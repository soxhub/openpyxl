from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

"""Reader for a single worksheet."""
from io import BytesIO
from warnings import warn

# compatibility imports
from openpyxl.xml.functions import iterparse

# package imports
from openpyxl.cell import Cell
from openpyxl.cell.read_only import _cast_number
from openpyxl.cell.text import Text
from openpyxl.worksheet import Worksheet
from openpyxl.worksheet.dimensions import (
    ColumnDimension,
    RowDimension,
    SheetFormatProperties,
)

from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
    EXT_TYPES,
)
from openpyxl.xml.functions import safe_iterator
from openpyxl.styles import Color
from openpyxl.styles import is_date_format
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.formatting import Rule
from openpyxl.formatting.formatting import ConditionalFormatting
from openpyxl.formula.translate import Translator
from openpyxl.utils import (
    get_column_letter,
    coordinate_to_tuple,
    )
from openpyxl.utils.datetime import from_excel, from_ISO8601, WINDOWS_EPOCH
from openpyxl.descriptors.excel import ExtensionList, Extension

from .filters import AutoFilter, SortState
from .header_footer import HeaderFooter
from .hyperlink import HyperlinkList
from .merge import MergeCells
from .cell_range import CellRange
from .page import PageMargins, PrintOptions, PrintPageSetup
from .pagebreak import PageBreak
from .protection import SheetProtection
from .scenario import ScenarioList
from .views import SheetViewList
from .datavalidation import DataValidationList
from .table import TablePartList
from .properties import WorksheetProperties
from .dimensions import SheetDimension
from .related import Related


CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
MERGE_TAG = '{%s}mergeCells' % SHEET_MAIN_NS
INLINE_STRING = "{%s}is" % SHEET_MAIN_NS
COL_TAG = '{%s}col' % SHEET_MAIN_NS
ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CF_TAG = '{%s}conditionalFormatting' % SHEET_MAIN_NS
LEGACY_TAG = '{%s}legacyDrawing' % SHEET_MAIN_NS
PROT_TAG = '{%s}sheetProtection' % SHEET_MAIN_NS
EXT_TAG = "{%s}extLst" % SHEET_MAIN_NS
HYPERLINK_TAG = "{%s}hyperlinks" % SHEET_MAIN_NS
TABLE_TAG = "{%s}tableParts" % SHEET_MAIN_NS
PRINT_TAG = '{%s}printOptions' % SHEET_MAIN_NS
MARGINS_TAG = '{%s}pageMargins' % SHEET_MAIN_NS
PAGE_TAG = '{%s}pageSetup' % SHEET_MAIN_NS
HEADER_TAG = '{%s}headerFooter' % SHEET_MAIN_NS
FILTER_TAG = '{%s}autoFilter' % SHEET_MAIN_NS
VALIDATION_TAG = '{%s}dataValidations' % SHEET_MAIN_NS
PROPERTIES_TAG = '{%s}sheetPr' % SHEET_MAIN_NS
VIEWS_TAG = '{%s}sheetViews' % SHEET_MAIN_NS
FORMAT_TAG = '{%s}sheetFormatPr' % SHEET_MAIN_NS
ROW_BREAK_TAG = '{%s}rowBreaks' % SHEET_MAIN_NS
SCENARIOS_TAG = '{%s}scenarios' % SHEET_MAIN_NS
DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS


class WorkSheetParser(object):

    def __init__(self, xml_source, shared_strings, data_only=False,
                 epoch=WINDOWS_EPOCH, styles={}):
        self.min_row = self.min_col = self.max_row = self.max_column = None
        self.epoch = epoch
        self.source = xml_source
        self.shared_strings = shared_strings
        self.data_only = data_only
        self.shared_formulae = {}
        self.array_formulae = {}
        self.max_row = self.max_column = 0
        self.tables = TablePartList()
        self._number_format_cache = {}
        self.row_dimensions = {}
        self.column_dimensions = {}
        self.styles = styles
        self.number_formats = []
        self.keep_vba = False
        self.hyperlinks = HyperlinkList()
        self.formatting = []
        self.legacy_drawing = None


    def _is_date(self, style_id):
        """
        Check whether a particular style has a date format
        """
        if style_id in self._number_format_cache:
            return self._number_format_cache[style_id]

        style = self.styles[style_id]
        key = style.numFmtId
        if key < 164:
            fmt = BUILTIN_FORMATS.get(key, "General")
        else:
            fmt = self.number_formats[key - 164]
        is_date = is_date_format(fmt)
        self._number_format_cache[style_id] = is_date
        return is_date

    def parse(self):
        dispatcher = {
            COL_TAG: self.parse_column_dimensions,
            PROT_TAG: self.parse_sheet_protection,
            EXT_TAG: self.parse_extensions,
            CF_TAG: self.parse_formatting,
                      }

        properties = {
            PRINT_TAG: ('print_options', PrintOptions),
            MARGINS_TAG: ('page_margins', PageMargins),
            PAGE_TAG: ('page_setup', PrintPageSetup),
            HEADER_TAG: ('HeaderFooter', HeaderFooter),
            FILTER_TAG: ('auto_filter', AutoFilter),
            VALIDATION_TAG: ('data_validations', DataValidationList),
            PROPERTIES_TAG: ('sheet_properties', WorksheetProperties),
            VIEWS_TAG: ('views', SheetViewList),
            FORMAT_TAG: ('sheet_format', SheetFormatProperties),
            ROW_BREAK_TAG: ('page_breaks', PageBreak),
            SCENARIOS_TAG: ('scenarios', ScenarioList),
            LEGACY_TAG: ('legacy_drawing', Related),
            TABLE_TAG: ('tables', TablePartList),
            HYPERLINK_TAG: ('hyperlinks', HyperlinkList),
            MERGE_TAG: ('merged_cells', MergeCells),
        }

        it = iterparse(self.source, tag=dispatcher)

        for _, element in it:
            tag_name = element.tag
            if tag_name in dispatcher:
                dispatcher[tag_name](element)
                element.clear()
            elif tag_name in properties:
                prop = properties[tag_name]
                obj = prop[1].from_tree(element)
                setattr(self, prop[0], obj)
                element.clear()
            elif tag_name == ROW_TAG:
                yield self.parse_row(element)


    def parse_dimensions(self):
        """
        Get worksheet dimensions if they are provided.
        """
        it = iterparse(self.source)

        for _event, element in it:
            if element.tag == DIMENSION_TAG:
                dim = SheetDimension.from_tree(element)
                return dim.boundaries

            elif element.tag == DATA_TAG:
                # Dimensions missing
                break
            element.clear()


    def parse_cell(self, element):
        value = element.findtext(VALUE_TAG)
        data_type = element.get('t', 'n')
        coordinate = element.get('r')
        self.max_column += 1
        style_id = element.get('s', 0)

        if style_id is not None:
            style_id = int(style_id)

        if coordinate:
            row, column = coordinate_to_tuple(coordinate)
        else:
            row, column = self.max_row, self.max_column

        if data_type == "str" and not self.data_only:
            data_type = 'f'
            value = self.parse_formula(element)

        if value is not None:
            if data_type == 'n':
                value = _cast_number(value)
                if self._is_date(style_id):
                    data_type = 'd'
                    value = from_excel(value, self.epoch)
            elif data_type == 's':
                value = self.shared_strings[int(value)]
            elif data_type == 'b':
                value = bool(int(value))
            elif data_type == "str":
                try:
                    value = _cast_number(value)
                    data_type = "n"
                except ValueError:
                    data_type = "s"
            elif data_type == 'd':
                value = from_ISO8601(value)

        elif data_type == 'inlineStr':
                child = element.find(INLINE_STRING)
                if child is not None:
                    data_type = 's'
                    richtext = Text.from_tree(child)
                    value = richtext.content

        return {'row':row, 'column':column, 'value':value, 'data_type':data_type, 'style_id':style_id}


    def parse_formula(self, element):
        """
        possible formulae types: shared, array, datatable
        """
        formula = element.find(FORMULA_TAG)
        data_type = 'f'
        formula_type = formula.get('t')
        coordinate = element.get('r')
        value = "="
        if formula.text is not None:
            value += formula.text

        if formula_type == "array":
            self.array_formulae[coordinate] = dict(formula.attrib)

        elif formula_type == "shared":
            idx = formula.get('si')
            if idx in self.shared_formulae:
                trans = self.shared_formulae[idx]
                value = trans.translate_formula(coordinate)
            elif value != "=":
                self.shared_formulae[idx] = Translator(value, coordinate)
        return value


    def parse_column_dimensions(self, col):
        attrs = dict(col.attrib)
        attrs['worksheet'] = None
        column = get_column_letter(int(attrs['min']))
        attrs['index'] = column
        if 'style' in attrs:
            attrs['style'] = self.styles[int(attrs['style'])]
        dim = ColumnDimension(**attrs)
        self.column_dimensions[column] = dim


    def parse_row(self, row):
        attrs = dict(row.attrib)

        if "r" in attrs:
            self.max_row = int(attrs['r'])
        else:
            self.max_row += 1
        self.max_column = 0
        keys = set(attrs)
        for key in keys:
            if key == "s":
                attrs['s'] = self.styles[int(attrs['s'])]
            elif key.startswith('{'):
                del attrs[key]

        keys = set(attrs)
        if keys != set(['r', 'spans']) and keys != set(['r']):
            # don't create dimension objects unless they have relevant information
            attrs['worksheet'] = None
            dim = RowDimension(**attrs)
            self.row_dimensions[dim.index] = dim

        cells = []
        for cell in safe_iterator(row, CELL_TAG):
            cells.append(self.parse_cell(cell))
        return cells


    def parse_formatting(self, element):
        cf = ConditionalFormatting.from_tree(element)
        self.formatting.append(cf)


    def parse_sheet_protection(self, element):
        protection = SheetProtection.from_tree(element)
        password = element.get("password")
        if password is not None:
            protection.set_password(password, True)
        self.protection = protection


    def parse_extensions(self, element):
        extLst = ExtensionList.from_tree(element)
        for e in extLst.ext:
            ext_type = EXT_TYPES.get(e.uri.upper(), "Unknown")
            msg = "{0} extension is not supported and will be removed".format(ext_type)
            warn(msg)


class WorksheetReader(object):
    """
    Create a parser and apply it to a workbook
    """

    def __init__(self, ws, xml_source, shared_strings, data_only):
        self.ws = ws
        self.parser = WorkSheetParser(xml_source, shared_strings, data_only, ws.parent.epoch, ws.parent._cell_styles)
        self.tables = []


    def bind_cells(self):
        for row in self.parser.parse():
            for cell in row:
                style = self.ws.parent._cell_styles[cell['style_id']]
                c = Cell(self.ws, row=cell['row'], column=cell['column'], style_array=style)
                c._value = cell['value']
                c.data_type = cell['data_type']
                self.ws._cells[(cell['row'], cell['column'])] = c


    def bind_formatting(self):
        for cf in self.parser.formatting:
            for rule in cf.rules:
                if rule.dxfId is not None:
                    rule.dxf = self.ws.parent._differential_styles[rule.dxfId]
            self.ws.conditional_formatting[cf] = rule


    def bind_tables(self):
        for t in self.parser.tables.tablePart:
            rel = self.ws._rels[t.id]
            self.tables.append(rel.Target)


    def bind_merged_cells(self):
        self.ws.merged_cells.ranges = self.parser.merged_cells.mergeCell
        for cr in self.parser.merged_cells.mergeCell:
            self.ws._clean_merge_range(cr)


    def bind_hyperlinks(self):
        for link in self.parser.hyperlinks.hyperlink:
            if link.id:
                rel = self.ws._rels[link.id]
                link.target = rel.Target
            if ":" in link.ref:
                # range of cells
                for row in self.ws[link.ref]:
                    for cell in row:
                        cell.hyperlink = link
            else:
                self.ws[link.ref].hyperlink = link


    def bind_col_dimensions(self):
        self.ws.column_dimensions = self.parser.column_dimensions
        for cd in self.ws.column_dimensions.values():
            cd.worksheet = self.ws


    def bind_row_dimensions(self):
        self.ws.row_dimensions = self.parser.row_dimensions
        for rd in self.ws.row_dimensions.values():
            rd.worksheet = self.ws


    def bind_properties(self):
        for k in ('print_options', 'page_margins', 'page_setup',
                  'HeaderFooter', 'auto_filter', 'data_validations',
                  'sheet_properties', 'views', 'sheet_format',
                  'page_breaks', 'scenarios', 'legacy_drawing'):
            v = getattr(self.parser, k, None)
            if v is not None:
                setattr(self.ws, k, v)


    def bind_all(self):
        self.bind_cells()
        self.bind_merged_cells()
        self.bind_hyperlinks()
        self.bind_formatting()
        self.bind_col_dimensions()
        self.bind_row_dimensions()
        self.bind_tables()
        self.bind_properties()
