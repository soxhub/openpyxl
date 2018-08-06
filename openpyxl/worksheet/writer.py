from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from io import BytesIO
from openpyxl.xml.functions import xmlfile
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.styles.differential import DifferentialStyle
from .dimensions import SheetDimension
from .hyperlink import HyperlinkList
from .merge import MergeCell, MergeCells


class WorksheetWriter:


    def __init__(self, ws, out=None):
        self.ws = ws
        if out is None:
            out = BytesIO()
        self.out = out
        self._rels = RelationshipList()
        self._hyperlinks = []
        self.xf = self.write()
        next(self.xf) # start generator


    def write_properties(self):
        props = self.ws.sheet_properties
        self.xf.send(props.to_tree())


    def write_dimensions(self):
        """
        Write worksheet size if known
        """
        ref = getattr(self.ws, 'calculate_dimension')
        if ref:
            dim = SheetDimension(ref())
            self.xf.send(dim.to_tree())


    def write_format(self):
        self.ws.sheet_format.outlineLevelCol = self.ws.column_dimensions.max_outline
        fmt = self.ws.sheet_format
        self.xf.send(fmt.to_tree())


    def write_views(self):
        views = self.ws.views
        self.xf.send(views.to_tree())


    def write_cols(self):
        cols = self.ws.column_dimensions
        self.xf.send(cols.to_tree())


    def write_top(self):
        """
        Write all elements up to rows:
        properties
        dimensions
        views
        format
        cols
        """
        self.write_properties()
        self.write_dimensions()
        self.write_format()
        self.write_views()
        self.write_cols()


    def write_rows(self):
        pass


    def write_protection(self):
        prot = self.ws.protection
        if prot:
            self.xf.send(prot.to_tree())


    def write_filter(self):
        flt = self.ws.auto_filter
        if flt:
            self.xf.send(flt.to_tree())


    def write_sort(self):
        """
        As per discusion with the OOXML Working Group global sort state is not required.
        openpyxl never reads it from existing files
        """
        pass


    def write_merged_cells(self):
        merged = self.ws.merged_cells
        if merged:
            cells = [MergeCell(str(ref)) for ref in self.ws.merged_cells]
            self.xf.send(MergeCells(mergeCell=cells).to_tree())


    def write_formatting(self):
        df = DifferentialStyle()
        wb = self.ws.parent
        for cf in self.ws.conditional_formatting:
            for rule in cf.rules:
                if rule.dxf and rule.dxf != df:
                    rule.dxfId = wb._differential_styles.add(rule.dxf)
            self.xf.send(cf.to_tree())


    def write_validations(self):
        dv = self.ws.data_validations
        if dv:
            self.xf.send(dv.to_tree())


    def write_hyperlinks(self):
        links = HyperlinkList()

        for link in self._hyperlinks:
            if link.target:
                rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
                self._rels.append(rel)
                link.id = rel.id
            links.hyperlink.append(link)

        if links:
            self.xf.send(links.to_tree())


    def write_print(self):
        print_options = self.ws.print_options
        if print_options:
            self.xf.send(print_options.to_tree())


    def write_margin(self):
        pass


    def write_page(self):
        pass


    def write_header(self):
        pass


    def write_breaks(self):
        pass


    def write_drawings(self):
        pass


    def write_tables(self):
        pass


    def write(self):
        with xmlfile(self.out) as xf:
            with xf.element("worksheet", xmlns=SHEET_MAIN_NS):
                try:
                    while True:
                        el = (yield)
                        xf.write(el)
                except GeneratorExit:
                    pass


    def write_tail(self):
        """
        Write all elements after the rows
        calc properties
        protection
        protected ranges #
        scenarios #
        filters
        sorts # always ignored
        data consolidation #
        custom views #
        merged cells
        phonetic properties #
        conditional formatting
        data validation
        hyperlinks
        print options
        page margins
        page setup
        header
        row breaks
        col breaks
        custom properties #
        cell watches #
        ignored errors #
        smart tags #
        drawing
        drawingHF #
        background #
        OLE objects #
        controls #
        web publishing #
        tables
        """
        pass
