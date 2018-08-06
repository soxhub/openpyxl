from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from io import BytesIO
from openpyxl.xml.functions import xmlfile
from openpyxl.xml.constants import SHEET_MAIN_NS


class WorksheetWriter:


    def __init__(self, ws, out=None):
        self.ws = ws
        if out is None:
            out = BytesIO()
        self.out = out
        #self._rels = RelationshipList()
        self._hyperlinks = []
        self.xf = self.write()
        next(self.xf) # start generator


    def write_properties(self):
        props = self.ws.sheet_properties
        self.xf.send(props.to_tree())


    def write_cols(self):
        cols = self.ws.column_dimensions
        self.xf.send(cols.to_tree())


    def write_views(self):
        pass


    def write_columns(self):
        pass


    def write_rows(self):
        pass


    def write_protection(self):
        pass


    def write_filter(self):
        pass


    def write_sort(self):
        pass


    def write_merged_cells(self):
        pass


    def write_formatting(self):
        pass


    def write_validations(self):
        pass


    def write_hyperlinks(self):
        pass


    def write_print(self):
        pass


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
