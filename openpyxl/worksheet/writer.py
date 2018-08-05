from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl


from io import BytesIO


class WorksheetWriter:


    def __init__(self, out=None):
        if out is None:
            out = BytesIO()
        self.out = out
        self._rels = RelationshipList()
        self._hyperlinks = []


    def write_properties(self):
        pass


    def write_dimensions(self):
        pass


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
        pass
