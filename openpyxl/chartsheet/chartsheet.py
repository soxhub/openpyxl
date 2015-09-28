from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.worksheet.page import (
    PageMargins,
    PrintPageSetup
)
from openpyxl.packaging.relationship import Relationship
from openpyxl.worksheet.drawing import Drawing
from openpyxl.worksheet.header_footer import HeaderFooter
from openpyxl.workbook.child import _WorkbookChild

from .relation import DrawingHF, SheetBackgroundPicture
from .properties import ChartsheetProperties
from .protection import ChartsheetProtection
from .views import ChartsheetViews
from .custom import CustomChartsheetViews
from .publish import WebPublishItems


class Chartsheet(_WorkbookChild, Serialisable):

    tagname = "chartsheet"
    _default_title = "Chart"

    sheetPr = Typed(expected_type=ChartsheetProperties, allow_none=True)
    sheetViews = Typed(expected_type=ChartsheetViews, allow_none=True)
    sheetProtection = Typed(expected_type=ChartsheetProtection, allow_none=True)
    customSheetViews = Typed(expected_type=CustomChartsheetViews, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)
    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    drawing = Typed(expected_type=Drawing, allow_none=True)
    drawingHF = Typed(expected_type=DrawingHF, allow_none=True)
    picture = Typed(expected_type=SheetBackgroundPicture, allow_none=True)
    webPublishItems = Typed(expected_type=WebPublishItems, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = (
        'sheetPr', 'sheetViews', 'sheetProtection', 'customSheetViews',
        'pageMargins', 'pageSetup', 'headerFooter', 'drawing', 'drawingHF',
        'picture', 'webPublishItems')

    def __init__(self,
                 sheetPr=None,
                 sheetViews=None,
                 sheetProtection=None,
                 customSheetViews=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 drawing=None,
                 drawingHF=None,
                 picture=None,
                 webPublishItems=None,
                 extLst=None,
                 parent=None,
                 title="",
                 ):
        # hack to simplify testing
        if parent is not None:
            super(Chartsheet, self).__init__(parent, title)
        self._charts = []
        self.sheetPr = sheetPr
        self.sheetViews = sheetViews
        self.sheetProtection = sheetProtection
        self.customSheetViews = customSheetViews
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.headerFooter = headerFooter
        self.drawing = drawing
        self.drawingHF = drawingHF
        self.picture = picture
        self.webPublishItems = webPublishItems


    def add_chart(self, chart):
        self._charts.append(chart)
        self.parent._charts.append(chart)


    def to_tree(self):
        self._rels = []
        if self._charts:
            rel = Relationship(type="drawing", target="")
            self._rels.append(rel)
            self.drawing = Drawing()
            self.drawing.id = "rId%s" % len(self._rels)
        return super(Chartsheet, self).to_tree()
