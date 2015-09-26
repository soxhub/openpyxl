from __future__ import absolute_import

from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.worksheet.page import (PageMargins, PrintPageSetup)
from openpyxl.worksheet.header_footer import HeaderFooter

from .relation import Drawing, DrawingHF, SheetBackgroundPicture
from .properties import ChartsheetProperties
from .protection import ChartsheetProtection
from .views import ChartsheetViews
from .custom import CustomChartsheetViews
from .publish import WebPublishItems


class Chartsheet(Serialisable):
    tagname = "chartsheet"

    sheetPr = Typed(expected_type=ChartsheetProperties, allow_none=True)
    sheetViews = Typed(expected_type=ChartsheetViews, )
    sheetProtection = Typed(expected_type=ChartsheetProtection, allow_none=True)
    customSheetViews = Typed(expected_type=CustomChartsheetViews, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)
    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    drawing = Typed(expected_type=Drawing, )
    drawingHF = Typed(expected_type=DrawingHF, allow_none=True)
    picture = Typed(expected_type=SheetBackgroundPicture, allow_none=True)
    webPublishItems = Typed(expected_type=WebPublishItems, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = (
        'sheetPr', 'sheetViews', 'sheetProtection', 'customSheetViews', 'pageMargins', 'pageSetup', 'headerFooter',
        'drawing', 'drawingHF', 'picture', 'webPublishItems')

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
                 ):
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
