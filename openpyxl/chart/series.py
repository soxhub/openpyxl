from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedInteger

from .shapes import ShapeProperties, Shape
from .chartBase import AxDataSource, NumDataSource, NumRef
from .error_bar import ErrorBars
from .label import DataLabels
from .marker import DataPoint, PictureOptions, Marker
from .trendline import Trendline

attribute_mapping = {'area': ('idx', 'order', 'tx', 'spPr', 'pictureOptions', 'dPt', 'dLbls', 'errBars',
 'trendline', 'cat', 'val',),
                     'bar':('idx', 'order','tx', 'spPr', 'invertIfNegative', 'pictureOptions', 'dPt',
 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'shape'),
                     'bubble':('idx','order', 'tx', 'spPr', 'invertIfNegative', 'dPt', 'dLbls',
 'trendline', 'errBars', 'xVal', 'yVal', 'bubbleSize', 'bubble3D'),
                     'line':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline',
 'errBars', 'cat', 'val', 'smooth'),
                     'pie':('idx', 'order', 'tx', 'spPr', 'explosion', 'dPt', 'dLbls', 'cat', 'val'),
                     'radar':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'cat', 'val'),
                     'scatter':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline',
 'errBars', 'xVal', 'yVal', 'smooth'),
                     'surface':('idx', 'order', 'tx', 'spPr', 'cat', 'val'),
                     }


def Series(name_ref=None, cat_ref=None, values=None, order=None):
    """
    High level function for creating series

    See http://exceluser.com/excel_help/functions/function-series.htm for a description
    """
    series = _SeriesBase()
    series.val = NumDataSource(numRef=NumRef(f=values))
    return series


class StrVal(Serialisable):

    idx = Integer()
    v = Typed(expected_type=String(), )

    def __init__(self,
                 idx=None,
                 v=None,
                ):
        self.idx = idx
        self.v = v


class StrData(Serialisable):

    ptCount = Integer(allow_none=True, nested=True)
    pt = Typed(expected_type=StrVal, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'pt')

    def __init__(self,
                 ptCount=None,
                 pt=None,
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.pt = pt


class StrRef(Serialisable):

    f = Typed(expected_type=String, )
    strCache = Typed(expected_type=StrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'strCache')

    def __init__(self,
                 f=None,
                 strCache=None,
                 extLst=None,
                ):
        self.f = f
        self.strCache = strCache


class SerTx(Serialisable):

    strRef = Typed(expected_type=StrRef)


class _SeriesBase(Serialisable):

    tagname = "ser"

    idx = NestedInteger()
    order = NestedInteger()
    tx = Typed(expected_type=SerTx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    ShapeProperties = Alias('spPr')

    # area chart
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    #bar chart
    invertIfNegative = Bool(nested=True, allow_none=True)
    shape = Typed(expected_type=Shape, allow_none=True)

    #bubble chart
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    bubbleSize = Typed(expected_type=NumDataSource, allow_none=True)
    bubble3D = Bool(nested=True, allow_none=True)

    #line chart
    marker = Typed(expected_type=Marker, allow_none=True)
    smooth = Bool(allow_none=True, nested=True)

    #pie chart
    explosion = Integer(allow_none=True, nested=True)

    __elements__ = ('bubble3D', 'bubbleSize', 'cat', 'dLbls', 'dPt', 'errBars',
                 'explosion', 'idx', 'invertIfNegative', 'marker', 'order',
                 'pictureOptions', 'shape', 'smooth', 'spPr', 'trendline', 'tx', 'val',
                 'xVal', 'yVal')


    def __init__(self,
                 idx=0,
                 order=0,
                 tx=None,
                 spPr=None,
                 pictureOptions=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 invertIfNegative=None,
                 shape=None,
                 xVal=None,
                 yVal=None,
                 bubbleSize=None,
                 bubble3D=None,
                 marker=None,
                 smooth=None,
                 explosion=None
                ):
        self.idx = idx
        self.order = order
        self.tx = tx
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.invertIfNegative = invertIfNegative
        self.shape = shape
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D
        self.marker = marker
        self.smooth = smooth
        self.explosion = explosion

    def to_tree(self, tagname=None, idx=None):
        if idx is not None:
            if self.order == self.idx:
                self.order = idx
            self.idx = idx
        return super(_SeriesBase, self).to_tree(tagname)


class AreaSer(_SeriesBase):

    pictureOptions = _SeriesBase.pictureOptions
    dPt = _SeriesBase.dPt
    dLbls = _SeriesBase.dLbls
    trendline = _SeriesBase.trendline
    errBars = _SeriesBase.errBars
    cat = _SeriesBase.cat
    val = _SeriesBase.val
    extLst = _SeriesBase.extLst

    __elements__ = attribute_mapping['area']

    def __init__(self,
                 pictureOptions=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.pictureOptions = pictureOptions
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val


class BarSer(_SeriesBase):

    idx = _SeriesBase.idx
    order = _SeriesBase.order
    tx = _SeriesBase.tx
    spPr = _SeriesBase.spPr

    invertIfNegative = _SeriesBase.invertIfNegative
    pictureOptions = _SeriesBase.pictureOptions
    dPt = _SeriesBase.dPt
    dLbls = _SeriesBase.dLbls
    trendline = _SeriesBase.dLbls
    errBars = _SeriesBase.errBars
    cat = _SeriesBase.cat
    val = _SeriesBase.val
    shape = _SeriesBase.shape
    extLst = _SeriesBase.extLst


    def __init__(self, **kw):
        self.__elements__ = attribute_mapping['bar']
        super(BarSer, self).__init__(**kw)


class BubbleSer(_SeriesBase):

    invertIfNegative = Bool(nested=True, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    bubbleSize = Typed(expected_type=NumDataSource, allow_none=True)
    bubble3D = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['bubble']


    def __init__(self,
                 invertIfNegative=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 xVal=None,
                 yVal=None,
                 bubbleSize=None,
                 bubble3D=None,
                 extLst=None,
                ):
        self.invertIfNegative = invertIfNegative
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D


class PieSer(_SeriesBase):

    explosion = Integer(allow_none=True, nested=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['pie']

    def __init__(self,
                 explosion=None,
                 dPt=None,
                 dLbls=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.explosion = explosion
        self.dPt = dPt
        self.dLbls = dLbls
        self.cat = cat
        self.val = val


class RadarSer(_SeriesBase):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['pie']

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.cat = cat
        self.val = val


class ScatterSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['scatter']

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 xVal=None,
                 yVal=None,
                 smooth=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.xVal = xVal
        self.yVal = yVal
        self.smooth = smooth


class SurfaceSer(Serialisable):

    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['surface']

    def __init__(self,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.cat = cat
        self.val = val


class LineSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Bool(allow_none=True, nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = attribute_mapping['line']

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 smooth=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.smooth = smooth
