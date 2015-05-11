from __future__ import absolute_import

from openpyxl.compat import basestring, unicode

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Integer,
    Set,
    MinMax,
)
from openpyxl.descriptors.excel import Percentage
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedValue,
)

from openpyxl.styles.colors import RGB
from openpyxl.xml.constants import DRAWING_NS

from .drawing import OfficeArtExtensionList

PRESET_COLORS = [
        'aliceBlue', 'antiqueWhite', 'aqua', 'aquamarine',
        'azure', 'beige', 'bisque', 'black', 'blanchedAlmond', 'blue',
        'blueViolet', 'brown', 'burlyWood', 'cadetBlue', 'chartreuse',
        'chocolate', 'coral', 'cornflowerBlue', 'cornsilk', 'crimson', 'cyan',
        'darkBlue', 'darkCyan', 'darkGoldenrod', 'darkGray', 'darkGrey',
        'darkGreen', 'darkKhaki', 'darkMagenta', 'darkOliveGreen', 'darkOrange',
        'darkOrchid', 'darkRed', 'darkSalmon', 'darkSeaGreen', 'darkSlateBlue',
        'darkSlateGray', 'darkSlateGrey', 'darkTurquoise', 'darkViolet',
        'dkBlue', 'dkCyan', 'dkGoldenrod', 'dkGray', 'dkGrey', 'dkGreen',
        'dkKhaki', 'dkMagenta', 'dkOliveGreen', 'dkOrange', 'dkOrchid', 'dkRed',
        'dkSalmon', 'dkSeaGreen', 'dkSlateBlue', 'dkSlateGray', 'dkSlateGrey',
        'dkTurquoise', 'dkViolet', 'deepPink', 'deepSkyBlue', 'dimGray',
        'dimGrey', 'dodgerBlue', 'firebrick', 'floralWhite', 'forestGreen',
        'fuchsia', 'gainsboro', 'ghostWhite', 'gold', 'goldenrod', 'gray',
        'grey', 'green', 'greenYellow', 'honeydew', 'hotPink', 'indianRed',
        'indigo', 'ivory', 'khaki', 'lavender', 'lavenderBlush', 'lawnGreen',
        'lemonChiffon', 'lightBlue', 'lightCoral', 'lightCyan',
        'lightGoldenrodYellow', 'lightGray', 'lightGrey', 'lightGreen',
        'lightPink', 'lightSalmon', 'lightSeaGreen', 'lightSkyBlue',
        'lightSlateGray', 'lightSlateGrey', 'lightSteelBlue', 'lightYellow',
        'ltBlue', 'ltCoral', 'ltCyan', 'ltGoldenrodYellow', 'ltGray', 'ltGrey',
        'ltGreen', 'ltPink', 'ltSalmon', 'ltSeaGreen', 'ltSkyBlue',
        'ltSlateGray', 'ltSlateGrey', 'ltSteelBlue', 'ltYellow', 'lime',
        'limeGreen', 'linen', 'magenta', 'maroon', 'medAquamarine', 'medBlue',
        'medOrchid', 'medPurple', 'medSeaGreen', 'medSlateBlue',
        'medSpringGreen', 'medTurquoise', 'medVioletRed', 'mediumAquamarine',
        'mediumBlue', 'mediumOrchid', 'mediumPurple', 'mediumSeaGreen',
        'mediumSlateBlue', 'mediumSpringGreen', 'mediumTurquoise',
        'mediumVioletRed', 'midnightBlue', 'mintCream', 'mistyRose', 'moccasin',
        'navajoWhite', 'navy', 'oldLace', 'olive', 'oliveDrab', 'orange',
        'orangeRed', 'orchid', 'paleGoldenrod', 'paleGreen', 'paleTurquoise',
        'paleVioletRed', 'papayaWhip', 'peachPuff', 'peru', 'pink', 'plum',
        'powderBlue', 'purple', 'red', 'rosyBrown', 'royalBlue', 'saddleBrown',
        'salmon', 'sandyBrown', 'seaGreen', 'seaShell', 'sienna', 'silver',
        'skyBlue', 'slateBlue', 'slateGray', 'slateGrey', 'snow', 'springGreen',
        'steelBlue', 'tan', 'teal', 'thistle', 'tomato', 'turquoise', 'violet',
        'wheat', 'white', 'whiteSmoke', 'yellow', 'yellowGreen'
    ]


SCHEME_COLORS= ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3',
                'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr', 'dk1', 'lt1',
                'dk2', 'lt2'
                ]


class SystemColor(Serialisable):

    val = Set(values=(['scrollBar', 'background', 'activeCaption',
                       'inactiveCaption', 'menu', 'window', 'windowFrame', 'menuText',
                       'windowText', 'captionText', 'activeBorder', 'inactiveBorder',
                       'appWorkspace', 'highlight', 'highlightText', 'btnFace', 'btnShadow',
                       'grayText', 'btnText', 'inactiveCaptionText', 'btnHighlight',
                       '3dDkShadow', '3dLight', 'infoText', 'infoBk', 'hotLight',
                       'gradientActiveCaption', 'gradientInactiveCaption', 'menuHighlight',
                       'menuBar']))
    lastClr = Typed(expected_type=RGB, allow_none=True)

    def __init__(self,
                 val=None,
                 lastClr=None,
                ):
        self.val = val
        self.lastClr = lastClr


class HslColor(Serialisable):

    hue = Integer()
    sat = MinMax(min=0, max=100)
    lum = MinMax(min=0, max=100)

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum



class RGBPercent(Serialisable):

    r = MinMax(min=0, max=100)
    g = MinMax(min=0, max=100)
    b = MinMax(min=0, max=100)

    def __init__(self,
                 r=None,
                 g=None,
                 b=None,
                ):
        self.r = r
        self.g = g
        self.b = b


class ColorChoice(Serialisable):

    tagname = "colorChoice"
    namespace = DRAWING_NS

    scrgbClr = Typed(expected_type=RGBPercent, allow_none=True)
    RGBPercent = Alias('scrgbClr')
    srgbClr = NestedValue(expected_type=unicode, allow_none=True)
    RGB = Alias('srgbClr')
    hslClr = Typed(expected_type=HslColor, allow_none=True)
    sysClr = Typed(expected_type=SystemColor, allow_none=True)
    schemeClr = NestedNoneSet(values=SCHEME_COLORS)
    prstClr = NestedNoneSet(values=PRESET_COLORS)

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 scrgbClr=None,
                 srgbClr=None,
                 hslClr=None,
                 sysClr=None,
                 schemeClr=None,
                 prstClr=None,
                ):
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr

_COLOR_SET = ('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
               'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')


class ColorMapping(Serialisable):

    bg1 = Set(values=_COLOR_SET)
    tx1 = Set(values=_COLOR_SET)
    bg2 = Set(values=_COLOR_SET)
    tx2 = Set(values=_COLOR_SET)
    accent1 = Set(values=_COLOR_SET)
    accent2 = Set(values=_COLOR_SET)
    accent3 = Set(values=_COLOR_SET)
    accent4 = Set(values=_COLOR_SET)
    accent5 = Set(values=_COLOR_SET)
    accent6 = Set(values=_COLOR_SET)
    hlink = Set(values=_COLOR_SET)
    folHlink = Set(values=_COLOR_SET)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bg1=None,
                 tx1=None,
                 bg2=None,
                 tx2=None,
                 accent1=None,
                 accent2=None,
                 accent3=None,
                 accent4=None,
                 accent5=None,
                 accent6=None,
                 hlink=None,
                 folHlink=None,
                 extLst=None,
                ):
        self.bg1 = bg1
        self.tx1 = tx1
        self.bg2 = bg2
        self.tx2 = tx2
        self.accent1 = accent1
        self.accent2 = accent2
        self.accent3 = accent3
        self.accent4 = accent4
        self.accent5 = accent5
        self.accent6 = accent6
        self.hlink = hlink
        self.folHlink = folHlink
        self.extLst = extLst


class ColorChoiceDescriptor(Typed):
    """
    Objects can choose from 7 different kinds of color system.
    Assume RGBHex if a string is passed in.
    """

    expected_type = ColorChoice
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, basestring):
            value = ColorChoice(srgbClr=value)
        super(ColorChoiceDescriptor, self).__set__(instance, value)
