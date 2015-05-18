from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.styles.colors import Color, BLACK, WHITE

from openpyxl.utils.units import (
    pixels_to_EMU,
    EMU_to_pixels,
    short_color,
)

from openpyxl.compat import deprecated


class Shape(object):
    """ a drawing inside a chart
        coordiantes are specified by the user in the axis units
    """

    MARGIN_LEFT = 6 + 13 + 1
    MARGIN_BOTTOM = 17 + 11

    FONT_WIDTH = 7
    FONT_HEIGHT = 8

    ROUND_RECT = 'roundRect'
    RECT = 'rect'

    # other shapes to define :
    '''
    "line"
    "lineInv"
    "triangle"
    "rtTriangle"
    "diamond"
    "parallelogram"
    "trapezoid"
    "nonIsoscelesTrapezoid"
    "pentagon"
    "hexagon"
    "heptagon"
    "octagon"
    "decagon"
    "dodecagon"
    "star4"
    "star5"
    "star6"
    "star7"
    "star8"
    "star10"
    "star12"
    "star16"
    "star24"
    "star32"
    "roundRect"
    "round1Rect"
    "round2SameRect"
    "round2DiagRect"
    "snipRoundRect"
    "snip1Rect"
    "snip2SameRect"
    "snip2DiagRect"
    "plaque"
    "ellipse"
    "teardrop"
    "homePlate"
    "chevron"
    "pieWedge"
    "pie"
    "blockArc"
    "donut"
    "noSmoking"
    "rightArrow"
    "leftArrow"
    "upArrow"
    "downArrow"
    "stripedRightArrow"
    "notchedRightArrow"
    "bentUpArrow"
    "leftRightArrow"
    "upDownArrow"
    "leftUpArrow"
    "leftRightUpArrow"
    "quadArrow"
    "leftArrowCallout"
    "rightArrowCallout"
    "upArrowCallout"
    "downArrowCallout"
    "leftRightArrowCallout"
    "upDownArrowCallout"
    "quadArrowCallout"
    "bentArrow"
    "uturnArrow"
    "circularArrow"
    "leftCircularArrow"
    "leftRightCircularArrow"
    "curvedRightArrow"
    "curvedLeftArrow"
    "curvedUpArrow"
    "curvedDownArrow"
    "swooshArrow"
    "cube"
    "can"
    "lightningBolt"
    "heart"
    "sun"
    "moon"
    "smileyFace"
    "irregularSeal1"
    "irregularSeal2"
    "foldedCorner"
    "bevel"
    "frame"
    "halfFrame"
    "corner"
    "diagStripe"
    "chord"
    "arc"
    "leftBracket"
    "rightBracket"
    "leftBrace"
    "rightBrace"
    "bracketPair"
    "bracePair"
    "straightConnector1"
    "bentConnector2"
    "bentConnector3"
    "bentConnector4"
    "bentConnector5"
    "curvedConnector2"
    "curvedConnector3"
    "curvedConnector4"
    "curvedConnector5"
    "callout1"
    "callout2"
    "callout3"
    "accentCallout1"
    "accentCallout2"
    "accentCallout3"
    "borderCallout1"
    "borderCallout2"
    "borderCallout3"
    "accentBorderCallout1"
    "accentBorderCallout2"
    "accentBorderCallout3"
    "wedgeRectCallout"
    "wedgeRoundRectCallout"
    "wedgeEllipseCallout"
    "cloudCallout"
    "cloud"
    "ribbon"
    "ribbon2"
    "ellipseRibbon"
    "ellipseRibbon2"
    "leftRightRibbon"
    "verticalScroll"
    "horizontalScroll"
    "wave"
    "doubleWave"
    "plus"
    "flowChartProcess"
    "flowChartDecision"
    "flowChartInputOutput"
    "flowChartPredefinedProcess"
    "flowChartInternalStorage"
    "flowChartDocument"
    "flowChartMultidocument"
    "flowChartTerminator"
    "flowChartPreparation"
    "flowChartManualInput"
    "flowChartManualOperation"
    "flowChartConnector"
    "flowChartPunchedCard"
    "flowChartPunchedTape"
    "flowChartSummingJunction"
    "flowChartOr"
    "flowChartCollate"
    "flowChartSort"
    "flowChartExtract"
    "flowChartMerge"
    "flowChartOfflineStorage"
    "flowChartOnlineStorage"
    "flowChartMagneticTape"
    "flowChartMagneticDisk"
    "flowChartMagneticDrum"
    "flowChartDisplay"
    "flowChartDelay"
    "flowChartAlternateProcess"
    "flowChartOffpageConnector"
    "actionButtonBlank"
    "actionButtonHome"
    "actionButtonHelp"
    "actionButtonInformation"
    "actionButtonForwardNext"
    "actionButtonBackPrevious"
    "actionButtonEnd"
    "actionButtonBeginning"
    "actionButtonReturn"
    "actionButtonDocument"
    "actionButtonSound"
    "actionButtonMovie"
    "gear6"
    "gear9"
    "funnel"
    "mathPlus"
    "mathMinus"
    "mathMultiply"
    "mathDivide"
    "mathEqual"
    "mathNotEqual"
    "cornerTabs"
    "squareTabs"
    "plaqueTabs"
    "chartX"
    "chartStar"
    "chartPlus"
    '''

    @deprecated("Chart Drawings need a complete rewrite")
    def __init__(self,
                 chart,
                 coordinates=((0, 0), (1, 1)),
                 text=None,
                 scheme="accent1"):
        self.chart = chart
        self.coordinates = coordinates  # in axis units
        self.text = text
        self.scheme = scheme
        self.style = Shape.RECT
        self.border_width = 0
        self.border_color = BLACK  # "F3B3C5"
        self.color = WHITE
        self.text_color = BLACK

    @property
    def border_color(self):
        return self._border_color

    @border_color.setter
    def border_color(self, color):
        self._border_color = short_color(color)

    @property
    def color(self):
        return self._color

    @color.setter
    def color(self, color):
        self._color = short_color(color)

    @property
    def text_color(self):
        return self._text_color

    @text_color.setter
    def text_color(self, color):
        self._text_color = short_color(color)

    @property
    def border_width(self):
        return self._border_width

    @border_width.setter
    def border_width(self, w):
        self._border_width = w

    @property
    def coordinates(self):
        """Return coordindates in axis units"""
        return self._coordinates

    @coordinates.setter
    def coordinates(self, coords):
        """ set shape coordinates in percentages (left, top, right, bottom)
        """
        # this needs refactoring to reflect changes in charts
        self.axis_coordinates = coords
        (x1, y1), (x2, y2) = coords # bottom left, top right
        drawing_width = pixels_to_EMU(self.chart.drawing.width)
        drawing_height = pixels_to_EMU(self.chart.drawing.height)
        plot_width = drawing_width * self.chart.width
        plot_height = drawing_height * self.chart.height

        margin_left = self.chart._get_margin_left() * drawing_width
        xunit = plot_width / self.chart.get_x_units()

        margin_top = self.chart._get_margin_top() * drawing_height
        yunit = self.chart.get_y_units()

        x_start = (margin_left + (float(x1) * xunit)) / drawing_width
        y_start = ((margin_top
                    + plot_height
                    - (float(y1) * yunit))
                    / drawing_height)

        x_end = (margin_left + (float(x2) * xunit)) / drawing_width
        y_end = ((margin_top
                  + plot_height
                  - (float(y2) * yunit))
                  / drawing_height)

        # allow user to specify y's in whatever order
        # excel expect y_end to be lower
        if y_end < y_start:
            y_end, y_start = y_start, y_end

        self._coordinates = (
            self._norm_pct(x_start), self._norm_pct(y_start),
            self._norm_pct(x_end), self._norm_pct(y_end)
        )

    @staticmethod
    def _norm_pct(pct):
        """ force shapes to appear by truncating too large sizes """
        if pct > 1:
            return 1
        elif pct < 0:
            return 0
        return pct
