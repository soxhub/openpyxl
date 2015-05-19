from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.xml.functions import Element, SubElement, tostring
from openpyxl.xml.constants import (
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_NS,
    CHART_DRAWING_NS,
    PKG_REL_NS
)
from openpyxl.compat.strings import safe_string
from openpyxl.chart.spreadsheet_drawing import (
    SpreadsheetDrawing,
)
from openpyxl.packaging.relationship import Relationship


class DrawingWriter(object):
    """ one main drawing file per sheet """

    def __init__(self, sheet):
        self._sheet = sheet

    def write(self):
        """ write drawings for one sheet in one file """

        root = Element("wsDr", xmlns=SHEET_DRAWING_NS)

        for idx, chart in enumerate(self._sheet._charts, 1):
            node = self._write_chart(chart, idx)
            self.root.append(node)

        for idx, img in enumerate(self._sheet._images, 1):
            anchor = self._write_image(img)
            self.root.append(node)

        return tostring(root)


    def _write_chart(self, chart, idx):
        """Add a chart"""
        drawing = chart.drawing
        anchor = drawing.anchor
        _drawing = SpreadsheetDrawing()
        anchor.graphicFrame = _drawing._chart_frame(idx)

        return anchor


    def _write_image(self, img, idx):
        """Add an image"""
        anchor = img.drawing.anchor
        _drawing = SpreadsheetDrawing()
        anchor.pic = _drawing._picture_frame(idx)

        return anchor

    def write_rels(self, chart_id, image_id):

        root = Element("Relationships", xmlns=PKG_REL_NS)
        i = 0
        for i, chart in enumerate(self._sheet._charts):
            rel = Relationship(type="chart", target='../charts/chart%s.xml' %
                               (chart_id + i), id='rId%s' % (i + 1))
            root.append(rel.to_tree())
        for j, img in enumerate(self._sheet._images):
            rel = Relationship(type="image", target='../media/image%s.png' %
                       (image_id + i), id='rId%s' % (i + j + 1))
            root.append(rel.to_tree())
        return tostring(root)


class ShapeWriter(object):
    """ one file per shape """

    def __init__(self, shapes):

        self._shapes = shapes

    def write(self, shape_id):

        root = Element('{%s}userShapes' % CHART_NS)

        for shape in self._shapes:
            anchor = SubElement(root, '{%s}relSizeAnchor' % CHART_DRAWING_NS)

            xstart, ystart, xend, yend = shape.coordinates

            _from = SubElement(anchor, '{%s}from' % CHART_DRAWING_NS)
            SubElement(_from, '{%s}x' % CHART_DRAWING_NS).text = str(xstart)
            SubElement(_from, '{%s}y' % CHART_DRAWING_NS).text = str(ystart)

            _to = SubElement(anchor, '{%s}to' % CHART_DRAWING_NS)
            SubElement(_to, '{%s}x' % CHART_DRAWING_NS).text = str(xend)
            SubElement(_to, '{%s}y' % CHART_DRAWING_NS).text = str(yend)

            sp = SubElement(anchor, '{%s}sp' % CHART_DRAWING_NS, {'macro':'', 'textlink':''})
            nvspr = SubElement(sp, '{%s}nvSpPr' % CHART_DRAWING_NS)
            SubElement(nvspr, '{%s}cNvPr' % CHART_DRAWING_NS, {'id':str(shape_id), 'name':'shape %s' % shape_id})
            SubElement(nvspr, '{%s}cNvSpPr' % CHART_DRAWING_NS)

            sppr = SubElement(sp, '{%s}spPr' % CHART_DRAWING_NS)
            frm = SubElement(sppr, '{%s}xfrm' % DRAWING_NS,)
            # no transformation
            SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
            SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

            prstgeom = SubElement(sppr, '{%s}prstGeom' % DRAWING_NS, {'prst':str(shape.style)})
            SubElement(prstgeom, '{%s}avLst' % DRAWING_NS)

            fill = SubElement(sppr, '{%s}solidFill' % DRAWING_NS, )
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.color})

            border = SubElement(sppr, '{%s}ln' % DRAWING_NS, {'w':str(shape._border_width)})
            sf = SubElement(border, '{%s}solidFill' % DRAWING_NS)
            SubElement(sf, '{%s}srgbClr' % DRAWING_NS, {'val':shape.border_color})

            self._write_style(sp)
            self._write_text(sp, shape)

            shape_id += 1

        return tostring(root)

    def _write_text(self, node, shape):
        """ write text in the shape """

        tx_body = SubElement(node, '{%s}txBody' % CHART_DRAWING_NS)
        SubElement(tx_body, '{%s}bodyPr' % DRAWING_NS, {'vertOverflow':'clip'})
        SubElement(tx_body, '{%s}lstStyle' % DRAWING_NS)
        p = SubElement(tx_body, '{%s}p' % DRAWING_NS)
        if shape.text:
            r = SubElement(p, '{%s}r' % DRAWING_NS)
            rpr = SubElement(r, '{%s}rPr' % DRAWING_NS, {'lang':'en-US'})
            fill = SubElement(rpr, '{%s}solidFill' % DRAWING_NS)
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.text_color})

            SubElement(r, '{%s}t' % DRAWING_NS).text = shape.text
        else:
            SubElement(p, '{%s}endParaRPr' % DRAWING_NS, {'lang':'en-US'})

    def _write_style(self, node):
        """ write style theme """

        style = SubElement(node, '{%s}style' % CHART_DRAWING_NS)

        ln_ref = SubElement(style, '{%s}lnRef' % DRAWING_NS, {'idx':'2'})
        scheme_clr = SubElement(ln_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})
        SubElement(scheme_clr, '{%s}shade' % DRAWING_NS, {'val':'50000'})

        fill_ref = SubElement(style, '{%s}fillRef' % DRAWING_NS, {'idx':'1'})
        SubElement(fill_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        effect_ref = SubElement(style, '{%s}effectRef' % DRAWING_NS, {'idx':'0'})
        SubElement(effect_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        font_ref = SubElement(style, '{%s}fontRef' % DRAWING_NS, {'idx':'minor'})
        SubElement(font_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'lt1'})
