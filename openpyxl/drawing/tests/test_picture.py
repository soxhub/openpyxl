# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def PictureLocking():
    from ..picture import PictureLocking
    return PictureLocking

class TestPictureLocking:

    def test_ctor(self, PictureLocking):
        graphic = PictureLocking(noChangeAspect=True)
        xml = tostring(graphic.to_tree())
        expected = """
        <picLocks xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureLocking):
        src = """
        <picLocks noRot="1" />
        """
        node = fromstring(src)
        graphic = PictureLocking.from_tree(node)
        assert graphic == PictureLocking(noRot=1)


@pytest.fixture
def NonVisualPictureProperties():
    from ..picture import NonVisualPictureProperties
    return NonVisualPictureProperties


class TestNonVisualPictureProperties:

    def test_ctor(self, NonVisualPictureProperties):
        graphic = NonVisualPictureProperties()
        xml = tostring(graphic.to_tree())
        expected = """
        <cNvPicPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualPictureProperties):
        src = """
        <cNvPicPr />
        """
        node = fromstring(src)
        graphic = NonVisualPictureProperties.from_tree(node)
        assert graphic == NonVisualPictureProperties()



@pytest.fixture
def PictureNonVisual():
    from ..picture import PictureNonVisual
    return PictureNonVisual


class TestPictureNonVisual:

    def test_ctor(self, PictureNonVisual):
        graphic = PictureNonVisual()
        xml = tostring(graphic.to_tree())
        expected = """
        <nvPicPr>
          <cNvPr descr="Name of file" id="0" name="Image 1" />
          <cNvPicPr />
        </nvPicPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureNonVisual):
        src = """
        <nvPicPr>
          <cNvPr descr="Name of file" id="0" name="Image 1" />
          <cNvPicPr />
        </nvPicPr>
        """
        node = fromstring(src)
        graphic = PictureNonVisual.from_tree(node)
        assert graphic == PictureNonVisual()



@pytest.fixture
def PictureFrame():
    from ..picture import PictureFrame
    return PictureFrame


class TestPicture:

    def test_ctor(self, PictureFrame):
        graphic = PictureFrame()
        xml = tostring(graphic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <nvPicPr>
            <cNvPr descr="Name of file" id="0" name="Image 1" />
            <cNvPicPr />
          </nvPicPr>
          <blipFill>
             <a:stretch >
               <a:fillRect/>
            </a:stretch>
          </blipFill>
          <spPr>
            <a:ln>
              <a:prstDash val="solid" />
            </a:ln>
          </spPr>
        </pic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureFrame):
        src = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <nvPicPr>
            <cNvPr descr="Picture" id="1" name="Image 1"/>
            <cNvPicPr/>
        </nvPicPr>
        <blipFill>
            <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" cstate="print" r:embed="rId1"/>
            <a:stretch>
                <a:fillRect/>
            </a:stretch>
        </blipFill>
        <spPr>
            <a:xfrm>
                <a:off x="303" y="0"/>
                <a:ext cx="321" cy="88"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
            <a:ln>
              <a:prstDash val="solid" />
            </a:ln>
        </spPr>
        </pic>
        """
        node = fromstring(src)
        graphic = PictureFrame.from_tree(node)
        xml = tostring(graphic.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff
