from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Scaling():
    from ..axis import Scaling
    return Scaling


def test_scaling(Scaling):

    scale = Scaling()
    xml = tostring(scale.to_tree())
    expected = """
    <scaling>
       <orientation val="minMax"></orientation>
    </scaling>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def _BaseAxis():
    from ..axis import _BaseAxis
    return _BaseAxis


class TestAxis:

    def test_ctor(self, _BaseAxis, Scaling):
        axis = _BaseAxis(axId=10, crossAx=100)
        xml = tostring(axis.to_tree(tagname="baseAxis"))
        expected = """
        <baseAxis>
            <axId val="10"></axId>
            <scaling>
              <orientation val="minMax"></orientation>
            </scaling>
            <crossAx val="100"></crossAx>
        </baseAxis>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



@pytest.fixture
def CatAx():
    from ..axis import CatAx
    return CatAx


class TestCatAx:

    def test_ctor(self, CatAx):
        axis = CatAx(axId=10, crossAx=100)
        xml = tostring(axis.to_tree())
        expected = """
        <catAx>
         <axId val="10"></axId>
         <scaling>
           <orientation val="minMax"></orientation>
         </scaling>
         <crossAx val="100"></crossAx>
         <lblOffset val="100"></lblOffset>
        </catAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def from_xml(self, CatAx):
        src = """
        <catAx>
        <axId val="2065276984"/>
        <scaling>
          <orientation val="minMax"/>
        </scaling>
        <delete val="0"/>
        <axPos val="b"/>
        <majorTickMark val="out"/>
        <minorTickMark val="none"/>
        <tickLblPos val="nextTo"/>
        <crossAx val="2056619928"/>
        <crosses val="autoZero"/>
        <auto val="1"/>
        <lblAlgn val="ctr"/>
        <lblOffset val="100"/>
        <noMultiLvlLbl val="0"/>
        </catAx>
        """
        node = fromstring(src)
        axis = CatAx.from_tree(node)
        assert axis.scaling.orientation == "minMax"
        assert axis.auto is True
        assert axis.majorTickMark == "out"
        assert axis.minorTickMark is None


@pytest.fixture
def ValAx():
    from ..axis import ValAx
    return ValAx


class TestValAx:

    def test_ctor(self, ValAx):
        axis = ValAx(axId=100, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <valAx>
          <axId val="100"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <crossAx val="10"></crossAx>
        </valAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ValAx):
        src = """
        <valAx>
            <axId val="2056619928"/>
            <scaling>
                <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="l"/>
            <majorGridlines/>
            <numFmt formatCode="General" sourceLinked="1"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="2065276984"/>
            <crosses val="autoZero"/>
            <crossBetween val="between"/>
        </valAx>
        """
        node = fromstring(src)
        axis = ValAx.from_tree(node)
        assert axis.delete is False
        assert axis.crossAx == 2065276984
        assert axis.crossBetween == "between"


@pytest.fixture
def DateAx():
    from ..axis import DateAx
    return DateAx


class TestDateAx:


    def test_ctor(self, DateAx):
        axis = DateAx(axId=500, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <dateAx>
           <axId val="500"></axId>
           <scaling>
             <orientation val="minMax"></orientation>
           </scaling>
           <crossAx val="10"></crossAx>
        </dateAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DateAx):
        src = """
        <dateAx>
          <axId val="20"/>
          <scaling>
            <orientation val="minMax"/>
          </scaling>
          <delete val="0"/>
          <axPos val="b"/>
          <numFmt formatCode="d\-mmm" sourceLinked="1"/>
          <majorTickMark val="out"/>
          <minorTickMark val="none"/>
          <tickLblPos val="nextTo"/>
          <crossAx val="10"/>
          <crosses val="autoZero"/>
          <auto val="1"/>
          <lblOffset val="100"/>
          <baseTimeUnit val="months"/>
        </dateAx>
        """
        node = fromstring(src)
        axis = DateAx.from_tree(node)
        assert dict(axis) == {}
        assert axis.axId == 20
        assert axis.crossAx == 10


@pytest.fixture
def SerAx():
    from ..axis import SerAx
    return SerAx


class TestSerAx:

    def test_ctor(self, SerAx):
        axis = SerAx(axId=1000, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <serAx>
          <axId val="1000"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <crossAx val="10"></crossAx>
        </serAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SerAx):
        src = """
        <serAx>
          <axId val="1000"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <crossAx val="10"></crossAx>
        </serAx>
        """
        node = fromstring(src)
        axis = SerAx.from_tree(node)
        assert dict(axis) == {}


@pytest.fixture
def DisplayUnitsLabel():
    from ..axis import DisplayUnitsLabel
    return DisplayUnitsLabel


class TestDispUnitsLbl:

    def test_ctor(self, DisplayUnitsLabel):
        axis = DisplayUnitsLabel()
        xml = tostring(axis.to_tree())
        expected = """
        <dispUnitsLbl />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DisplayUnitsLabel):
        src = """
        <dispUnitsLbl />
        """
        node = fromstring(src)
        axis = DisplayUnitsLabel.from_tree(node)
        assert axis == DisplayUnitsLabel()


@pytest.fixture
def DisplayUnits():
    from ..axis import DisplayUnits
    return DisplayUnits


class TestDisplayUnits:

    def test_ctor(self, DisplayUnits):
        axis = DisplayUnits()
        xml = tostring(axis.to_tree())
        expected = """
        <dispUnits />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DisplayUnits):
        src = """
        <dispUnits />
        """
        node = fromstring(src)
        axis = DisplayUnits.from_tree(node)
        assert axis == DisplayUnits()
