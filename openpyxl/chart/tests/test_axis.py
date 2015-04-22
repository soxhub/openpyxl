from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring
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
