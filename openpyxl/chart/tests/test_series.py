from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def BarSer():
    from ..series import BarSer
    return BarSer


class TestBarSer:

    def test_from_tree(self, BarSer):
        src = """
        <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:invertIfNegative val="0"/>
          <c:val>
            <c:numRef>
                <c:f>Blatt1!$A$1:$A$12</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        """
        node = fromstring(src)
        ser = BarSer.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        #assert ser.val.numRef.ref == ""
