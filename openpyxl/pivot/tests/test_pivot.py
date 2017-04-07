
from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotField():
    from ..pivot import PivotField
    return PivotField


class TestPivotField:

    def test_ctor(self, PivotField):
        field = PivotField()
        xml = tostring(field.to_tree())
        expected = """
        <pivotField allDrilled="0" autoShow="0" avgSubtotal="0" compact="1" countASubtotal="0" countSubtotal="0" dataField="0" defaultAttributeDrillState="0" defaultSubtotal="1" dragOff="1" dragToCol="1" dragToData="1" dragToPage="1" dragToRow="1" hiddenLevel="0" hideNewItems="0" includeNewItemsInFilter="0" insertBlankRow="0" insertPageBreak="0" itemPageCount="10" maxSubtotal="0" measureFilter="0" minSubtotal="0" multipleItemSelectionAllowed="0" nonAutoSortDefault="0" outline="1" productSubtotal="0" serverField="0" showAll="1" showDropDowns="1" showPropAsCaption="0" showPropCell="0" showPropTip="0" sortType="manual" stdDevPSubtotal="0" stdDevSubtotal="0" subtotalTop="1" sumSubtotal="0" topAutoShow="0" varPSubtotal="0" varSubtotal="0"></pivotField>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotField):
        src = """
        <pivotField />
        """
        node = fromstring(src)
        field = PivotField.from_tree(node)
        assert field == PivotField()

