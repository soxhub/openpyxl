from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.mark.xfail
def test_style_copy():
    from .. import Style
    st1 = Style()
    st2 = st1.copy()
    assert st1 == st2
    assert st1.font is not st2.font
