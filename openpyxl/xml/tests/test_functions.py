import pytest


def test_safe_iterator_none():
    from .. functions import safe_iterator
    seq = safe_iterator(None)
    assert seq == []


@pytest.mark.parametrize("xml, tag",
                         [
                             ("<root xmlns='http://openpyxl.org/ns' />", "root"),
                             ("<root />", "root"),
                         ]
                         )
def test_localtag(xml, tag):
    from .. functions import localname
    from .. functions import fromstring
    node = fromstring(xml)
    assert localname(node) == tag
