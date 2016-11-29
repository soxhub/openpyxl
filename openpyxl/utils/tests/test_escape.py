import pytest


def test_escape():
    from ..escape import escape
    value = "My name is \nNewline\r"
    assert escape(value) == "My name is _x000a_Newline_x000d_"


def test_unescape():
    from ..escape import unescape
    value = "My name is _x000a_Newline_x000d_"
    assert unescape(value) == "My name is \nNewline\r"
