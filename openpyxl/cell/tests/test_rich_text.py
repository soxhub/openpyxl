# Copyright (c) 2010-2021 openpyxl

# 3rd party imports
import pytest

# package imports

from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.styles.fonts import Font

def test_rich_text_create_single():
    text = CellRichText("ABC")
    assert text[0] == "ABC"

def test_rich_text_create_multi():
    text = CellRichText("ABC", "DEF", "GHI")
    assert len(text) == 3

def test_rich_text_create_text_block():
    text = CellRichText(TextBlock(font=Font(), text="ABC"))
    assert getattr(text[0], "text") == "ABC"

def test_rich_text_append():
    text = CellRichText()
    text.append(TextBlock(font=Font(), text="ABC"))
    assert getattr(text[0], "text") == "ABC"

def test_rich_text_extend():
    text = CellRichText()
    text.extend(("ABC", "DEF"))
    assert len(text) == 2
