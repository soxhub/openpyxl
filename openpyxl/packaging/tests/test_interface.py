from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

def test_interface():
    from ..interface import ISerialisableFile

    class DummyFile(ISerialisableFile):

        pass

    with pytest.raises(TypeError):

        df = DummyFile()
