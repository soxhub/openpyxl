from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from .. protection import hash_password


def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'
