from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
from openpyxl.compat import unicode

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
)


class AuthorList(Serialisable):

    tagname = "authors"

    author = Sequence(expected_type=unicode)

    def __init__(self,
                 author=(),
                ):
        self.author = author
