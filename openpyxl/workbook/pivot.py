from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import Relation

class PivotCache(Serialisable):

    tagname = "pivotCache"

    cacheId = Integer()
    id = Relation()

    def __init__(self,
                 cacheId=None,
                 id=None
                ):
        self.cacheId = cacheId
        self.id = id


class PivotCacheList(Serialisable):

    tagname = "pivotCaches"

    pivotCache = Sequence(expected_type=PivotCache, )

    __elements__ = ('pivotCache',)

    def __init__(self,
                 pivotCache=(),
                ):
        self.pivotCache = pivotCache
