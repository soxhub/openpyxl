from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Integer,
    Sequence,
)

class PivotCache(Serialisable):

    cacheId = Integer()

    def __init__(self,
                 cacheId=None,
                ):
        self.cacheId = cacheId


class PivotCacheList(Serialisable):

    pivotCache = Sequence(expected_type=PivotCache, )

    __elements__ = ('pivotCache',)

    def __init__(self,
                 pivotCache=None,
                ):
        self.pivotCache = pivotCache
