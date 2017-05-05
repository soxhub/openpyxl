from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    Sequence,
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedBool,
)

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring

from .fields import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    TupleList,
    DateTimeField,
    SharedItem,
    Index,
)


class Record(Serialisable):

    tagname = "r"

    # some elements are choice
    m = Sequence(expected_type=Missing)
    n = Sequence(expected_type=Number)
    b = Sequence(expected_type=Boolean)
    e = Sequence(expected_type=Error)
    s = Sequence(expected_type=Text)
    d = Sequence(expected_type=DateTimeField)
    x = Sequence(expected_type=Index)

    __elements__ = ('m', 'n', 'b', 'e', 's', 'd', 'x')

    def __init__(self,
                 m=(),
                 n=(),
                 b=(),
                 e=(),
                 s=(),
                 d=(),
                 x=(),
                ):
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d
        self.x = x


class RecordList(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
    _id = 1
    _path = "/xl/pivotCache/pivotCacheRecords{0}.xml"

    tagname ="pivotCacheRecords"

    r = Sequence(expected_type=Record, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('r', )
    __attrs__ = ('count', )

    def __init__(self,
                 count=None,
                 r=(),
                 extLst=None,
                ):
        self.r = r
        self.extLst = extLst


    @property
    def count(self):
        return len(self.r)


    def to_tree(self):
        tree = super(RecordList, self).to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Write to zipfile and update manifest
        """
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        pass
