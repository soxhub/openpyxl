from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    DateTime,
    Bool,
    Float,
    String,
    Integer,
    Sequence,
)

from openpyxl.descriptors.excel import HexBinary, ExtensionList
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedBool,
)

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring


class Index(Serialisable):

    tagname = "x"

    v = Integer(allow_none=True)

    def __init__(self,
                 v=0,
                ):
        self.v = v


class Tuple(Serialisable):

    fld = Integer()
    hier = Integer()
    item = Integer()

    def __init__(self,
                 fld=None,
                 hier=None,
                 item=None,
                ):
        self.fld = fld
        self.hier = hier
        self.item = item


class TupleList(Serialisable):

    c = Integer(allow_none=True)
    tpl = Typed(expected_type=Tuple, )

    __elements__ = ('tpl',)

    def __init__(self,
                 c=None,
                 tpl=None,
                ):
        self.c = c
        self.tpl = tpl


class SharedItem(Serialisable):

    pass


class Missing(SharedItem):

    tagname = "m"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b

class Number(SharedItem):

    tagname = "n"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    v = Float()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Error(SharedItem):

    tagname = "e"

    tpls = Typed(expected_type=TupleList, allow_none=True)
    x = Sequence(expected_type=Index)
    v = String()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=None,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Boolean(SharedItem):

    tagname = "b"

    x = Sequence(expected_type=Index)
    v = Bool()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                ):
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp


class Text(SharedItem):

    tagname = "s"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    v = String()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                 ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class PivotDateTime(Serialisable):

    x = Sequence(expected_type=Index)
    v = DateTime()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 ):
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp


class Record(Serialisable):

    tagname = "r"

    # some elements are choice
    m = Sequence(expected_type=Missing)
    n = Sequence(expected_type=Number)
    b = Sequence(expected_type=Boolean)
    e = Sequence(expected_type=Error)
    s = Sequence(expected_type=Text)
    d = Sequence(expected_type=PivotDateTime)
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
