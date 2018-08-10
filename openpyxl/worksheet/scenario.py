from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
)
from .cell_range import MultiCellRange


class InputCells(Serialisable):

    tagname = "inputCells"

    r = String()
    deleted = Bool(allow_none=True)
    undone = Bool(allow_none=True)
    val = String()
    numFmtId = Integer()

    def __init__(self,
                 r=None,
                 deleted=False,
                 undone=False,
                 val=None,
                 numFmtId=None,
                ):
        self.r = r
        self.deleted = deleted
        self.undone = undone
        self.val = val
        self.numFmtId = numFmtId


class Scenario(Serialisable):

    tagname = "scenario"

    inputCells = Sequence(expected_type=InputCells)
    name = String()
    locked = Bool(allow_none=True)
    hidden = Bool(allow_none=True)
    count = Integer(allow_none=True)
    user = String(allow_none=True)
    comment = String(allow_none=True)

    __elements__ = ('inputCells',)

    def __init__(self,
                 inputCells=(),
                 name=None,
                 locked=False,
                 hidden=False,
                 count=None,
                 user=None,
                 comment=None,
                ):
        self.inputCells = inputCells
        self.name = name
        self.locked = locked
        self.hidden = hidden
        self.count = count
        self.user = user
        self.comment = comment


class Scenarios(Serialisable):

    tagname = "scenarios"

    scenario = Sequence(expected_type=Scenario)
    current = Integer(allow_none=True)
    show = Integer(allow_none=True)
    sqref = Convertible(expected_type=MultiCellRange)

    __elements__ = ('scenario',)

    def __init__(self,
                 scenario=(),
                 current=None,
                 show=None,
                 sqref=None,
                ):
        self.scenario = scenario
        self.current = current
        self.show = show
        self.sqref = sqref
