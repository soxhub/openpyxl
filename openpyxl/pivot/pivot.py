from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    NoneSet,
    Set,
    Float,
    Bool,
    DateTime,
    String,
    Alias,
    Bool,
    Sequence,
)

from openpyxl.descriptors.excel import ExtensionList, Relation
from openpyxl.descriptors.nested import NestedInteger
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring
from openpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)

from openpyxl.worksheet.filters import (
    AutoFilter,
    CellRange,
    ColorFilter,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
    DynamicFilter,
    FilterColumn,
    Filters,
    IconFilter,
    SortCondition,
    SortState,
    Top10,
)


class HierarchyUsage(Serialisable):

    hierarchyUsage = Integer()

    def __init__(self,
                 hierarchyUsage=None,
                ):
        self.hierarchyUsage = hierarchyUsage


class ColHierarchiesUsage(Serialisable):

    count = Integer()
    colHierarchyUsage = Typed(expected_type=HierarchyUsage, )

    __elements__ = ('colHierarchyUsage',)

    def __init__(self,
                 count=None,
                 colHierarchyUsage=None,
                ):
        self.colHierarchyUsage = colHierarchyUsage


class RowHierarchiesUsage(Serialisable):

    count = Integer()
    rowHierarchyUsage = Typed(expected_type=HierarchyUsage, )

    __elements__ = ('rowHierarchyUsage',)

    def __init__(self,
                 count=None,
                 rowHierarchyUsage=None,
                ):
        self.rowHierarchyUsage = rowHierarchyUsage


class PivotFilter(Serialisable):

    fld = Integer()
    mpFld = Integer(allow_none=True)
    type = Set(values=(['unknown', 'count', 'percent', 'sum', 'captionEqual',
                        'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith',
                        'captionEndsWith', 'captionNotEndsWith', 'captionContains',
                        'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual',
                        'captionLessThan', 'captionLessThanOrEqual', 'captionBetween',
                        'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan',
                        'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual',
                        'valueBetween', 'valueNotBetween', 'dateEqual', 'dateNotEqual',
                        'dateOlderThan', 'dateOlderThanOrEqual', 'dateNewerThan',
                        'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween', 'tomorrow',
                        'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth',
                        'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter',
                        'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
                        'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11',
                        'M12']))
    evalOrder = Integer(allow_none=True)
    id = Integer()
    iMeasureHier = Integer(allow_none=True)
    iMeasureFld = Integer(allow_none=True)
    name = String(allow_none=True)
    description = String()
    stringValue1 = String()
    stringValue2 = String()
    autoFilter = Typed(expected_type=AutoFilter, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter', 'extLst')

    def __init__(self,
                 fld=None,
                 mpFld=None,
                 type=None,
                 evalOrder=None,
                 id=None,
                 iMeasureHier=None,
                 iMeasureFld=None,
                 name=None,
                 description=None,
                 stringValue1=None,
                 stringValue2=None,
                 autoFilter=None,
                 extLst=None,
                ):
        self.fld = fld
        self.mpFld = mpFld
        self.type = type
        self.evalOrder = evalOrder
        self.id = id
        self.iMeasureHier = iMeasureHier
        self.iMeasureFld = iMeasureFld
        self.name = name
        self.description = description
        self.stringValue1 = stringValue1
        self.stringValue2 = stringValue2
        self.autoFilter = autoFilter
        self.extLst = extLst


class PivotFilters(Serialisable):

    count = Integer()
    filter = Typed(expected_type=PivotFilter, allow_none=True)

    __elements__ = ('filter',)

    def __init__(self,
                 count=None,
                 filter=None,
                ):
        self.filter = filter


class PivotTableStyle(Serialisable):

    tagname = "pivotTableStyleInfo"

    name = String()
    showRowHeaders = Bool()
    showColHeaders = Bool()
    showRowStripes = Bool()
    showColStripes = Bool()
    showLastColumn = Bool(allow_none=True)

    def __init__(self,
                 name=None,
                 showRowHeaders=None,
                 showColHeaders=None,
                 showRowStripes=None,
                 showColStripes=None,
                 showLastColumn=None,
                ):
        self.name = name
        self.showRowHeaders = showRowHeaders
        self.showColHeaders = showColHeaders
        self.showRowStripes = showRowStripes
        self.showColStripes = showColStripes
        self.showLastColumn = showLastColumn


class Member(Serialisable):

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class MemberList(Serialisable):

    level = Integer(allow_none=True)
    member = Typed(expected_type=Member, )

    __elements__ = ('member',)

    def __init__(self,
                 count=None,
                 level=None,
                 member=None,
                ):
        self.level = level
        self.member = member

    @property
    def count(self):
        return len(self.member)


class MemberProperty(Serialisable):

    name = String(allow_none=True)
    showCell = Bool(allow_none=True)
    showTip = Bool(allow_none=True)
    showAsCaption = Bool(allow_none=True)
    nameLen = Integer(allow_none=True)
    pPos = Integer(allow_none=True)
    pLen = Integer(allow_none=True)
    level = Integer(allow_none=True)
    field = Integer()

    def __init__(self,
                 name=None,
                 showCell=None,
                 showTip=None,
                 showAsCaption=None,
                 nameLen=None,
                 pPos=None,
                 pLen=None,
                 level=None,
                 field=None,
                ):
        self.name = name
        self.showCell = showCell
        self.showTip = showTip
        self.showAsCaption = showAsCaption
        self.nameLen = nameLen
        self.pPos = pPos
        self.pLen = pLen
        self.level = level
        self.field = field


class MemberPropertyList(Serialisable):

    mp = Typed(expected_type=MemberProperty, )

    __elements__ = ('mp',)

    def __init__(self,
                 count=None,
                 mp=None,
                ):
        self.mp = mp


    @property
    def count(self):
        return len(self.mp)


class PivotHierarchy(Serialisable):

    outline = Bool()
    multipleItemSelectionAllowed = Bool()
    subtotalTop = Bool()
    showInFieldList = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    includeNewItemsInFilter = Bool()
    caption = String(allow_none=True)
    mps = Typed(expected_type=MemberPropertyList, allow_none=True)
    members = Typed(expected_type=MemberList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('mps', 'members', 'extLst')

    def __init__(self,
                 outline=None,
                 multipleItemSelectionAllowed=None,
                 subtotalTop=None,
                 showInFieldList=None,
                 dragToRow=None,
                 dragToCol=None,
                 dragToPage=None,
                 dragToData=None,
                 dragOff=None,
                 includeNewItemsInFilter=None,
                 caption=None,
                 mps=None,
                 members=None,
                 extLst=None,
                ):
        self.outline = outline
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.subtotalTop = subtotalTop
        self.showInFieldList = showInFieldList
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.caption = caption
        self.mps = mps
        self.members = members
        self.extLst = extLst


class ChartFormat(Serialisable):

    chart = Integer()
    format = Integer()
    series = Bool()
    #pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 chart=None,
                 format=None,
                 series=None,
                 pivotArea=None,
                ):
        self.chart = chart
        self.format = format
        self.series = series
        self.pivotArea = pivotArea


class PivotAreaReference(Serialisable):

    field = Integer(allow_none=True)
    selected = Bool()
    byPosition = Bool()
    relative = Bool()
    defaultSubtotal = Bool()
    sumSubtotal = Bool()
    countASubtotal = Bool()
    avgSubtotal = Bool()
    maxSubtotal = Bool()
    minSubtotal = Bool()
    productSubtotal = Bool()
    countSubtotal = Bool()
    stdDevSubtotal = Bool()
    stdDevPSubtotal = Bool()
    varSubtotal = Bool()
    varPSubtotal = Bool()
    x = NestedInteger(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 field=None,
                 count=None,
                 selected=None,
                 byPosition=None,
                 relative=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 x=None,
                 extLst=None,
                ):
        self.field = field
        self.selected = selected
        self.byPosition = byPosition
        self.relative = relative
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.x = x
        self.extLst = extLst


    @property
    def count(self):
        return len(self.field)


class PivotAreaReferenceList(Serialisable):

    reference = Typed(expected_type=PivotAreaReference, )

    __elements__ = ('reference',)

    def __init__(self,
                 count=None,
                 reference=None,
                ):
        self.reference = reference


    @property
    def count(self):
        return len(self.reference)


class PivotArea(Serialisable):

    field = Integer(allow_none=True)
    type = NoneSet(values=(['normal', 'data', 'all', 'origin', 'button', 'topEnd', 'topRight']))
    dataOnly = Bool()
    labelOnly = Bool()
    grandRow = Bool()
    grandCol = Bool()
    cacheIndex = Bool()
    outline = Bool()
    offset = String()
    collapsedLevelsAreSubtotals = Bool()
    axis = Set(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    fieldPosition = Integer(allow_none=True)
    references = Typed(expected_type=PivotAreaReferenceList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('references', 'extLst')

    def __init__(self,
                 field=None,
                 type=None,
                 dataOnly=None,
                 labelOnly=None,
                 grandRow=None,
                 grandCol=None,
                 cacheIndex=None,
                 outline=None,
                 offset=None,
                 collapsedLevelsAreSubtotals=None,
                 axis=None,
                 fieldPosition=None,
                 references=None,
                 extLst=None,
                ):
        self.field = field
        self.type = type
        self.dataOnly = dataOnly
        self.labelOnly = labelOnly
        self.grandRow = grandRow
        self.grandCol = grandCol
        self.cacheIndex = cacheIndex
        self.outline = outline
        self.offset = offset
        self.collapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals
        self.axis = axis
        self.fieldPosition = fieldPosition
        self.references = references
        self.extLst = extLst


class PivotAreaList(Serialisable):

    pivotArea = Typed(expected_type=PivotArea, allow_none=True)

    __elements__ = ('pivotArea',)

    def __init__(self,
                 count=None,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


    @property
    def count(self):
        return len(self.pivotArea)


class ConditionalFormat(Serialisable):

    scope = Set(values=(['selection', 'data', 'field']))
    type = NoneSet(values=(['all', 'row', 'column']))
    priority = Integer()
    pivotAreas = Typed(expected_type=PivotAreaList, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotAreas', 'extLst')

    def __init__(self,
                 scope=None,
                 type=None,
                 priority=None,
                 pivotAreas=None,
                 extLst=None,
                ):
        self.scope = scope
        self.type = type
        self.priority = priority
        self.pivotAreas = pivotAreas
        self.extLst = extLst


class Format(Serialisable):

    tagname = "format"

    action = Set(values=(['blank', 'formatting', 'drill', 'formula']))
    dxfId = Integer()
    pivotArea = Typed(expected_type=PivotArea, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotArea',)

    def __init__(self,
                 action=None,
                 dxfId=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.action = action
        self.dxfId = dxfId
        self.pivotArea = pivotArea
        self.extLst = extLst


class DataField(Serialisable):

    tagname = "dataField"

    name = String(allow_none=True)
    fld = Integer()
    subtotal = Set(values=(['average', 'count', 'countNums', 'max', 'min',
                            'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp']))
    showDataAs = Set(values=(['normal', 'difference', 'percent',
                              'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol',
                              'percentOfTotal', 'index']))
    baseField = Integer()
    baseItem = Integer()
    numFmtId = Integer(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)


    def __init__(self,
                 name=None,
                 fld=None,
                 subtotal="sum",
                 showDataAs="normal",
                 baseField=-1,
                 baseItem=1048832,
                 numFmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fld = fld
        self.subtotal = subtotal
        self.showDataAs = showDataAs
        self.baseField = baseField
        self.baseItem = baseItem
        self.numFmtId = numFmtId
        self.extLst = extLst


class PageField(Serialisable):

    tagname = "PageField"

    fld = Integer()
    item = Integer(allow_none=True)
    hier = Integer()
    name = String()
    cap = String()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 fld=None,
                 item=None,
                 hier=None,
                 name=None,
                 cap=None,
                 extLst=None,
                ):
        self.fld = fld
        self.item = item
        self.hier = hier
        self.name = name
        self.cap = cap
        self.extLst = extLst


class RowColItem(Serialisable):

    tagname = "i"

    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min',
                     'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand',
                     'blank']))
    r = Integer()
    i = Integer()
    x = NestedInteger(allow_none=True, attribute="v")

    __elements__ = ('x',)

    def __init__(self,
                 t="data",
                 r=0,
                 i=0,
                 x=None,
                ):
        self.t = t
        self.r = r
        self.i = i
        self.x = x


class RowColField(Serialisable):

    tagname = "field"

    x = Integer()

    def __init__(self,
                 x=None,
                ):
        self.x = x


class AutoSortScope(Serialisable):

    pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


class FieldItem(Serialisable):

    tagname = "item"

    n = String(allow_none=True)
    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min',
                     'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand',
                     'blank']))
    h = Bool()
    s = Bool()
    sd = Bool()
    f = Bool()
    m = Bool()
    c = Bool()
    x = Integer(allow_none=True)
    d = Bool()
    e = Bool()

    def __init__(self,
                 n=None,
                 t="data",
                 h=False,
                 s=False,
                 sd=True,
                 f=False,
                 m=False,
                 c=False,
                 x=None,
                 d=False,
                 e=False,
                ):
        self.n = n
        self.t = t
        self.h = h
        self.s = s
        self.sd = sd
        self.f = f
        self.m = m
        self.c = c
        self.x = x
        self.d = d
        self.e = e


class PivotField(Serialisable):

    tagname = "pivotField"

    name = String(allow_none=True)
    axis = NoneSet(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    dataField = Bool()
    subtotalCaption = String(allow_none=True)
    showDropDowns = Bool()
    hiddenLevel = Bool()
    uniqueMemberProperty = String(allow_none=True)
    compact = Bool()
    allDrilled = Bool()
    numFmtId = Integer(allow_none=True)
    outline = Bool()
    subtotalTop = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    multipleItemSelectionAllowed = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    showAll = Bool()
    insertBlankRow = Bool()
    serverField = Bool()
    insertPageBreak = Bool()
    autoShow = Bool()
    topAutoShow = Bool()
    hideNewItems = Bool()
    measureFilter = Bool()
    includeNewItemsInFilter = Bool()
    itemPageCount = Integer(allow_none=True)
    sortType = Set(values=(['manual', 'ascending', 'descending']))
    dataSourceSort = Bool(allow_none=True)
    nonAutoSortDefault = Bool()
    rankBy = Integer(allow_none=True)
    defaultSubtotal = Bool(allow_none=True)
    sumSubtotal = Bool(allow_none=True)
    countASubtotal = Bool(allow_none=True)
    avgSubtotal = Bool(allow_none=True)
    maxSubtotal = Bool(allow_none=True)
    minSubtotal = Bool(allow_none=True)
    productSubtotal = Bool(allow_none=True)
    countSubtotal = Bool(allow_none=True)
    stdDevSubtotal = Bool(allow_none=True)
    stdDevPSubtotal = Bool(allow_none=True)
    varSubtotal = Bool(allow_none=True)
    varPSubtotal = Bool(allow_none=True)
    showPropCell = Bool(allow_none=True)
    showPropTip = Bool(allow_none=True)
    showPropAsCaption = Bool(allow_none=True)
    defaultAttributeDrillState = Bool(allow_none=True)
    items = NestedSequence(expected_type=FieldItem, count=True)
    autoSortScope = Typed(expected_type=AutoSortScope, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('items', 'autoSortScope', )

    def __init__(self,
                 name=None,
                 axis=None,
                 dataField=False,
                 subtotalCaption=None,
                 showDropDowns=True,
                 hiddenLevel=False,
                 uniqueMemberProperty=None,
                 compact=True,
                 allDrilled=False,
                 numFmtId=None,
                 outline=True,
                 subtotalTop=True,
                 dragToRow=True,
                 dragToCol=True,
                 multipleItemSelectionAllowed=False,
                 dragToPage=True,
                 dragToData=True,
                 dragOff=True,
                 showAll=True,
                 insertBlankRow=False,
                 serverField=False,
                 insertPageBreak=False,
                 autoShow=False,
                 topAutoShow=False,
                 hideNewItems=None,
                 measureFilter=False,
                 includeNewItemsInFilter=False,
                 itemPageCount=10,
                 sortType="manual",
                 dataSourceSort=None,
                 nonAutoSortDefault=False,
                 rankBy=None,
                 defaultSubtotal=True,
                 sumSubtotal=False,
                 countASubtotal=False,
                 avgSubtotal=False,
                 maxSubtotal=False,
                 minSubtotal=False,
                 productSubtotal=False,
                 countSubtotal=False,
                 stdDevSubtotal=False,
                 stdDevPSubtotal=False,
                 varSubtotal=False,
                 varPSubtotal=False,
                 showPropCell=False,
                 showPropTip=False,
                 showPropAsCaption=False,
                 defaultAttributeDrillState=False,
                 items=(),
                 autoSortScope=None,
                 extLst=None,
                ):
        self.name = name
        self.axis = axis
        self.dataField = dataField
        self.subtotalCaption = subtotalCaption
        self.showDropDowns = showDropDowns
        self.hiddenLevel = hiddenLevel
        self.uniqueMemberProperty = uniqueMemberProperty
        self.compact = compact
        self.allDrilled = allDrilled
        self.numFmtId = numFmtId
        self.outline = outline
        self.subtotalTop = subtotalTop
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.showAll = showAll
        self.insertBlankRow = insertBlankRow
        self.serverField = serverField
        self.insertPageBreak = insertPageBreak
        self.autoShow = autoShow
        self.topAutoShow = topAutoShow
        self.hideNewItems = hideNewItems
        self.measureFilter = measureFilter
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.itemPageCount = itemPageCount
        self.sortType = sortType
        self.dataSourceSort = dataSourceSort
        self.nonAutoSortDefault = nonAutoSortDefault
        self.rankBy = rankBy
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.showPropCell = showPropCell
        self.showPropTip = showPropTip
        self.showPropAsCaption = showPropAsCaption
        self.defaultAttributeDrillState = defaultAttributeDrillState
        self.items = items
        self.autoSortScope = autoSortScope
        self.extLst = extLst


class Location(Serialisable):

    tagname = "location"

    ref = String()
    firstHeaderRow = Integer()
    firstDataRow = Integer()
    firstDataCol = Integer()
    rowPageCount = Integer(allow_none=True)
    colPageCount = Integer(allow_none=True)

    def __init__(self,
                 ref=None,
                 firstHeaderRow=None,
                 firstDataRow=None,
                 firstDataCol=None,
                 rowPageCount=None,
                 colPageCount=None,
                ):
        self.ref = ref
        self.firstHeaderRow = firstHeaderRow
        self.firstDataRow = firstDataRow
        self.firstDataCol = firstDataCol
        self.rowPageCount = rowPageCount
        self.colPageCount = colPageCount


class PivotTableDefinition(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
    _id = 1
    _path = "/xl/pivotTables/pivotTable{0}.xml"

    tagname = "pivotTableDefinition"
    cache = None

    name = String()
    cacheId = Integer()
    dataOnRows = Bool()
    dataPosition = Integer(allow_none=True)
    dataCaption = String()
    grandTotalCaption = String(allow_none=True)
    errorCaption = String(allow_none=True)
    showError = Bool()
    missingCaption = String(allow_none=True)
    showMissing = Bool()
    pageStyle = String(allow_none=True)
    pivotTableStyle = String(allow_none=True)
    vacatedStyle = String(allow_none=True)
    tag = String(allow_none=True)
    updatedVersion = Integer()
    minRefreshableVersion = Integer()
    asteriskTotals = Bool()
    showItems = Bool()
    editData = Bool()
    disableFieldList = Bool()
    showCalcMbrs = Bool()
    visualTotals = Bool()
    showMultipleLabel = Bool()
    showDataDropDown = Bool()
    showDrill = Bool()
    printDrill = Bool()
    showMemberPropertyTips = Bool()
    showDataTips = Bool()
    enableWizard = Bool()
    enableDrill = Bool()
    enableFieldProperties = Bool()
    preserveFormatting = Bool()
    useAutoFormatting = Bool()
    pageWrap = Integer()
    pageOverThenDown = Bool()
    subtotalHiddenItems = Bool()
    rowGrandTotals = Bool()
    colGrandTotals = Bool()
    fieldPrintTitles = Bool()
    itemPrintTitles = Bool()
    mergeItem = Bool()
    showDropZones = Bool()
    createdVersion = Integer()
    indent = Integer()
    showEmptyRow = Bool()
    showEmptyCol = Bool()
    showHeaders = Bool()
    compact = Bool()
    outline = Bool()
    outlineData = Bool()
    compactData = Bool()
    published = Bool()
    gridDropZones = Bool()
    immersive = Bool()
    multipleFieldFilters = Bool()
    chartFormat = Integer()
    rowHeaderCaption = String(allow_none=True)
    colHeaderCaption = String(allow_none=True)
    fieldListSortAscending = Bool()
    mdxSubqueries = Bool()
    customListSort = Bool(allow_none=True)
    autoFormatId = Integer(allow_none=True)
    applyNumberFormats = Bool()
    applyBorderFormats = Bool()
    applyFontFormats = Bool()
    applyPatternFormats = Bool()
    applyAlignmentFormats = Bool()
    applyWidthHeightFormats = Bool()
    location = Typed(expected_type=Location, )
    pivotFields = NestedSequence(expected_type=PivotField, count=True)
    rowFields = NestedSequence(expected_type=RowColField, count=True)
    rowItems = NestedSequence(expected_type=RowColItem, count=True)
    colFields = NestedSequence(expected_type=RowColField, count=True)
    colItems = NestedSequence(expected_type=RowColItem, count=True)
    pageFields = NestedSequence(expected_type=RowColField, count=True)
    dataFields = NestedSequence(expected_type=DataField, count=True)
    formats = NestedSequence(expected_type=Format, count=True)
    conditionalFormats = NestedSequence(expected_type=ConditionalFormat, count=True)
    chartFormats = NestedSequence(expected_type=ChartFormat, count=True)
    pivotHierarchies = NestedSequence(expected_type=PivotHierarchy, count=True)
    pivotTableStyleInfo = Typed(expected_type=PivotTableStyle, allow_none=True)
    filters = NestedSequence(expected_type=PivotFilter, count=True)
    rowHierarchiesUsage = Typed(expected_type=RowHierarchiesUsage, allow_none=True)
    colHierarchiesUsage = Typed(expected_type=ColHierarchiesUsage, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    id = Relation()

    __elements__ = ('location', 'pivotFields', 'rowFields', 'rowItems',
                    'colFields', 'colItems', 'pageFields', 'dataFields', 'formats',
                    'conditionalFormats', 'chartFormats', 'pivotHierarchies',
                    'pivotTableStyleInfo', 'filters', 'rowHierarchiesUsage',
                    'colHierarchiesUsage',)

    def __init__(self,
                 name=None,
                 cacheId=None,
                 dataOnRows=False,
                 dataPosition=None,
                 dataCaption=None,
                 grandTotalCaption=None,
                 errorCaption=None,
                 showError=False,
                 missingCaption=None,
                 showMissing=True,
                 pageStyle=None,
                 pivotTableStyle=None,
                 vacatedStyle=None,
                 tag=None,
                 updatedVersion=0,
                 minRefreshableVersion=0,
                 asteriskTotals=False,
                 showItems=True,
                 editData=False,
                 disableFieldList=False,
                 showCalcMbrs=True,
                 visualTotals=True,
                 showMultipleLabel=True,
                 showDataDropDown=True,
                 showDrill=True,
                 printDrill=False,
                 showMemberPropertyTips=True,
                 showDataTips=True,
                 enableWizard=True,
                 enableDrill=True,
                 enableFieldProperties=True,
                 preserveFormatting=True,
                 useAutoFormatting=False,
                 pageWrap=0,
                 pageOverThenDown=False,
                 subtotalHiddenItems=False,
                 rowGrandTotals=True,
                 colGrandTotals=True,
                 fieldPrintTitles=False,
                 itemPrintTitles=False,
                 mergeItem=False,
                 showDropZones=True,
                 createdVersion=0,
                 indent=1,
                 showEmptyRow=False,
                 showEmptyCol=False,
                 showHeaders=None,
                 compact=True,
                 outline=False,
                 outlineData=False,
                 compactData=True,
                 published=False,
                 gridDropZones=False,
                 immersive=True,
                 multipleFieldFilters=None,
                 chartFormat=0,
                 rowHeaderCaption=None,
                 colHeaderCaption=None,
                 fieldListSortAscending=None,
                 mdxSubqueries=None,
                 customListSort=None,
                 autoFormatId=None,
                 applyNumberFormats=False,
                 applyBorderFormats=False,
                 applyFontFormats=False,
                 applyPatternFormats=False,
                 applyAlignmentFormats=False,
                 applyWidthHeightFormats=False,
                 location=None,
                 pivotFields=(),
                 rowFields=(),
                 rowItems=(),
                 colFields=(),
                 colItems=(),
                 pageFields=(),
                 dataFields=(),
                 formats=(),
                 conditionalFormats=(),
                 chartFormats=(),
                 pivotHierarchies=(),
                 pivotTableStyleInfo=None,
                 filters=(),
                 rowHierarchiesUsage=None,
                 colHierarchiesUsage=None,
                 extLst=None,
                 id=None,
                ):
        self.name = name
        self.cacheId = cacheId
        self.dataOnRows = dataOnRows
        self.dataPosition = dataPosition
        self.dataCaption = dataCaption
        self.grandTotalCaption = grandTotalCaption
        self.errorCaption = errorCaption
        self.showError = showError
        self.missingCaption = missingCaption
        self.showMissing = showMissing
        self.pageStyle = pageStyle
        self.pivotTableStyle = pivotTableStyle
        self.vacatedStyle = vacatedStyle
        self.tag = tag
        self.updatedVersion = updatedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.asteriskTotals = asteriskTotals
        self.showItems = showItems
        self.editData = editData
        self.disableFieldList = disableFieldList
        self.showCalcMbrs = showCalcMbrs
        self.visualTotals = visualTotals
        self.showMultipleLabel = showMultipleLabel
        self.showDataDropDown = showDataDropDown
        self.showDrill = showDrill
        self.printDrill = printDrill
        self.showMemberPropertyTips = showMemberPropertyTips
        self.showDataTips = showDataTips
        self.enableWizard = enableWizard
        self.enableDrill = enableDrill
        self.enableFieldProperties = enableFieldProperties
        self.preserveFormatting = preserveFormatting
        self.useAutoFormatting = useAutoFormatting
        self.pageWrap = pageWrap
        self.pageOverThenDown = pageOverThenDown
        self.subtotalHiddenItems = subtotalHiddenItems
        self.rowGrandTotals = rowGrandTotals
        self.colGrandTotals = colGrandTotals
        self.fieldPrintTitles = fieldPrintTitles
        self.itemPrintTitles = itemPrintTitles
        self.mergeItem = mergeItem
        self.showDropZones = showDropZones
        self.createdVersion = createdVersion
        self.indent = indent
        self.showEmptyRow = showEmptyRow
        self.showEmptyCol = showEmptyCol
        self.showHeaders = showHeaders
        self.compact = compact
        self.outline = outline
        self.outlineData = outlineData
        self.compactData = compactData
        self.published = published
        self.gridDropZones = gridDropZones
        self.immersive = immersive
        self.multipleFieldFilters = multipleFieldFilters
        self.chartFormat = chartFormat
        self.rowHeaderCaption = rowHeaderCaption
        self.colHeaderCaption = colHeaderCaption
        self.fieldListSortAscending = fieldListSortAscending
        self.mdxSubqueries = mdxSubqueries
        self.customListSort = customListSort
        self.autoFormatId = autoFormatId
        self.applyNumberFormats = applyNumberFormats
        self.applyBorderFormats = applyBorderFormats
        self.applyFontFormats = applyFontFormats
        self.applyPatternFormats = applyPatternFormats
        self.applyAlignmentFormats = applyAlignmentFormats
        self.applyWidthHeightFormats = applyWidthHeightFormats
        self.location = location
        self.pivotFields = pivotFields
        self.rowFields = rowFields
        self.rowItems = rowItems
        self.colFields = colFields
        self.colItems = colItems
        self.pageFields = pageFields
        self.dataFields = dataFields
        self.formats = formats
        self.conditionalFormats = conditionalFormats
        self.chartFormats = chartFormats
        self.pivotHierarchies = pivotHierarchies
        self.pivotTableStyleInfo = pivotTableStyleInfo
        self.filters = filters
        self.rowHierarchiesUsage = rowHierarchiesUsage
        self.colHierarchiesUsage = colHierarchiesUsage
        self.extLst = extLst
        self.id = id


    def to_tree(self):
        tree = super(PivotTableDefinition, self).to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Add to zipfile and update manifest
        """
        self._write_rels(archive, manifest)
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        """
        Write the relevant child objects and add links
        """
        if self.cache is None:
            return

        rels = RelationshipList()
        r = Relationship(Type=self.cache.rel_type)
        rels.append(r)
        self.id = r.id
        self.cache._id = self._id
        self.cache._write(archive, manifest)

        path = get_rels_path(self.path)
        xml = tostring(rels.to_tree())
        archive._writestr(path, xml)
