Defined Names
=============


The specification has the following to say about defined names:

    "Defined names are descriptive text that is used to represents a cell, range
    of cells, formula, or constant value."

This means they are very loosely defined. They might contain a constant, a
formula, a single cell reference, a range of cells or multiple ranges of
cells across different worksheets. Or all of the above. Cell references or
ranges must use absolute coordinates and **always** include the name of the worksheet
they're in. Use the utilities :obj:`absolute_coordinate()` and :obj:`quote_sheetname()`
to do this.

Defined names can either be restricted to individual worksheets or available
globally for the whole workbook. Names must be unique within a collection; new
items will replace existing ones with the name.


Accessing Global Definitions
----------------------------

Global definitions are stored in the workbook collection::

    defn = wb.defined_names["my_range"]
    # the destinations attribute contains a list of ranges in the definitions
    dests = defn.destinations # returns a generator of (worksheet title, cell range) tuples

    cells = []
    for title, coord in dests:
        ws = wb[title]
        cells.append(ws[coord])


Accessing Worksheet Definitions
-------------------------------

Definitions are assigned to a specific worksheet are only accessible from
that worksheet::

    ws = wb["Sheet"]
    defn = ws.defined_names["private_range"]

Creating a Global Definition
----------------------------

Global definitions are assigned to the workbook collection::

    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.utils import quote_sheetname, absolute_coordinate
    wb = Workbook()
    ws = wb.active
    # make sure sheetnames and cell references are quoted correctly
    ref =  "{quote_sheetname(ws.title)}!{absolute_coordinate('A1:A5')}"

    defn = DefinedName("global_range", attr_text=ref)
    wb.defined_names["global_range"] = defn

    # key and `name` must be the same, the `.add()` method makes this easy
    wb.defined_names.add(new_range)


Creating a Worksheet Definition
-------------------------------

Definitions are assigned to a specific worksheet are only accessible from
that worksheet::

    # create a local named range (only valid for a specific sheet)
    ws = wb["Sheet"]
    ws.title = "My Sheet"
    # make sure sheetnames  and cell referencesare quoted correctly
    ref = f"{quote_sheetname(ws.title)}!{absolute_coordinate('A6')}"

    defn = DefinedName("private_range", attr_text=ref)
    ws.defined_names.add(defn)
    print(ws.defined_names["private_range"].attr_text)

