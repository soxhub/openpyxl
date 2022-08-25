Defined Names
=============


The specification has the following to say about defined names:

    "Defined names are descriptive text that is used to represents a cell, range
    of cells, formula, or constant value."

This means they are very loosely defined. They might contain a constant, a
formula, a single cell reference, a range of cells or multiple ranges of
cells across different worksheets. Or all of the above. Cell references or
ranges should always include the name of the worksheet they're in.

Defined names can either be restricted to individual worksheets or available
globally for the whole workbook. Names must be unique within a collection; new
items will replace existing ones with the name.


Sample use for ranges
---------------------

Accessing a range called "my_range"::

    my_range = wb.defined_names['my_range']
    # if this contains a range of cells then the destinations attribute is not None
    dests = my_range.destinations # returns a generator of (worksheet title, cell range) tuples

    cells = []
    for title, coord in dests:
        ws = wb[title]
        cells.append(ws[coord])


Creating new named ranges
-------------------------

.. testcode::

    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.utils import quote_sheetname, absolute_coordinate
    wb = Workbook()
    new_range = DefinedName("newrange", attr_text="Sheet!$A$1:$A$5")
    wb.defined_names["newrange"] = new_range

    # key and name must be the same, the `.add()` method makes this easy
    wb.defined_names.add(new_range)

    # create a local named range (only valid for a specific sheet)
    ws = wb["Sheet"]
    ws.title = "My Sheet"
    # make sure sheetnames are quoted correctly
    ref = f"{quote_sheetname(ws.title)}!{absolute_coordinate('A6')}"
    private_range = DefinedName("private_range", attr_text=ref)
    ws.defined_names.add(private_range)
    print(ws.defined_names["private_range"].attr_text)

.. testoutput::

    'My Sheet'!$A$6
