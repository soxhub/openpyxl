Defined Names
=============


The specification has the following to say about defined names:

    "Defined names are descriptive text that is used to represents a cell, range
    of cells, formula, or constant value."

This means they are very loosely defined. They might contain a constant, a
formula, a single cell reference, a range of cells or multiple ranges of
cells across different worksheets. Or all of the above. They are defined
globally for a workbook and accessed from there `defined_names` attribue.


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

    import openpyxl
    wb = openpyxl.Workbook()
    new_range = openpyxl.workbook.defined_name.DefinedName('newrange', attr_text='Sheet!$A$1:$A$5')
    wb.defined_names.append(new_range)

    # create a local named range (only valid for a specific sheet)
    sheetid = wb.sheetnames.index('Sheet')
    private_range = openpyxl.workbook.defined_name.DefinedName('privaterange', attr_text='Sheet!$A$6', localSheetId=sheetid)
    wb.defined_names.append(private_range)
    # this local range can't be retrieved from the global defined names
    assert('privaterange' not in wb.defined_names)

    # the scope has to be supplied to retrieve local ranges:
    print(wb.defined_names.localnames(sheetid))
    print(wb.defined_names.get('privaterange', sheetid).attr_text)

.. testoutput::

    ['privaterange']
    Sheet!$A$6
