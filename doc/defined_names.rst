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
