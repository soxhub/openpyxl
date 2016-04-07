Working with Pandas and NumPy
=============================

openpyxl is able to work with the popular libraries `Pandas
<http://pandas.pydata.org>`_ and `NumPy <http://numpy.org>`_


NumPy Support
-------------

openpyxl has builtin support for the NumPy types float, integer and boolean.
DateTimes are supported using the Pandas' Timestamp type.


Working with Pandas Dataframes
------------------------------

The :func:`openpyxl.utils.dataframe.dataframe_to_rows` function provides a
simple way to work with Pandas Dataframes::

    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)


While Pandas itself supports conversion to Excel, this gives client code
additional flexibility including the ability to stream dataframes straight to
files.

To convert a dataframe into a worksheet highlighting the header and index::

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)

    # Create Pandas style formatting for headers and indices
    ft = Font(bold=True)
    al = Alignment(horizontal="center")
    side = Side(style="thin", color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)

    for c in ws['A'] + ws[1]:
        c.font = ft
        c.alignment = al
        c.border = border

    wb.save("pandas_openpyxl.xlsx")

Alternatively, if you just want to convert the data you can use write-only mode::

    from openpyxl.cell.cell import WriteOnlyCell
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    ft = Font(bold=True)
    al = Alignment(horizontal="center")
    side = Side(style="thin", color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)

    cell = WriteOnlyCell(ws)
    cell.font = ft
    cell.alignment = al
    cell.border = border

     def format_first_row(row, cell):

        for c in row:
            cell.value = c
            yield cell

    rows = dataframe_to_rows(df)
    first_row = format_first_row(next(rows), cell)
    ws.append(first_row)

    for row in rows:
        row = list(row)
        cell.value = row[0]
        row[0] = cell
        ws.append(row)

    wb.save("openpyxl_stream.xlsx")


This code will work just as well with a standard workbook.


Converting a worksheet to a Dataframe
-------------------------------------

To convert a worksheet to a Dataframe you can use the `values` property. This
is very easy if the worksheet has no headers or indices::

    df = DataFrame(ws.values)

If the worksheet does have headers or indices, such as one created by Pandas,
then a little more work is required::

    data = ws.values
    cols = next(data)[1:]
    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    df = DataFrame(data, index=idx, columns=cols)
