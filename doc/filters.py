from openpyxl import Workbook
from openpyxl.worksheet.filters import (
    FilterColumn,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
    Filters,
    )

wb = Workbook()
ws = wb.active

data = [
    ["Fruit", "Quantity"],
    ["Kiwi", 3],
    ["Grape", 15],
    ["Apple", 3],
    ["Peach", 3],
    ["Pomegranate", 3],
    ["Pear", 3],
    ["Tangerine", 3],
    ["Blueberry", 3],
    ["Mango", 3],
    ["Watermelon", 3],
    ["Blackberry", 3],
    ["Orange", 3],
    ["Raspberry", 3],
    ["Banana", 3]
]

for r in data:
    ws.append(r)

filters = ws.auto_filter
filters.ref = "A1:B15"
col = FilterColumn(colId=0) # for column A
col.filters = Filters(filter=["Kiwi", "Apple", "Mango"]) # add selected values
filters.filterColumn.append(col) # add filter to the worksheet

ws.auto_filter.add_sort_condition("B2:B15")

wb.save("filtered.xlsx")
