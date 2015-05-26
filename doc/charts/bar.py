from openpyxl import Workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.title import title_maker

wb = Workbook(write_only=True)
ws = wb.create_sheet()

rows = [
    ('Number', 'Batch 1', 'Batch 2'),
    (2, 10, 30),
    (3, 40, 60),
    (4, 50, 70),
    (5, 20, 10),
    (6, 10, 40),
    (7, 50, 30),
]


for row in rows:
    ws.append(row)


chart1 = BarChart()
chart1.type = "bar"
chart1.style = 11
chart1.title = "Bar chart"
chart1.y_axis.title = title_maker('Test number')
chart1.x_axis.title = title_maker('Sample length (mm)')

data = Reference(ws, min_col=2, min_row=1, max_row=7, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=7)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.shape = 4
ws.add_chart(chart1, "D2")


chart2 = BarChart()
chart2.type = "bar"
chart2.style = 12
chart2.grouping = "stacked"
chart2.overlap = 100
chart2.title = 'Stacked Chart'
chart2.y_axis.title = title_maker('Test number')
chart2.x_axis.title = title_maker('Sample length (mm)')
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "D18")


chart3 = BarChart()
chart3.type = "bar"
chart3.style = 13
chart3.grouping = "percentStacked"
chart3.overlap = 100
chart3.title = 'Percent Stacked Chart'
chart3.y_axis.title = title_maker('Test number')
chart3.x_axis.title = title_maker('Sample length (mm)')
chart3.add_data(data, titles_from_data=True)
chart3.set_categories(cats)
ws.add_chart(chart3, "D34")


wb.save("bar.xlsx")
