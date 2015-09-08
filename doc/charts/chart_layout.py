from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.layout import Layout, ManualLayout

wb = Workbook()
ws = wb.active

rows = [
    ['Size', 'Batch 1', 'Batch 2'],
    [2, 40, 30],
    [3, 40, 25],
    [4, 50, 30],
    [5, 30, 25],
    [6, 25, 35],
    [7, 20, 40],
]

for row in rows:
    ws.append(row)

ch = ScatterChart()
ch.scatterStyle = "marker"
ch.title = "Scatter Chart with Layout"
ch.style = 13
ch.x_axis.title = 'Size'
ch.y_axis.title = 'Percentage'

xvalues = Reference(ws, min_col=1, min_row=2, max_row=7)
for i in range(2, 4):
    values = Reference(ws, min_col=i, min_row=1, max_row=7)
    series = Series(values, xvalues, title_from_data=True)
    ch.series.append(series)
    
# Set layout of the legend    
ch.legend.layout=Layout(manualLayout=ManualLayout(yMode='edge', xMode='edge',x=0,y=0.9, h=0.1, w=1))
# x is the lateral position counted from the left in relative coordinates [0,1]
# y is the vertical position counted from the left in relative coordinates [0,1]
# h is the height of the box 
# w is the width of the box

# Set layout of the plot area
ch.plot_area.layout=Layout(manualLayout=ManualLayout(yMode='edge', xMode='edge',x=0.1,y=0.15, h=0.65, w=0.85))
# x is the lateral position counted from the left in relative coordinates [0,1]
# y is the vertical position counted from the left in relative coordinates [0,1]
# h is the height of the box 
# w is the width of the box

ws.add_chart(ch, "B10")

wb.save("chart_layout.xlsx")
