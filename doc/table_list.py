#Get table by nmae 
ws['Table1'] #Returns a table object with that name

ws.get(table_range="A1:E5") # Returns a table object with matching range
ws.get(name="Table1") #Return a table object with that name
