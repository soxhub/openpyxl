#Get table by nmae 
ws['Table1'] #Returns a table object with that name
wb.tables.get("Table1") #Return a table object with that name
wb.tables.delete("Table1") #Delete a table object with that name
wb.tables.items() #Returns a tuple with (table name, table ref)
