Additional Worksheet Properties
===============================

These are advanced properties for particular behaviours, the most used ones
are the "fitTopage" page setup property and the tabColor that define the
background color of the worksheet tab.

Available properties for worksheets
-----------------------------------

* "enableFormatConditionsCalculation"
* "filterMode"
* "published"
* "syncHorizontal"
* "syncRef"
* "syncVertical"
* "transitionEvaluation"
* "transitionEntry"
* "tabColor"

Available fields for page setup properties
------------------------------------------

"autoPageBreaks"
"fitToPage"

Available fields for outlines
-----------------------------

* "applyStyles"
* "summaryBelow"
* "summaryRight"
* "showOutlineSymbols"

see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetproperties%28v=office.14%29.aspx_ for details.

.. note::
        By default, outline properties are intitialized so you can directly modify each of their 4 attributes, while page setup properties don't.
        If you want modify the latter, you should first initialize a :class:`openpyxl.worksheet.properties.PageSetupProperties` object with the required parameters.
        Once done, they can be directly modified by the routine later if needed.


.. :: doctest

>>> from openpyxl.workbook import Workbook
>>> from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> wsprops = ws.sheet_properties
>>> wsprops.tabColor = "1072BA"
>>> wsprops.filterMode = False
>>> wsprops.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
>>> wsprops.outlinePr.summaryBelow = False
>>> wsprops.outlinePr.applyStyles = True
>>> wsprops.pageSetUpPr.autoPageBreaks = True
