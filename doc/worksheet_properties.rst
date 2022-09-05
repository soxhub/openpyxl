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

Search `ECMA-376 pageSetup` for more details.

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

Worksheet Views
---------------

There are also several convenient properties defined as worksheet views. You can use :class:`ws.sheet_view<openpyxl.worksheet.views.SheetView>` to set sheet attributes such as zoom, show formulas or if the tab is selected.

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.sheet_view.zoom = 85 # Sets 85% zoom
>>> ws.sheet_view.showFormulas = True
>>> ws.sheet_view.tabSelected = True