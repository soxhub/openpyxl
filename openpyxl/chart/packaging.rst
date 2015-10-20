Notes on packaging charts
=========================

In Excel charts are global objects even though they must be assigned to either a worksheet or a chart sheet. Assignment is further complicated because it is indirect: charts are referred to from drawings (also global objects), which are referred to from worksheets. Drawings can contain other objects. Relations are managed indirectly between worksheets and drawings, and drawings and charts.

The naive approach too packaging is to collect as groups (charts, images, etc.) and, thus calculate indices. This involves looping over sheets multiple times and is tightly coupled, involving the passing of relation ids around objects.

A better approach would be to serialise dependent objects sheet by sheet in reverse order of dependency: serialise charts (in /charts) and add the relation to the worksheet drawing relationship mapper; then serialise drawings (not sure if there can be more than one per worksheet, charts are normally bundled) and their relationships and add the relation to the drawing to the worksheet. All dependent objects must be either shadowed or referenced to both in the worksheet and in the workbook so that duplicate filenames can be easily avoided.