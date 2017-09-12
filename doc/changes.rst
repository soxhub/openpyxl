2.5.0-b1 (unreleased)
=====================


Bugfixes
--------
* `#876 <https://bitbucket.org/openpyxl/openpyxl/issues/876>`_ Unicode in chart axes not handled correctly in Python 2
* `#882 <https://bitbucket.org/openpyxl/openpyxl/issues/882>`_ ScatterCharts have defective axes
* `#885 <https://bitbucket.org/openpyxl/openpyxl/issues/885>`_ Charts with empty numVal elements cannot be read
* `#895 <https://bitbucket.org/openpyxl/openpyxl/issues/895>`_ Charts with PivotSource cannot be read


2.5.0-a3 (2017-08-14)
=====================


Bugfixes
--------
* `#848 <https://bitbucket.org/openpyxl/openpyxl/issues/848>`_ Reading workbooks with Pie Charts raises an exception
* `#857 <https://bitbucket.org/openpyxl/openpyxl/issues/857>`_ Pivot Tables without Worksheet Sources raise an exception


2.5.0-a2 (2017-06-25)
=====================


Major Changes
-------------

* Read support for charts


Bugfixes
--------
* `#833 <https://bitbucket.org/openpyxl/openpyxl/issues/833>`_ Cannot access chartsheets by title
* `#834 <https://bitbucket.org/openpyxl/openpyxl/issues/834>`_ Preserve workbook views
* `#841 <https://bitbucket.org/openpyxl/openpyxl/issues/841>`_ Incorrect classification of a datetime


2.5.0-a1 (2017-05-30)
=====================


Compatibility
-------------

* Dropped support for Python 2.6 and 3.3. openpyxl will not run with Python 2.6


Major Changes
-------------

* Read/write support for pivot tables


Deprecations
------------

* Dropped the anchor method from images and additional constructor arguments


Bugfixes
--------
* `#779 <https://bitbucket.org/openpyxl/openpyxl/issues/779>`_ Fails to recognise Chinese date format`
* `#828 <https://bitbucket.org/openpyxl/openpyxl/issues/828>`_ Include hidden cells in charts`


Pull requests
-------------
* `163 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/163>`_ Improved GradientFill


Minor changes
-------------

* Remove deprecated methods from Cell
* Remove deprecated methods from Worksheet
* Added read/write support for the datetime type for cells


2.4.9 (unreleased)
==================


Bugfixes
--------
* `#809 <https://bitbucket.org/openpyxl/openpyxl/issues/809>`_ Incomplete documentation of `copy_worksheet` method
* `#811 <https://bitbucket.org/openpyxl/openpyxl/issues/811>`_ Scoped definedNames not removed when worksheet is deleted
* `#824 <https://bitbucket.org/openpyxl/openpyxl/issues/824>`_ Raise an exception if a chart is used in multiple sheets
* `#842 <https://bitbucket.org/openpyxl/openpyxl/issues/842>`_ Non-ASCII table column headings cause an exception in Python 2
* `#846 <https://bitbucket.org/openpyxl/openpyxl/issues/846>`_ Conditional formats not supported in write-only mode
* `#849 <https://bitbucket.org/openpyxl/openpyxl/issues/849>`_ Conditional formats with no sqref cause an exception
* `#859 <https://bitbucket.org/openpyxl/openpyxl/issues/859>`_ Headers that start with a number conflict with font size


Pull requests
-------------
* `#179 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/179>`_ Cells kept in a set
* `#180 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/180>`_ Support for Workbook protection
* `#182 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/182>`_ Read support for page breaks
* `#183 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/183>`_ Improve documentation of `copy_worksheet` method


2.4.8 (2017-05-30)
==================


Bugfixes
--------
* AutoFilter.sortState being assignd to the ws.sortState
* `#766 <https://bitbucket.org/openpyxl/openpyxl/issues/666>`_ Sheetnames with apostrophes need additional escaping
* `#729 <https://bitbucket.org/openpyxl/openpyxl/issues/729>`_ Cannot open files created by Microsoft Dynamics
* `#819 <https://bitbucket.org/openpyxl/openpyxl/issues/819>`_ Negative percents not case correctly
* `#821 <https://bitbucket.org/openpyxl/openpyxl/issues/821>`_ Runtime imports can cause deadlock
* `#855 <https://bitbucket.org/openpyxl/openpyxl/issues/855>`_ Print area containing only columns leads to corrupt file


Minor changes
-------------
* Preserve any table styles


2.4.7 (2017-04-24)
==================


Bugfixes
--------
* `#807 <https://bitbucket.org/openpyxl/openpyxl/issues/807>`_ Sample files being included by mistake in sdist


2.4.6 (2017-04-14)
==================


Bugfixes
--------
* `#776 <https://bitbucket.org/openpyxl/openpyxl/issues/776>`_ Cannot apply formatting to plot area
* `#780 <https://bitbucket.org/openpyxl/openpyxl/issues/780>`_ Exception when element attributes are Python keywords
* `#781 <https://bitbucket.org/openpyxl/openpyxl/issues/781>`_ Exception raised when saving files with styled columns
* `#785 <https://bitbucket.org/openpyxl/openpyxl/issues/785>`_ Number formats for data labels are incorrect
* `#788 <https://bitbucket.org/openpyxl/openpyxl/issues/788>`_ Worksheet titles not quoted in defined names
* `#800 <https://bitbucket.org/openpyxl/openpyxl/issues/800>`_ Font underlines not read correctly


2.4.5 (2017-03-07)
==================


Bugfixes
--------
* `#750 <https://bitbucket.org/openpyxl/openpyxl/issues/750>`_ Adding images keeps file handles open
* `#772 <https://bitbucket.org/openpyxl/openpyxl/issues/772>`_ Exception for column-only ranges
* `#773 <https://bitbucket.org/openpyxl/openpyxl/issues/773>`_ Cannot copy worksheets with non-ascii titles on Python 2


Pull requests
-------------

* `161 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/161>`_ Support for non-standard names for Workbook part.
* `162 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/162>`_ Documentation correction


2.4.4 (2017-02-23)
==================


Bugfixes
--------

* `#673 <https://bitbucket.org/openpyxl/openpyxl/issues/673>`_ Add close method to workbooks
* `#762 <https://bitbucket.org/openpyxl/openpyxl/issues/762>`_ openpyxl can create files with invalid style indices
* `#729 <https://bitbucket.org/openpyxl/openpyxl/issues/729>`_ Allow images in write-only mode
* `#744 <https://bitbucket.org/openpyxl/openpyxl/issues/744>`_ Rounded corners for charts
* `#747 <https://bitbucket.org/openpyxl/openpyxl/issues/747>`_ Use repr when handling non-convertible objects
* `#764 <https://bitbucket.org/openpyxl/openpyxl/issues/764>`_ Hashing function is incorrect
* `#765 <https://bitbucket.org/openpyxl/openpyxl/issues/765>`_ Named styles share underlying array


Minor Changes
-------------

* Add roundtrip support for worksheet tables.


Pull requests
-------------

* `160 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/160>`_ Don't init mimetypes more than once.


2.4.3 (unreleased)
==================
bad release


2.4.2 (2017-01-31)
==================


Bug fixes
---------

* `#727 <https://bitbucket.org/openpyxl/openpyxl/issues/727>`_ DeprecationWarning is incorrect
* `#734 <https://bitbucket.org/openpyxl/openpyxl/issues/734>`_ Exception raised if userName is missing
* `#739 <https://bitbucket.org/openpyxl/openpyxl/issues/739>`_ Always provide a date1904 attribute
* `#740 <https://bitbucket.org/openpyxl/openpyxl/issues/740>`_ Hashes should be stored as Base64
* `#743 <https://bitbucket.org/openpyxl/openpyxl/issues/743>`_ Print titles broken on sheetnames with spaces
* `#748 <https://bitbucket.org/openpyxl/openpyxl/issues/748>`_ Workbook breaks when active sheet is removed
* `#754 <https://bitbucket.org/openpyxl/openpyxl/issues/754>`_ Incorrect descriptor for Filter values
* `#756 <https://bitbucket.org/openpyxl/openpyxl/issues/756>`_ Potential XXE vulerability
* `#758 <https://bitbucket.org/openpyxl/openpyxl/issues/758>`_ Cannot create files with page breaks and charts
* `#759 <https://bitbucket.org/openpyxl/openpyxl/issues/759>`_ Problems with worksheets with commas in their titles


Minor Changes
-------------

* Add unicode support for sheet name incrementation.


2.4.1 (2016-11-23)
==================


Bug fixes
---------

* `#643 <https://bitbucket.org/openpyxl/openpyxl/issues/643>`_ Make checking for duplicate sheet titles case insensitive
* `#647 <https://bitbucket.org/openpyxl/openpyxl/issues/647>`_ Trouble handling LibreOffice files with named styles
* `#687 <https://bitbucket.org/openpyxl/openpyxl/issues/682>`_ Directly assigned new named styles always refer to "Normal"
* `#690 <https://bitbucket.org/openpyxl/openpyxl/issues/690>`_ Cannot parse print titles with multiple sheet names
* `#691 <https://bitbucket.org/openpyxl/openpyxl/issues/691>`_ Cannot work with macro files created by LibreOffice
* Prevent duplicate differential styles
* `#694 <https://bitbucket.org/openpyxl/openpyxl/issues/694>`_ Allow sheet titles longer than 31 characters
* `#697 <https://bitbucket.org/openpyxl/openpyxl/issues/697>`_ Cannot unset hyperlinks
* `#699 <https://bitbucket.org/openpyxl/openpyxl/issues/699>`_ Exception raised when format objects use cell references
* `#703 <https://bitbucket.org/openpyxl/openpyxl/issues/703>`_ Copy height and width when copying comments
* `#705 <https://bitbucket.org/openpyxl/openpyxl/issues/705>`_ Incorrect content type for VBA macros
* `#707 <https://bitbucket.org/openpyxl/openpyxl/issues/707>`_ IndexError raised in read-only mode when accessing individual cells
* `#711 <https://bitbucket.org/openpyxl/openpyxl/issues/711>`_ Files with external links become corrupted
* `#715 <https://bitbucket.org/openpyxl/openpyxl/issues/715>`_ Cannot read files containing macro sheets
* `#717 <https://bitbucket.org/openpyxl/openpyxl/issues/717>`_ Details from named styles not preserved when reading files
* `#722 <https://bitbucket.org/openpyxl/openpyxl/issues/722>`_ Remove broken Print Title and Print Area definitions


Minor changes
-------------

* Add support for Python 3.6
* Correct documentation for headers and footers


Deprecations
------------

Worksheet methods `get_named_range()` and `get_sqaured_range()`


Bug fixes
---------


2.4.0 (2016-09-15)
==================


Bug fixes
---------

* `#652 <https://bitbucket.org/openpyxl/openpyxl/issues/652>`_ Exception raised when epoch is 1904
* `#642 <https://bitbucket.org/openpyxl/openpyxl/issues/642>`_ Cannot handle unicode in headers and footers in Python 2
* `#646 <https://bitbucket.org/openpyxl/openpyxl/issues/646>`_ Cannot handle unicode sheetnames in Python 2
* `#658 <https://bitbucket.org/openpyxl/openpyxl/issues/658>`_ Chart styles, and axis units should not be 0
* `#663 <https://bitbucket.org/openpyxl/openpyxl/issues/663>`_ Strings in external workbooks not unicode


Major changes
-------------

* Add support for builtin styles and include one for Pandas


Minor changes
-------------

* Add a `keep_links` option to `load_workbook`. External links contain cached
  copies of the external workbooks. If these are big it can be advantageous to
  be able to disable them.
* Provide an example for using cell ranges in DataValidation.
* PR 138 - add copy support to comments.


2.4.0-b1 (2016-06-08)
=====================


Minor changes
-------------

* Add an the alias `hide_drop_down` to DataValidation for `showDropDown` because that is how Excel works.


Bug fixes
---------

* `#625 <https://bitbucket.org/openpyxl/openpyxl/issues/625>`_ Exception raises when inspecting EmptyCells in read-only mode
* `#547 <https://bitbucket.org/openpyxl/openpyxl/issues/547>`_ Functions for handling OOXML "escaped" ST_XStrings
* `#629 <https://bitbucket.org/openpyxl/openpyxl/issues/629>`_ Row Dimensions not supported in write-only mode
* `#530 <https://bitbucket.org/openpyxl/openpyxl/issues/530>`_ Problems when removing worksheets with charts
* `#630 <https://bitbucket.org/openpyxl/openpyxl/issues/630>`_ Cannot use SheetProtection in write-only mode


Features
--------

* Add write support for worksheet tables


2.4.0-a1 (2016-04-11)
=====================


Minor changes
-------------

* Remove deprecated methods from DataValidation
* Remove deprecated methods from PrintPageSetup
* Convert AutoFilter to Serialisable and extend support for filters
* Add support for SortState
* Removed `use_iterators` keyword when loading workbooks. Use `read_only` instead.
* Removed `optimized_write` keyword for new workbooks. Use `write_only` instead.
* Improve print title support
* Add print area support
* New implementation of defined names
* New implementation of page headers and footers
* Add support for Python's NaN
* Added iter_cols method for worksheets
* ws.rows and ws.columns now always return generators and start at the top of the worksheet
* Add a `values` property for worksheets
* Default column width changed to 8 as per the specification


Deprecations
------------

* Cell anchor method
* Worksheet point_pos method
* Worksheet add_print_title method
* Worksheet HeaderFooter attribute, replaced by individual ones
* Flatten function for cells
* Workbook get_named_range, add_named_range, remove_named_range, get_sheet_names, get_sheet_by_name
* Comment text attribute
* Use of range strings deprecated for ws.iter_rows()
* Use of coordinates deprecated for ws.cell()
* Deprecate .copy() method for StyleProxy objects


Bug fixes
---------

* `#152 <https://bitbucket.org/openpyxl/openpyxl/issues/152>`_ Hyperlinks lost when reading files
* `#171 <https://bitbucket.org/openpyxl/openpyxl/issues/171>`_ Add function for copying worksheets
* `#386 <https://bitbucket.org/openpyxl/openpyxl/issues/386>`_ Cells with inline strings considered empty
* `#397 <https://bitbucket.org/openpyxl/openpyxl/issues/397>`_ Add support for ranges of rows and columns
* `#446 <https://bitbucket.org/openpyxl/openpyxl/issues/446>`_ Workbook with definedNames corrupted by openpyxl
* `#481 <https://bitbucket.org/openpyxl/openpyxl/issues/481>`_ "safe" reserved ranges are not read from workbooks
* `#501 <https://bitbucket.org/openpyxl/openpyxl/issues/501>`_ Discarding named ranges can lead to corrupt files
* `#574 <https://bitbucket.org/openpyxl/openpyxl/issues/574>`_ Exception raised when using the class method to parse Relationships
* `#579 <https://bitbucket.org/openpyxl/openpyxl/issues/579>`_ Crashes when reading defined names with no content
* `#597 <https://bitbucket.org/openpyxl/openpyxl/issues/597>`_ Cannot read worksheets without coordinates
* `#617 <https://bitbucket.org/openpyxl/openpyxl/issues/617>`_ Customised named styles not correctly preserved


2.3.5 (2016-04-11)
==================


Bug fixes
---------

* `#618 <https://bitbucket.org/openpyxl/openpyxl/issues/618>`_ Comments not written in write-only mode


2.3.4 (2016-03-16)
==================


Bug fixes
---------

* `#594 <https://bitbucket.org/openpyxl/openpyxl/issues/594>`_ Content types might be missing when keeping VBA
* `#599 <https://bitbucket.org/openpyxl/openpyxl/issues/599>`_ Cells with only one cell look empty
* `#607 <https://bitbucket.org/openpyxl/openpyxl/issues/607>`_ Serialise NaN as ''


Minor changes
-------------

* Preserve the order of external references because formualae use numerical indices.
* Typo corrected in cell unit tests (PR 118)


2.3.3 (2016-01-18)
==================


Bug fixes
---------

* `#540 <https://bitbucket.org/openpyxl/openpyxl/issues/540>`_ Cannot read merged cells in read-only mode
* `#565 <https://bitbucket.org/openpyxl/openpyxl/issues/565>`_ Empty styled text blocks cannot be parsed
* `#569 <https://bitbucket.org/openpyxl/openpyxl/issues/569>`_ Issue warning rather than raise Exception raised for unparsable definedNames
* `#575 <https://bitbucket.org/openpyxl/openpyxl/issues/575>`_ Cannot open workbooks with embdedded OLE files
* `#584 <https://bitbucket.org/openpyxl/openpyxl/issues/584>`_ Exception when saving borders with attributes


Minor changes
-------------

* `PR 103 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/103/>`_ Documentation about chart scaling and axis limits
* Raise an exception when trying to copy cells from other workbooks.


2.3.2 (2015-12-07)
==================


Bug fixes
---------

* `#554 <https://bitbucket.org/openpyxl/openpyxl/issues/554>`_ Cannot add comments to a worksheet when preserving VBA
* `#561 <https://bitbucket.org/openpyxl/openpyxl/issues/561>`_ Exception when reading phonetic text
* `#562 <https://bitbucket.org/openpyxl/openpyxl/issues/562>`_ DARKBLUE is the same as RED
* `#563 <https://bitbucket.org/openpyxl/openpyxl/issues/563>`_ Minimum for row and column indexes not enforced


Minor changes
-------------

* `PR 97 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/97/>`_ One VML file per worksheet.
* `PR 96 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/96/>`_ Correct descriptor for CharacterProperties.rtl
* `#498 <https://bitbucket.org/openpyxl/openpyxl/issues/498>`_ Metadata is not essential to use the package.


2.3.1 (2015-11-20)
==================


Bug fixes
---------

* `#534 <https://bitbucket.org/openpyxl/openpyxl/issues/534>`_ Exception when using columns property in read-only mode.
* `#536 <https://bitbucket.org/openpyxl/openpyxl/issues/536>`_ Incorrectly handle comments from Google Docs files.
* `#539 <https://bitbucket.org/openpyxl/openpyxl/issues/539>`_ Flexible value types for conditional formatting.
* `#542 <https://bitbucket.org/openpyxl/openpyxl/issues/542>`_ Missing content types for images.
* `#543 <https://bitbucket.org/openpyxl/openpyxl/issues/543>`_ Make sure images fit containers on all OSes.
* `#544 <https://bitbucket.org/openpyxl/openpyxl/issues/544>`_ Gracefully handle missing cell styles.
* `#546 <https://bitbucket.org/openpyxl/openpyxl/issues/546>`_ ExternalLink duplicated when editing a file with macros.
* `#548 <https://bitbucket.org/openpyxl/openpyxl/issues/548>`_ Exception with non-ASCII worksheet titles
* `#551 <https://bitbucket.org/openpyxl/openpyxl/issues/551>`_ Combine multiple LineCharts


Minor changes
-------------

* `PR 88 <https://bitbucket.org/openpyxl/openpyxl/pull-requests/88/>`_ Fix page margins in parser.


2.3.0 (2015-10-20)
==================


Major changes
-------------

* Support the creation of chartsheets


Bug fixes
---------

* `#532 <https://bitbucket.org/openpyxl/openpyxl/issues/532>`_ Problems when cells have no style in read-only mode.


Minor changes
-------------

* PR 79 Make PlotArea editable in charts
* Use graphicalProperties as the alias for spPr


2.3.0-b2 (2015-09-04)
=====================


Bug fixes
---------

* `#488 <https://bitbucket.org/openpyxl/openpyxl/issue/488>`_ Support hashValue attribute for sheetProtection
* `#493 <https://bitbucket.org/openpyxl/openpyxl/issue/493>`_ Warn that unsupported extensions will be dropped
* `#494 <https://bitbucket.org/openpyxl/openpyxl/issues/494/>`_ Cells with exponentials causes a ValueError
* `#497 <https://bitbucket.org/openpyxl/openpyxl/issues/497/>`_ Scatter charts are broken
* `#499 <https://bitbucket.org/openpyxl/openpyxl/issues/499/>`_ Inconsistent conversion of localised datetimes
* `#500 <https://bitbucket.org/openpyxl/openpyxl/issues/500/>`_ Adding images leads to unreadable files
* `#509 <https://bitbucket.org/openpyxl/openpyxl/issues/509/>`_ Improve handling of sheet names
* `#515 <https://bitbucket.org/openpyxl/openpyxl/issues/515/>`_ Non-ascii titles have bad repr
* `#516 <https://bitbucket.org/openpyxl/openpyxl/issues/516/>`_ Ignore unassigned worksheets


Minor changes
-------------

* Worksheets are now iterable by row.
* Assign individual cell styles only if they are explicitly set.


2.3.0-b1 (2015-06-29)
=====================


Major changes
-------------

* Shift to using (row, column) indexing for cells. Cells will at some point *lose* coordinates.
* New implementation of conditional formatting. Databars now partially preserved.
* et_xmlfile is now a standalone library.
* Complete rewrite of chart package
* Include a tokenizer for fomulae to be able to adjust cell references in them. PR 63


Minor changes
-------------

* Read-only and write-only worksheets renamed.
* Write-only workbooks support charts and images.
* `PR76 <https://bitbucket.org/openpyxl/openpyxl/pull-request/76>`_ Prevent comment images from conflicting with VBA


Bug fixes
---------

* `#81 <https://bitbucket.org/openpyxl/openpyxl/issue/81>`_ Support stacked bar charts
* `#88 <https://bitbucket.org/openpyxl/openpyxl/issue/88>`_ Charts break hyperlinks
* `#97 <https://bitbucket.org/openpyxl/openpyxl/issue/97>`_ Pie and combination charts
* `#99 <https://bitbucket.org/openpyxl/openpyxl/issue/99>`_ Quote worksheet names in chart references
* `#150 <https://bitbucket.org/openpyxl/openpyxl/issue/150>`_ Support additional chart options
* `#172 <https://bitbucket.org/openpyxl/openpyxl/issue/172>`_ Support surface charts
* `#381 <https://bitbucket.org/openpyxl/openpyxl/issue/381>`_ Preserve named styles
* `#470 <https://bitbucket.org/openpyxl/openpyxl/issue/470>`_ Adding more than 10 worksheets with the same name leads to duplicates sheet names and an invalid file


2.2.6 (unreleased)
==================


Bug fixes
---------

* `#502 <https://bitbucket.org/openpyxl/openpyxl/issue/502>`_ Unexpected keyword "mergeCell"
* `#503 <https://bitbucket.org/openpyxl/openpyxl/issue/503>`_ tostring missing in dump_worksheet
* `#506 <https://bitbucket.org/openpyxl/openpyxl/issues/506>`_ Non-ASCII formulae cannot be parsed
* `#508 <https://bitbucket.org/openpyxl/openpyxl/issues/508>`_ Cannot save files with coloured tabs
* Regex for ignoring named ranges is wrong (character class instead of prefix)


2.2.5 (2015-06-29)
==================


Bug fixes
---------

* `#463 <https://bitbucket.org/openpyxl/openpyxl/issue/463>`_ Unexpected keyword "mergeCell"
* `#484 <https://bitbucket.org/openpyxl/openpyxl/issue/484>`_ Unusual dimensions breaks read-only mode
* `#485 <https://bitbucket.org/openpyxl/openpyxl/issue/485>`_ Move return out of loop


2.2.4 (2015-06-17)
==================


Bug fixes
---------

* `#464 <https://bitbucket.org/openpyxl/openpyxl/issue/464>`_ Cannot use images when preserving macros
* `#465 <https://bitbucket.org/openpyxl/openpyxl/issue/465>`_ ws.cell() returns an empty cell on read-only workbooks
* `#467 <https://bitbucket.org/openpyxl/openpyxl/issue/467>`_ Cannot edit a file with ActiveX components
* `#471 <https://bitbucket.org/openpyxl/openpyxl/issue/471>`_ Sheet properties elements must be in order
* `#475 <https://bitbucket.org/openpyxl/openpyxl/issue/475>`_ Do not redefine class __slots__ in subclasses
* `#477 <https://bitbucket.org/openpyxl/openpyxl/issue/477>`_ Write-only support for SheetProtection
* `#478 <https://bitbucket.org/openpyxl/openpyxl/issue/477>`_ Write-only support for DataValidation
* Improved regex when checking for datetime formats


2.2.3 (2015-05-26)
==================


Bug fixes
---------

* `#451 <https://bitbucket.org/openpyxl/openpyxl/issue/451>`_ fitToPage setting ignored
* `#458 <https://bitbucket.org/openpyxl/openpyxl/issue/458>`_ Trailing spaces lost when saving files.
* `#459 <https://bitbucket.org/openpyxl/openpyxl/issue/459>`_ setup.py install fails with Python 3
* `#462 <https://bitbucket.org/openpyxl/openpyxl/issue/462>`_ Vestigial rId conflicts when adding charts, images or comments
* `#455 <https://bitbucket.org/openpyxl/openpyxl/issue/455>`_ Enable Zip64 extensions for all versions of Python


2.2.2 (2015-04-28)
==================


Bug fixes
---------

* `#447 <https://bitbucket.org/openpyxl/openpyxl/issue/447>`_ Uppercase datetime number formats not recognised.
* `#453 <https://bitbucket.org/openpyxl/openpyxl/issue/453>`_ Borders broken in shared_styles.


2.2.1 (2015-03-31)
==================


Minor changes
-------------

* `PR54 <https://bitbucket.org/openpyxl/openpyxl/pull-request/54>`_ Improved precision on times near midnight.
* `PR55 <https://bitbucket.org/openpyxl/openpyxl/pull-request/55>`_ Preserve macro buttons


Bug fixes
---------

* `#429 <https://bitbucket.org/openpyxl/openpyxl/issue/429>`_ Workbook fails to load because header and footers cannot be parsed.
* `#433 <https://bitbucket.org/openpyxl/openpyxl/issue/433>`_ File-like object with encoding=None
* `#434 <https://bitbucket.org/openpyxl/openpyxl/issue/434>`_ SyntaxError when writing page breaks.
* `#436 <https://bitbucket.org/openpyxl/openpyxl/issue/436>`_ Read-only mode duplicates empty rows.
* `#437 <https://bitbucket.org/openpyxl/openpyxl/issue/437>`_ Cell.offset raises an exception
* `#438 <https://bitbucket.org/openpyxl/openpyxl/issue/438>`_ Cells with pivotButton and quotePrefix styles cannot be read
* `#440 <https://bitbucket.org/openpyxl/openpyxl/issue/440>`_ Error when customised versions of builtin formats
* `#442 <https://bitbucket.org/openpyxl/openpyxl/issue/442>`_ Exception raised when a fill element contains no children
* `#444 <https://bitbucket.org/openpyxl/openpyxl/issue/442>`_ Styles cannot be copied


2.2.0 (2015-03-11)
==================


Bug fixes
---------
* `#415 <https://bitbucket.org/openpyxl/openpyxl/issue/415>`_ Improved exception when passing in invalid in memory files.


2.2.0-b1 (2015-02-18)
=====================


Major changes
-------------
* Cell styles deprecated, use formatting objects (fonts, fills, borders, etc.) directly instead
* Charts will no longer try and calculate axes by default
* Support for template file types - PR21
* Moved ancillary functions and classes into utils package - single place of reference
* `PR 34 <https://bitbucket.org/openpyxl/openpyxl/pull-request/34/>`_ Fully support page setup
* Removed SAX-based XML Generator. Special thanks to Elias Rabel for implementing xmlfile for xml.etree
* Preserve sheet view definitions in existing files (frozen panes, zoom, etc.)


Bug fixes
---------
* `#103 <https://bitbucket.org/openpyxl/openpyxl/issue/103>`_ Set the zoom of a sheet
* `#199 <https://bitbucket.org/openpyxl/openpyxl/issue/199>`_ Hide gridlines
* `#215 <https://bitbucket.org/openpyxl/openpyxl/issue/215>`_ Preserve sheet view setings
* `#262 <https://bitbucket.org/openpyxl/openpyxl/issue/262>`_ Set the zoom of a sheet
* `#392 <https://bitbucket.org/openpyxl/openpyxl/issue/392>`_ Worksheet header not read
* `#387 <https://bitbucket.org/openpyxl/openpyxl/issue/387>`_ Cannot read files without styles.xml
* `#410 <https://bitbucket.org/openpyxl/openpyxl/issue/410>`_ Exception when preserving whitespace in strings
* `#417 <https://bitbucket.org/openpyxl/openpyxl/issue/417>`_ Cannot create print titles
* `#420 <https://bitbucket.org/openpyxl/openpyxl/issue/420>`_ Rename confusing constants
* `#422 <https://bitbucket.org/openpyxl/openpyxl/issue/422>`_ Preserve color index in a workbook if it differs from the standard


Minor changes
-------------
* Use a 2-way cache for column index lookups
* Clean up tests in cells
* `PR 40 <https://bitbucket.org/openpyxl/openpyxl/pull-request/40/>`_ Support frozen panes and autofilter in write-only mode
* Use ws.calculate_dimension(force=True) in read-only mode for unsized worksheets


2.1.5 (2015-02-18)
==================


Bug fixes
---------
* `#403 <https://bitbucket.org/openpyxl/openpyxl/issue/403>`_ Cannot add comments in write-only mode
* `#401 <https://bitbucket.org/openpyxl/openpyxl/issue/401>`_ Creating cells in an empty row raises an exception
* `#408 <https://bitbucket.org/openpyxl/openpyxl/issue/408>`_ from_excel adjustment for Julian dates 1 < x < 60
* `#409 <https://bitbucket.org/openpyxl/openpyxl/issue/409>`_ refersTo is an optional attribute


Minor changes
-------------
* Allow cells to be appended to standard worksheets for code compatibility with write-only mode.


2.1.4 (2014-12-16)
==================


Bug fixes
---------

* `#393 <https://bitbucket.org/openpyxl/openpyxl/issue/393>`_ IterableWorksheet skips empty cells in rows
* `#394 <https://bitbucket.org/openpyxl/openpyxl/issue/394>`_ Date format is applied to all columns (while only first column contains dates)
* `#395 <https://bitbucket.org/openpyxl/openpyxl/issue/395>`_ temporary files not cleaned properly
* `#396 <https://bitbucket.org/openpyxl/openpyxl/issue/396>`_ Cannot write "=" in Excel file
* `#398 <https://bitbucket.org/openpyxl/openpyxl/issue/398>`_ Cannot write empty rows in write-only mode with LXML installed


Minor changes
-------------
* Add relation namespace to root element for compatibility with iWork
* Serialize comments relation in LXML-backend


2.1.3 (2014-12-09)
==================


Minor changes
-------------
* `PR 31 <https://bitbucket.org/openpyxl/openpyxl/pull-request/31/>`_ Correct tutorial
* `PR 32 <https://bitbucket.org/openpyxl/openpyxl/pull-request/32/>`_ See #380
* `PR 37 <https://bitbucket.org/openpyxl/openpyxl/pull-request/37/>`_ Bind worksheet to ColumnDimension objects


Bug fixes
---------
* `#379 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ ws.append() doesn't set RowDimension Correctly
* `#380 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ empty cells formatted as datetimes raise exceptions


2.1.2 (2014-10-23)
==================


Minor changes
-------------
* `PR 30 <https://bitbucket.org/openpyxl/openpyxl/pull-request/30/>`_ Fix regex for positive exponentials
* `PR 28 <https://bitbucket.org/openpyxl/openpyxl/pull-request/28/>`_ Fix for #328


Bug fixes
---------
* `#120 <https://bitbucket.org/openpyxl/openpyxl/issue/120>`_, `#168 <https://bitbucket.org/openpyxl/openpyxl/issue/168>`_ defined names with formulae raise exceptions, `#292 <https://bitbucket.org/openpyxl/openpyxl/issue/292>`_
* `#328 <https://bitbucket.org/openpyxl/openpyxl/issue/328/>`_ ValueError when reading cells with hyperlinks
* `#369 <https://bitbucket.org/openpyxl/openpyxl/issue/369>`_ IndexError when reading definedNames
* `#372 <https://bitbucket.org/openpyxl/openpyxl/issue/372>`_ number_format not consistently applied from styles


2.1.1 (2014-10-08)
==================


Minor changes
-------------
* PR 20 Support different workbook code names
* Allow auto_axis keyword for ScatterCharts


Bug fixes
---------

* `#332 <https://bitbucket.org/openpyxl/openpyxl/issue/332>`_ Fills lost in ConditionalFormatting
* `#360 <https://bitbucket.org/openpyxl/openpyxl/issue/360>`_ Support value="none" in attributes
* `#363 <https://bitbucket.org/openpyxl/openpyxl/issue/363>`_ Support undocumented value for textRotation
* `#364 <https://bitbucket.org/openpyxl/openpyxl/issue/364>`_ Preserve integers in read-only mode
* `#366 <https://bitbucket.org/openpyxl/openpyxl/issue/366>`_ Complete read support for DataValidation
* `#367 <https://bitbucket.org/openpyxl/openpyxl/issue/367>`_ Iterate over unsized worksheets


2.1.0 (2014-09-21)
==================

Major changes
-------------
* "read_only" and "write_only" new flags for workbooks
* Support for reading and writing worksheet protection
* Support for reading hidden rows
* Cells now manage their styles directly
* ColumnDimension and RowDimension object manage their styles directly
* Use xmlfile for writing worksheets if available - around 3 times faster
* Datavalidation now part of the worksheet package


Minor changes
-------------
* Number formats are now just strings
* Strings can be used for RGB and aRGB colours for Fonts, Fills and Borders
* Create all style tags in a single pass
* Performance improvement when appending rows
* Cleaner conversion of Python to Excel values
* PR6 reserve formatting for empty rows
* standard worksheets can append from ranges and generators


Bug fixes
---------
* `#153 <https://bitbucket.org/openpyxl/openpyxl/issue/153>`_ Cannot read visibility of sheets and rows
* `#181 <https://bitbucket.org/openpyxl/openpyxl/issue/181>`_ No content type for worksheets
* `241 <https://bitbucket.org/openpyxl/openpyxl/issue/241>`_ Cannot read sheets with inline strings
* `322 <https://bitbucket.org/openpyxl/openpyxl/issue/322>`_ 1-indexing for merged cells
* `339 <https://bitbucket.org/openpyxl/openpyxl/issue/339>`_ Correctly handle removal of cell protection
* `341 <https://bitbucket.org/openpyxl/openpyxl/issue/341>`_ Cells with formulae do not round-trip
* `347 <https://bitbucket.org/openpyxl/openpyxl/issue/347>`_ Read DataValidations
* `353 <https://bitbucket.org/openpyxl/openpyxl/issue/353>`_ Support Defined Named Ranges to external workbooks


2.0.5 (2014-08-08)
==================


Bug fixes
---------
* `#348 <https://bitbucket.org/openpyxl/openpyxl/issue/348>`_ incorrect casting of boolean strings
* `#349 <https://bitbucket.org/openpyxl/openpyxl/issue/349>`_ roundtripping cells with formulae


2.0.4 (2014-06-25)
==================

Minor changes
-------------
* Add a sample file illustrating colours


Bug fixes
---------

* `#331 <https://bitbucket.org/openpyxl/openpyxl/issue/331>`_ DARKYELLOW was incorrect
* Correctly handle extend attribute for fonts


2.0.3 (2014-05-22)
==================

Minor changes
-------------

* Updated docs


Bug fixes
---------

* `#319 <https://bitbucket.org/openpyxl/openpyxl/issue/319>`_ Cannot load Workbooks with vertAlign styling for fonts


2.0.2 (2014-05-13)
==================

2.0.1 (2014-05-13)  brown bag
=============================

2.0.0 (2014-05-13)  brown bag
=============================


Major changes
-------------

* This is last release that will support Python 3.2
* Cells are referenced with 1-indexing: A1 == cell(row=1, column=1)
* Use jdcal for more efficient and reliable conversion of datetimes
* Significant speed up when reading files
* Merged immutable styles
* Type inference is disabled by default
* RawCell renamed ReadOnlyCell
* ReadOnlyCell.internal_value and ReadOnlyCell.value now behave the same as Cell
* Provide no size information on unsized worksheets
* Lower memory footprint when reading files


Minor changes
-------------

* All tests converted to pytest
* Pyflakes used for static code analysis
* Sample code in the documentation is automatically run
* Support GradientFills
* BaseColWidth set


Pull requests
-------------
* #70 Add filterColumn, sortCondition support to AutoFilter
* #80 Reorder worksheets parts
* #82 Update API for conditional formatting
* #87 Add support for writing Protection styles, others
* #89 Better handling of content types when preserving macros


Bug fixes
---------
* `#46 <https://bitbucket.org/openpyxl/openpyxl/issue/46>`_ ColumnDimension style error
* `#86 <https://bitbucket.org/openpyxl/openpyxl/issue/86>`_ reader.worksheet.fast_parse sets booleans to integers
* `#98 <https://bitbucket.org/openpyxl/openpyxl/issue/98>`_ Auto sizing column widths does not work
* `#137 <https://bitbucket.org/openpyxl/openpyxl/issue/137>`_ Workbooks with chartsheets
* `#185 <https://bitbucket.org/openpyxl/openpyxl/issue/185>`_  Invalid PageMargins
* `#230 <https://bitbucket.org/openpyxl/openpyxl/issue/230>`_ Using \v in cells creates invalid files
* `#243 <https://bitbucket.org/openpyxl/openpyxl/issue/243>`_ - IndexError when loading workbook
* `#263 <https://bitbucket.org/openpyxl/openpyxl/issue/263>`_ - Forded conversion of line breaks
* `#267 <https://bitbucket.org/openpyxl/openpyxl/issue/267>`_ - Raise exceptions when passed invalid types
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Cannot open files which use non-standard sheet names or reference Ids
* `#269 <https://bitbucket.org/openpyxl/openpyxl/issue/269>`_ - Handling unsized worksheets in IterableWorksheet
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Handling Workbooks with non-standard references
* `#275 <https://bitbucket.org/openpyxl/openpyxl/issue/275>`_ - Handling auto filters where there are only custom filters
* `#277 <https://bitbucket.org/openpyxl/openpyxl/issue/277>`_ - Harmonise chart and cell coordinates
* `#280 <https://bitbucket.org/openpyxl/openpyxl/issue/280>`_- Explicit exception raising for invalid characters
* `#286 <https://bitbucket.org/openpyxl/openpyxl/issue/286>`_ - Optimized writer can not handle a datetime.time value
* `#296 <https://bitbucket.org/openpyxl/openpyxl/issue/296>`_ - Cell coordinates not consistent with documentation
* `#300 <https://bitbucket.org/openpyxl/openpyxl/issue/300>`_ - Missing column width causes load_workbook() exception
* `#304 <https://bitbucket.org/openpyxl/openpyxl/issue/304>`_ - Handling Workbooks with absolute paths for worksheets (from Sharepoint)


1.8.6 (2014-05-05)
==================

Minor changes
-------------
Fixed typo for import Elementtree

Bugfixes
--------
* `#279 <https://bitbucket.org/openpyxl/openpyxl/issue/279>`_ Incorrect path for comments files on Windows


1.8.5 (2014-03-25)
==================

Minor changes
-------------
* The '=' string is no longer interpreted as a formula
* When a client writes empty xml tags for cells (e.g. <c r='A1'></c>), reader will not crash


1.8.4 (2014-02-25)
==================

Bugfixes
--------
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ better handling of undimensioned worksheets
* `#268 <https://bitbucket.org/openpyxl/openpyxl/issue/268>`_ non-ascii in formualae
* `#282 <https://bitbucket.org/openpyxl/openpyxl/issue/282>`_ correct implementation of register_namepsace for Python 2.6


1.8.3 (2014-02-09)
==================

Major changes
-------------
Always parse using cElementTree

Minor changes
-------------
Slight improvements in memory use when parsing

* `#256 <https://bitbucket.org/openpyxl/openpyxl/issue/256>`_ - error when trying to read comments with optimised reader
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ - unsized worksheets
* `#264 <https://bitbucket.org/openpyxl/openpyxl/issue/264>`_ - only numeric cells can be dates


1.8.2 (2014-01-17)
==================

* `#247 <https://bitbucket.org/openpyxl/openpyxl/issue/247>`_ - iterable worksheets open too many files
* `#252 <https://bitbucket.org/openpyxl/openpyxl/issue/252>`_ - improved handling of lxml
* `#253 <https://bitbucket.org/openpyxl/openpyxl/issue/253>`_ - better handling of unique sheetnames


1.8.1 (2014-01-14)
==================

* `#246 <https://bitbucket.org/openpyxl/openpyxl/issue/246>`_


1.8.0 (2014-01-08)
==================

Compatibility
-------------

Support for Python 2.5 dropped.

Major changes
-------------

* Support conditional formatting
* Support lxml as backend
* Support reading and writing comments
* pytest as testrunner now required
* Improvements in charts: new types, more reliable


Minor changes
-------------

* load_workbook now accepts data_only to allow extracting values only from
  formulae. Default is false.
* Images can now be anchored to cells
* Docs updated
* Provisional benchmarking
* Added convenience methods for accessing worksheets and cells by key


1.7.0 (2013-10-31)
==================


Major changes
-------------

Drops support for Python < 2.5 and last version to support Python 2.5


Compatibility
-------------

Tests run on Python 2.5, 2.6, 2.7, 3.2, 3.3


Merged pull requests
--------------------

* 27 Include more metadata
* 41 Able to read files with chart sheets
* 45 Configurable Worksheet classes
* 3 Correct serialisation of Decimal
* 36 Preserve VBA macros when reading files
* 44 Handle empty oddheader and oddFooter tags
* 43 Fixed issue that the reader never set the active sheet
* 33 Reader set value and type explicitly and TYPE_ERROR checking
* 22 added page breaks, fixed formula serialization
* 39 Fix Python 2.6 compatibility
* 47 Improvements in styling


Known bugfixes
--------------

* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#165 <https://bitbucket.org/openpyxl/openpyxl/issue/165>`_
* `#179 <https://bitbucket.org/openpyxl/openpyxl/issue/179>`_
* `#209 <https://bitbucket.org/openpyxl/openpyxl/issue/209>`_
* `#112 <https://bitbucket.org/openpyxl/openpyxl/issue/112>`_
* `#166 <https://bitbucket.org/openpyxl/openpyxl/issue/166>`_
* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#223 <https://bitbucket.org/openpyxl/openpyxl/issue/223>`_
* `#124 <https://bitbucket.org/openpyxl/openpyxl/issue/124>`_
* `#157 <https://bitbucket.org/openpyxl/openpyxl/issue/157>`_


Miscellaneous
-------------

Performance improvements in optimised writer

Docs updated
