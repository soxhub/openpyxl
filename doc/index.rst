openpyxl - A Python library to read/write Excel 2010 xlsx/xlsm files
===========================================================================


:Author: Eric Gazoni, Charlie Clark
:Source code: https://foss.heptapod.net/openpyxl/openpyxl
:Issues: https://foss.heptapod.net/openpyxl/openpyxl/-/issues
:Generated: |today|
:License: MIT/Expat
:Version: |release|


.. include:: ../README.rst


Support
-------

This is an open source project, maintained by volunteers in their spare time.
This may well mean that particular features or functions that you would like
are missing. But things don't have to stay that way. You can contribute the
project :doc:`development` yourself or contract a developer for particular
features.


Professional support for openpyxl is available from
`Clark Consulting & Research <http://www.clark-consulting.eu/>`_ and
`Adimian <http://www.adimian.com>`_. Donations to the project to support further
development and maintenance are welcome.


Bug reports and feature requests should be submitted using the `issue tracker
<https://foss.heptapod.net/openpyxl/openpyxl/-/issues>`_. Please provide a full
traceback of any error you see and if possible a sample file. If for reasons
of confidentiality you are unable to make a file publicly available then
contact of one the developers.

The repository is being provided by `Octobus <https://octobus.net/>`_ and
`Clever Cloud <https://clever-cloud.com/>`_.


How to Contribute
-----------------

Any help will be greatly appreciated, just follow those steps:

    1.
    Please join the group and create a branch (https://foss.heptapod.net/openpyxl/openpyxl/) and
    follow the `Merge Request Start Guide <https://heptapod.net/pages/quick-start-guide.html>`_.
    for each independent feature, don't try to fix all problems at the same
    time, it's easier for those who will review and merge your changes ;-)

    2.
    Hack hack hack

    3.
    Don't forget to add unit tests for your changes! (YES, even if it's a
    one-liner, changes without tests will **not** be accepted.) There are plenty
    of examples in the source if you lack know-how or inspiration.

    4.
    If you added a whole new feature, or just improved something, you can
    be proud of it, so add yourself to the AUTHORS file :-)

    5.
    Let people know about the shiny thing you just implemented, update the
    docs!

    6.
    When it's done, just issue a pull request (click on the large "pull
    request" button on *your* repository) and wait for your code to be
    reviewed, and, if you followed all theses steps, merged into the main
    repository.


For further information see :doc:`development`


Other ways to help
++++++++++++++++++

There are several ways to contribute, even if you can't code (or can't code well):

    * triaging bugs on the bug tracker: closing bugs that have already been
      closed, are not relevant, cannot be reproduced, ...

    * updating documentation in virtually every area: many large features have
      been added (mainly about charts and images at the moment) but without any
      documentation, it's pretty hard to do anything with it

    * proposing compatibility fixes for different versions of Python: we support
      3.6, 3.7, 3.8 and 3.9.


.. toctree::
    :maxdepth: 1
    :caption: Introduction
    :hidden:

    tutorial
    usage
    faq


.. toctree::
    :caption: Styling
    :maxdepth: 1
    :hidden:

    styles
    rich_text
    formatting

.. toctree::
    :maxdepth: 1
    :caption: Worksheets
    :hidden:

    editing_worksheets
    worksheet_properties
    validation
    worksheet_tables
    filters
    print_settings
    pivot
    comments

.. toctree::
    :maxdepth: 1
    :caption: Workbooks
    :hidden:

    defined_names
    workbook_custom_doc_props
    protection

.. toctree::
    :maxdepth: 1
    :caption: Charts
    :hidden:

    charts/introduction

.. toctree::
    :caption: Advanced Settings
    :maxdepth: 1
    :hidden:

    optimized
    performance
    pandas
    formula
    datetime

.. toctree::
    :caption: Developers
    :maxdepth: 1
    :hidden:

    development
    api/openpyxl
       
.. toctree::
    :maxdepth: 1
    :caption: Release Notes
    :hidden:

    changes


API Documentation
------------------

Key Classes
+++++++++++

* :class:`openpyxl.workbook.workbook.Workbook`
* :class:`openpyxl.worksheet.worksheet.Worksheet`
* :class:`openpyxl.cell.cell.Cell`


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
