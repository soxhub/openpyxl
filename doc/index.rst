:mod:`openpyxl` - A Python library to read/write Excel 2010 xlsx/xlsm files
===========================================================================

.. module:: openpyxl
.. moduleauthor:: Eric Gazoni, Charlie Clark

:Author: Eric Gazoni, Charlie Clark
:Source code: http://bitbucket.org/openpyxl/openpyxl/src
:Issues: http://bitbucket.org/openpyxl/openpyxl/issues
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
<https://bitbucket.org/openpyxl/openpyxl/issues>`_. Please provide a full
traceback of any error you see and if possible a sample file. If for reasons
of confidentiality you are unable to make a file publicly available then
contact of one the developers.


How to Contribute
-----------------

Any help will be greatly appreciated, just follow those steps:

    1.
    Please start a new fork (https://bitbucket.org/openpyxl/openpyxl/fork)
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
      2.7, 3.4, 3.5, 3.6 and 3.7


Installation
------------

Install openpyxl using pip. It is advisable to do this in a Python virtualenv
without system packages::

    $ pip install openpyxl

.. note::

    There is support for the popular `lxml`_ library which will be used if it
    is installed. This is particular useful when creating large files.

.. _lxml: http://lxml.de

.. warning::

    To be able to include images (jpeg, png, bmp,...) into an openpyxl file,
    you will also need the "pillow" library that can be installed with::

    $ pip install pillow

    or browse https://pypi.python.org/pypi/Pillow/, pick the latest version
    and head to the bottom of the page for Windows binaries.


Working with a checkout
+++++++++++++++++++++++

Sometimes you might want to work with the checkout of a particular version.
This may be the case if bugs have been fixed but a release has not yet been
made.

.. parsed-literal::
    $ pip install -e hg+https://bitbucket.org/openpyxl/openpyxl@\ |version|\ #egg=openpyxl


Usage examples
--------------


Tutorial
++++++++

.. toctree::

    tutorial


Cookbook
++++++++

.. toctree::

    usage


Performance
-----------

.. toctree::

    performance


Other topics
------------

    .. toctree::
        :maxdepth: 1

        optimized


    .. toctree::
        :maxdepth: 1

        editing_worksheets

    .. toctree::
        :maxdepth: 1

        pandas

    .. toctree::
        :maxdepth: 1

        charts/introduction

    .. toctree::
        :maxdepth: 1

        comments

    .. toctree::
        :maxdepth: 1

        styles

    .. toctree::
        :maxdepth: 1

        worksheet_properties

    .. toctree::
        :maxdepth: 1

        formatting

    .. toctree::
        :maxdepth: 1

        print_settings

    .. toctree::
        :maxdepth: 1

        filters

    .. toctree::
        :maxdepth: 1

        validation


    .. toctree::
        :maxdepth: 1

        defined_names

    .. toctree::
        :maxdepth: 1

        worksheet_tables

    .. toctree::
        :maxdepth: 1

        pivot

    .. toctree::
        :maxdepth: 1

        formula

    .. toctree::
        :maxdepth: 1

        protection


Information for Developers
--------------------------

    .. toctree::
        :maxdepth: 1

        development
        windows-development


API Documentation
------------------

Key Classes
+++++++++++

* :class:`openpyxl.workbook.workbook.Workbook`
* :class:`openpyxl.worksheet.worksheet.Worksheet`
* :class:`openpyxl.cell.cell.Cell`


Full API
++++++++

.. toctree::
    :maxdepth: 2

    api/openpyxl


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`


Release Notes
=============

.. toctree::
    :maxdepth: 1

    changes
