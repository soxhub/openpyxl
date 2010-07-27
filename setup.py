#!/usr/bin/env python
"""Requires setuptools.

To build the setuptools egg use
    python setup.py bdist_egg
and either upload it to the PyPI with:
    python setup.py upload
or upload to your own server and register the release with PyPI:
    python setup.py register

A source distribution (.zip) can be built with
    python setup.py sdist --format=zip
That uses the manifest.in file for data files rather than searching for them here
"""

from setuptools import setup, Extension, find_packages
import sys
sys.path.append('src')#so we can import openpyxl
import openpyxl#to fetch __version__ etc

setup(name="openpyxl",
    packages=find_packages('src'),
    include_package_data =True,
    package_dir = {'':'src'},
    package_data = {'': ['tests/*.xml','tests/*.xslx']},#DOESN'T AFFECT zip DISTRIBUTION. MUST MODIFY MANIFEST.in TOO
    #metadata
    version = openpyxl.__version__,
    description = "A Python library to read/write Excel 2007 xlsx/xlsm files",
    long_description = "openpyxl is a pure python reader and writer of Excel OpenXML files." + \
        " It is ported from the PHPExcel project",
    author= openpyxl.__author__,
    author_email= openpyxl.__author_email__,
    url=openpyxl.__url__,
    license=openpyxl.__license__,
    download_url=openpyxl.__downloadUrl__,
    classifiers=['Development Status :: 4 - Beta',
          'Operating System :: MacOS :: MacOS X',
          'Operating System :: Microsoft :: Windows',
          'Operating System :: POSIX',
          'Programming Language :: Python'],
    )
