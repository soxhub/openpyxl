from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

try:
    # Python 2
    long = long
except NameError:
    # Python 3
    long = int

from decimal import Decimal

NUMERIC_TYPES = (int, float, long, Decimal)

def numpy_available():
    try:
        import numpy
        return True
    except ImportError:
        return False

NUMPY = numpy_available()

if NUMPY:
    NUMERIC_TYPES = NUMERIC_TYPES + (numpy.bool_, numpy.floating, numpy.integer)
