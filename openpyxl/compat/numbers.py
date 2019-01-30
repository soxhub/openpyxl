# Copyright (c) 2010-2019 openpyxl

from decimal import Decimal

NUMERIC_TYPES = (int, float, Decimal)


try:
    import numpy
    NUMPY = True
except ImportError:
    NUMPY = False


if NUMPY:
    NUMERIC_TYPES = NUMERIC_TYPES + (numpy.bool_, numpy.floating, numpy.integer)


try:
    import pandas
    PANDAS = True
except ImportError:
    PANDAS = False
