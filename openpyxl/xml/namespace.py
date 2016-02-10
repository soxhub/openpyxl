from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl
#

try:
    from xml.etree.cElementTree import register_namespace
except ImportError:
    from xml.etree.ElementTree import register_namespace
