from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import re

"""
Base class for worksheets, chartsheets, etc. that can be added to workbooks
"""

INVALID_TITLE_REGEX = re.compile(r'[\\*?:/\[\]]')

def avoid_duplicate_name(names, value):
    # check if sheet_name already exists
    # do this *before* length check
    if value in names:
        names = ",".join(names)
        sheet_title_regex = re.compile("(?P<title>%s)(?P<count>\d*),?" % re.escape(value))
        matches = sheet_title_regex.findall(names)
        if matches:
            # use name, but append with the next highest integer
            counts = [int(idx) for (t, idx) in matches if idx.isdigit()]
            highest = max(counts) or 0
            value = "%s%d" % (value, highest + 1)
    return value


class _WorkbookChild(object):

    __title = ""
    __parent = None

    def __init__(self, parent=None, title=None):
        self.__parent = parent
        if title is not None:
            self.title = title


    @property
    def parent(self):
        return self.__parent


    @property
    def encoding(self):
        return self.__parent.encoding


    @property
    def title(self):
        return self.__title


    @title.setter
    def title(self, value):
        """
        Set a sheet title, ensuring it is valid.
        Limited to 31 characters, no special characters.
        Duplicate titles will be incremented numerically
        """
        if hasattr(value, "decode"):
            if not isinstance(value, unicode):
                try:
                    value = value.decode("ascii")
                except UnicodeDecodeError:
                    raise ValueError("Worksheet titles must be unicode")

        m = INVALID_TITLE_REGEX.search(value)
        if m:
            msg = "Invalid character {0} found in sheet title".format(m.group(0))
            raise ValueError(msg)

        sheets = self.parent.sheetnames

        if self.title is not None and self.title != value:
            value = avoid_duplicate_name(sheets, value)

        if len(value) > 31:
            msg = 'Maximum 31 characters allowed in sheet title'
            raise ValueError(msg)

        self.__title = value
