'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@author: Eric Gazoni
'''

import re

def coordinate_from_string(coord_string):

    matches = re.match(pattern = '[$]?([A-Z]+)[$]?(\d+)', string = coord_string)

    if not matches:
        raise Exception('invalid cell coordinates')
    else:
        return matches.groups()


class Cell(object):

    def __init__(self, worksheet, column, row, value = None, data_type = None):

        self.column = column.upper()
        self.row = row

        self.value = value

        self.parent = worksheet

        self.data_type = data_type

    def coordinate(self):

        return '%s%s' % (self.column, self.row)

