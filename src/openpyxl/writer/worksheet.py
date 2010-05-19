# coding=UTF-8
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

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

from xml.etree.cElementTree import ElementTree, Element, SubElement

from openpyxl.shared.xmltools import get_document_content


def write_worksheet(worksheet, string_table):

    root = Element('workseet', {'xml:space':'preserve',
                                'xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                                'xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})

    # sheet pr
    sheet_pr = SubElement(root, 'sheetPr')
    SubElement(sheet_pr, 'outlinePr', {'summaryBelow' : '%d' % (1 if worksheet.show_summary_below else 0),
                                       'summaryRight' : '%d' % (1 if worksheet.show_summary_right else 0)
                                       })

    # dimensions
    SubElement(root, 'dimension', {'ref' : '%s' % worksheet.calculate_dimension()})

    # sheet views
    sheet_views = SubElement(root, 'sheetViews')
    sheet_view = SubElement(sheet_views, 'sheetView', {'workbookViewId' : '0'})
    SubElement(sheet_view, 'selection', {'activeCell' : worksheet.active_cell,
                                         'sqref' : worksheet.selected_cell})

    # sheet format pr
    SubElement(root, 'sheetFormatPr', {'defaultRowHeight' : '15'})

    # sheet data
    write_worksheet_data(root, worksheet, string_table)

    return get_document_content(xml_node = root)

def write_worksheet_data(root_node, worksheet, string_table):

    sheet_data = SubElement(root_node, 'sheetData')

    max_column = worksheet.get_highest_column()

    cells_by_row = {}
    for cell in worksheet.get_cell_collection():
        cells_by_row.setdefault(cell.row, []).append(cell)

    for row_idx in cells_by_row:
        row_dimension = worksheet.row_dimensions[row_idx]

        row = SubElement(sheet_data, 'row', {'r' : '%d' % row_idx,
                                             'spans' : '1:%d' % max_column})

        for cell in cells_by_row[row_idx]:

            c = SubElement(row, 'c', {'r' : cell.get_coordinate(),
                                      't' : cell.data_type})

            if cell.data_type == cell.TYPE_STRING:
                SubElement(c, 'v').text = '%s' % string_table[cell.value]
            else:
                SubElement(c, 'v').text = '%s' % cell.value
