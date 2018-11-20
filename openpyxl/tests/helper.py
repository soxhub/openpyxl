from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from io import BytesIO
from zipfile import ZipFile

# Python stdlib imports
from lxml.doctestcompare import LXMLOutputChecker, PARSE_XML


def compare_xml(generated, expected):
    """Use doctest checking from lxml for comparing XML trees. Returns diff if the two are not the same"""
    checker = LXMLOutputChecker()

    class DummyDocTest():
        pass

    ob = DummyDocTest()
    ob.want = expected

    check = checker.check_output(expected, generated, PARSE_XML)
    if check is False:
        diff = checker.output_difference(ob, generated, PARSE_XML)
        return diff


def _modify_zip_file(zip_path, filename, replacement_content=None):
    """
    Replace the contents of a file inside a ZIP-container. The original file
    will not be overwritten. A BytesIO() file will be returned containing
    the new ZIP-container with the file replaced.

    :param zip_path The path to the ZIP-File to be modified
    :param filename Filename inside the container to be modified
    :param replacement_content The new container which will be stored in `filename`
    :return BytesIO contain the modified zip file.
    """
    vuln_zipfile = BytesIO()
    zw = ZipFile(vuln_zipfile, 'w')
    zr = ZipFile(zip_path)

    names_to_copy = set(zr.namelist()) - {filename, }
    for name in names_to_copy:
        content = zr.read(name)
        zw.writestr(name, content)

    if replacement_content:
        zw.writestr(filename, replacement_content)

    zw.close()
    vuln_zipfile.seek(0)

    return vuln_zipfile


def modify_zip_file(zip_path, filename, replacements):
    """
    Performan a list of replacments on `filename` the original file in
    `zip_path` will remain unaffected and a BytesIO object will be returned
    containing the new zip file.

    :param zip_path Path to the zip-file.
    :param filename Filename where the replacements will be performed on
    :param replacements An iterable containing replacements using two-tuples.
                        e.g. [(orig_string, new_string), ]
    :return BytesIO contain the modified zip file.
    """
    zr = ZipFile(zip_path)
    original_content = zr.read(filename).decode('utf-8')
    new_content = original_content

    for old, new in replacements:
        new_content = new_content.replace(old, new)

    return _modify_zip_file(zip_path, filename, replacement_content=new_content.encode('utf-8'))
