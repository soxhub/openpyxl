# Copyright (c) 2010-2024 openpyxl

import pytest

# compatibility imports
import datetime

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.packaging.custom import CustomPropertyList, IntProperty, DateTimeProperty
from openpyxl.packaging.manifest import Manifest
from openpyxl.tests.helper import compare_xml
from openpyxl.workbook._writer import WorkbookWriter
from openpyxl.xml.constants import (
    ARC_CUSTOM,
    CPROPS_TYPE,
)


def test_read_custom_doc_props(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('example_vba_and_custom_doc_props.xlsm', read_only=False, keep_vba=True)
    custom_doc_props_dict = {
        "PropName1": {'value': datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22)},
        "PropName2": {'value': "ExampleName"},
        "PropName3": {'value': "Foo"},
    }
    for prop in wb.custom_doc_props:
        assert prop.value == custom_doc_props_dict[prop.name]['value']


def test_write_custom_doc_props(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('example_vba_and_no_custom_doc_props.xlsm')
    assert len(wb.custom_doc_props) == 0

    wb.custom_doc_props.append(DateTimeProperty(name="PropName1", value="2020-08-24T20:19:22Z"))
    wb.custom_doc_props.append(IntProperty(name="PropName2", value=2))

    writer = WorkbookWriter(wb)
    root_rels = writer.write_root_rels()

    custom_doc_props = tostring(wb.custom_doc_props.to_tree())
    class CustomOverride():
        path = "/" + ARC_CUSTOM #PartName
        mime_type = CPROPS_TYPE #ContentType

    custom_override = CustomOverride()
    # custom_override = Override(PartName="/docProps/custom.xml", ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml")
    manifest = Manifest()
    manifest.append(custom_override)
    root_manifest = tostring(manifest.to_tree())

    datadir.join("writer").chdir()
    with open('workbook_custom_doc_props.xml', 'r') as infile:
        file_xml = tostring(fromstring(infile.read()))
        diff = compare_xml(custom_doc_props, file_xml)
        assert diff is None, diff
    with open('workbook_root_rels_custom_doc_props.xml', 'r') as infile:
        file_xml = tostring(fromstring(infile.read()))
        diff = compare_xml(root_rels, file_xml)
        assert diff is None, diff
    # the source file will be created without knowing about vba project, sheet1 or the workbook,
    # so they are not included in the manifest example, as we are not testing those parts.
    with open('workbook_manifest_custom_doc_props.xml', 'r') as infile:
        file_xml = tostring(fromstring(infile.read()))
        diff = compare_xml(root_manifest, file_xml)
        assert diff is None, diff


def test_append_custom_props():
    """Tests the append method of the CustomPropertyList class.
    Appends properties to an empty property list and verifies the result."""
    props_list = CustomPropertyList()
    # Append custom properties
    n_properties = 10
    custom_props = {f"PropName{i}": i for i in range(n_properties)}
    for name, value in custom_props.items():
        props_list.append(IntProperty(name=name, value=value))
    # Assert that the properties were properly appended
    for prop in props_list:
        assert prop.value == custom_props[prop.name]


def test_append_repeated_prop():
    """Appends an existing custom property (with repeated name) and asserts
     that a ValueError is raised and that 'already exists' appears in the
     error message."""
    props_list = CustomPropertyList()
    props_list.append(IntProperty(name='foo', value=0))
    with pytest.raises(ValueError) as err:
        props_list.append(IntProperty(name='foo', value=1))

    assert 'already exists' in str(err).lower()
