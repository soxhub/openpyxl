# copyright openpyxl 2014

import datetime

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.constants import DCTERMS_PREFIX, DCTERMS_NS, XSI_NS
from openpyxl.xml.functions import (
    fromstring,
    tostring,
    register_namespace,
    NS_REGEX,
)


@pytest.fixture()
def SampleProperties():
    from .. core import DocumentProperties
    props = DocumentProperties()
    props.keywords = "one, two, three"
    props.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
    props.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
    props.lastPrinted = datetime.datetime(2014, 10, 14, 10, 30)
    props.category = "The category"
    props.contentStatus = "The status"
    props.creator = 'TEST_USER'
    props.lastModifiedBy = "SOMEBODY"
    props.revision = "0"
    props.version = "2.5"
    props.description = "The description"
    props.identifier = "The identifier"
    props.language = "The language"
    props.subject = "The subject"
    props.title = "The title"
    return props


def test_ctor(SampleProperties):
    expected = """
    <coreProperties
        xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        xmlns:dc="http://purl.org/dc/elements/1.1/"
        xmlns:dcterms="http://purl.org/dc/terms/"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        <dc:creator>TEST_USER</dc:creator>
        <dc:title>The title</dc:title>
        <dc:description>The description</dc:description>
        <dc:subject>The subject</dc:subject>
        <dc:identifier>The identifier</dc:identifier>
        <dc:language>The language</dc:language>
        <dcterms:created xsi:type="dcterms:W3CDTF">2010-04-01T20:30:00Z</dcterms:created>
        <dcterms:modified xsi:type="dcterms:W3CDTF">2010-04-05T14:05:30Z</dcterms:modified>
        <lastModifiedBy>SOMEBODY</lastModifiedBy>
        <category>The category</category>
        <contentStatus>The status</contentStatus>
        <version>2.5</version>
        <revision>0</revision>
        <keywords>one, two, three</keywords>
        <lastPrinted>2014-10-14T10:30:00Z</lastPrinted>
    </coreProperties>
    """
    xml = tostring(SampleProperties.to_tree())
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_from_tree(datadir, SampleProperties):
    datadir.chdir()
    with open("core.xml") as src:
        content = src.read()

    content = fromstring(content)
    props = SampleProperties.from_tree(content)
    assert props == SampleProperties


def test_qualified_datetime():
    from ..core import QualifiedDateTime
    dt = QualifiedDateTime()
    tree = dt.to_tree("time", datetime.datetime(2015, 7, 20, 12, 30))
    xml = tostring(tree)
    expected = """
    <time xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
      2015-07-20T12:30:00Z
    </time>"""

    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture(params=['abc', 'dct', 'dcterms', 'xyz'])
def dcterms_prefix(request):
    register_namespace(request.param, DCTERMS_NS)
    yield request.param
    register_namespace(DCTERMS_PREFIX, DCTERMS_NS)


@pytest.mark.no_pypy
def test_qualified_datetime_ns(dcterms_prefix):
    from ..core import QualifiedDateTime
    dt = QualifiedDateTime()
    tree = dt.to_tree("time", datetime.datetime(2015, 7, 20, 12, 30))
    xsi = tree.attrib["{%s}type" % XSI_NS]
    prefix = xsi.split(":")[0]
    assert prefix == dcterms_prefix
