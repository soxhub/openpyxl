from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import datetime

from openpyxl.compat import safe_string, unicode
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, datetime_to_W3CDTF, W3CDTF_to_datetime
from openpyxl.descriptors import (
    String,
    DateTime,
    Alias,
    )
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.nested import NestedText
from openpyxl.xml.functions import ElementTree, Element, SubElement, tostring, fromstring, safe_iterator, localname
from openpyxl.xml.constants import COREPROPS_NS, DCORE_NS, XSI_NS, DCTERMS_NS, DCTERMS_PREFIX


class NestedDateTime(DateTime, NestedText):

    expected_type = datetime.datetime

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)
        el = Element(tagname)
        if value is not None:
            el.text = datetime_to_W3CDTF(value)
        return el


class DocumentProperties(Serialisable):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    tagname = "coreProperties"
    namespace = COREPROPS_NS

    category = NestedText(expected_type=unicode, allow_none=True)
    contentStatus = NestedText(expected_type=unicode, allow_none=True)
    keywords = NestedText(expected_type=unicode, allow_none=True)
    lastModifiedBy = NestedText(expected_type=unicode, allow_none=True)
    lastPrinted = NestedDateTime(allow_none=True)
    revision = NestedText(expected_type=unicode, allow_none=True)
    version = NestedText(expected_type=unicode, allow_none=True)
    last_modified_by = Alias("lastModifiedBy")

    # Dublin Core Properties
    subject = NestedText(expected_type=unicode, allow_none=True, namespace=DCORE_NS)
    title = NestedText(expected_type=unicode, allow_none=True, namespace=DCORE_NS)
    creator = NestedText(expected_type=unicode, allow_none=True, namespace=DCORE_NS)
    description = NestedText(expected_type=unicode, allow_none=True, namespace=DCORE_NS)
    identifier = NestedText(expected_type=unicode, allow_none=True, namespace=DCORE_NS)
    language = NestedText(expected_type=unicode,allow_none=True, namespace=DCORE_NS)
    # Dubline Core Terms
    created = NestedDateTime(allow_none=True, namespace=DCTERMS_NS)
    modified = NestedDateTime(allow_none=True, namespace=DCTERMS_NS)

    __elements__ = ("creator","title", "description", "subject","identifier",
                  "language", "created", "modified", "lastModifiedBy", "category",
                  "contentStatus", "version", "revision", "keywords", "lastPrinted",
                  )

    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=None,
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=datetime.datetime.now(),
                 creator="openpyxl",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=datetime.datetime.now(),
                 subject=None,
                 title=None,
                 ):
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.modified = modified
        self.created = created
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self.keywords = keywords
        self.category = category
