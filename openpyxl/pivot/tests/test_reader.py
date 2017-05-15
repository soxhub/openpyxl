import pytest

from zipfile import ZipFile
from openpyxl.xml.constants import ARC_CONTENT_TYPES
from openpyxl.xml.functions import fromstring
from openpyxl.packaging.manifest import Manifest

from ..reader import read_pivot
from ..cache import CacheDefinition
from ..record import RecordList
from ..table import TableDefinition


def test_read_package(datadir):
    datadir.chdir()

    archive = ZipFile('pivot.xlsx')
    src = archive.read(ARC_CONTENT_TYPES)
    tree = fromstring(src)
    manifest = Manifest.from_tree(tree)
    ct = manifest.find(TableDefinition.mime_type)
    path = ct.PartName[1:]
    pivot = read_pivot(archive, path)

    assert isinstance(pivot, TableDefinition)
