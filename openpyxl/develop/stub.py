"""
Generate test stubs for Serialisable classes
"""
import argparse


stub = '''
from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def {clsname}():
    from ..{module} import {clsname}
    return {clsname}


class Test{clsname}:

    def test_ctor(self, {clsname}):
        {module} = {clsname}()
        xml = tostring({module}.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, {clsname}):
        src = """
        <root />
        """
        node = fromstring(src)
        {module} = {clsname}.from_tree(node)
        assert {module} == {clsname}()
'''


def generate(clsname, module):
    return stub.format(clsname=clsname, module=module)

commands = argparse.ArgumentParser(description="Generate stub tests for Serialisable classes")
commands.add_argument('clsname', help='Name of the class to be tested')
commands.add_argument('--module', help='Name of the module to be tested', default="fut")


if __name__ == "__main__":
    args = commands.parse_args()
    print(generate(args.clsname, args.module))
