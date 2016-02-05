from openpyxl.styles import HashableObject
import pytest


@pytest.fixture
def Immutable():

    class Immutable(HashableObject):

        __fields__ = __elements__ = ('value',)

        def __init__(self, value=None):
            self.value = value

    return Immutable


class TestHashable:

    def test_ctor(self, Immutable):
        d = Immutable()
        d.value = 1
        assert d.value == 1


    def test_copy(self, Immutable):
        d = Immutable()
        d.value = 1
        c = d.copy()
        assert c == d


    def test_hash(self, Immutable):
        d1 = Immutable()
        d2 = Immutable()
        assert hash(d1) == hash(d2)


    def test_str(self, Immutable):
        d = Immutable()
        assert str(d) == """<openpyxl.styles.tests.test_hashable.Immutable object>
Parameters:

value:None"""

        d2 = Immutable("hello")
        assert str(d2) == """<openpyxl.styles.tests.test_hashable.Immutable object>
Parameters:

value:'hello'"""


    def test_eq(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(1)
        assert d1 is not d2
        assert d1 == d2


    def test_ne(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(2)
        assert d1 != d2

