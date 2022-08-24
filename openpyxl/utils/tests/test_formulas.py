import pytest
from ..formulas import validate


class TestValidate:


    def test_invalid_name(self):
        with pytest.raises(ValueError):
            validate("=CHARGE()")


    def test_empty(self):
        validate("==")


    def test_extension(self):
        validate("=_xlfn.CONCAT()")


    def test_not_a_function(self):
        with pytest.raises(AssertionError):
            validate("SUM(A1, A2")
