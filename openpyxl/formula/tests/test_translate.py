import pytest

@pytest.fixture
def Translator():
    from .. import translate
    return translate.Translator

@pytest.fixture
def Tokenizer():
    from .. import tokenizer
    return tokenizer.Tokenizer

@pytest.fixture
def TranslatorError():
    from .. import translate
    return translate.TranslatorError

class TestTranslator(object):

    @pytest.mark.parametrize("origin, row, col", [
        ("A1", 1, 1),
        ("AA1", 1, 27),
        ("AA1001", 1001, 27),
        ("AA1001001001", 1001001001, 27),
        ("XFD111", 111, 16384),
    ])
    def test_init(self, Translator, origin, row, col):
        trans = Translator("=formula", origin)
        assert trans.row == row
        assert trans.col == col
        assert trans.tokenizer.formula == "=formula"

    @pytest.mark.parametrize("formula", [
        '=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
        "='Summary slices'!$C$3",
        '=-MAX(Pipeline!AA4:AA138)',
        '=TEXT(-S7/1000,"$#,##0""M""")',
        "=IF(A$3<1.3E-8,\"\",IF(ISNA('External Ref'!K7)," +
         '"N/A",TEXT(K7*1E+12,"0")&"bp"',
        '=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=' +
         'LOWER_BOUND)*($BR6>$S72123))',
        '=$A:$A,$C:$C',
        "Just text"
        "123.456",
        "31/12/1999",
        "",
    ])
    def test_get_tokens(self, Translator, Tokenizer, formula):
        trans = Translator(formula, "A1")
        tok = Tokenizer(formula)
        for t1, t2 in zip(tok.items, trans.get_tokens()):
            assert t1.value == t2.value
            assert t1.type == t2.type
            assert t1.subtype == t2.subtype

    @pytest.mark.parametrize("test_str, groups", [
        ["1:1", ("1", "1")],
        ["1234:5678", ("1234", "5678")],
        ["$1234:78910", ("$1234", "78910")],
        ["$12321:$23432", ("$12321", "$23432")],
        ["112233:$445566", ("112233", "$445566")],
        ["A:A", None],
        ["$ABC:AZZ", None],
        ["$DEF:$FOV",None],
        ["HA:$JA", None],
        ["named1", None],
        ["A15", None],
        ["$AB303", None],
        ["YY$101", None],
        ["$ZZ$99", None],
        ["B2:C3", None],
        ["$ATV25:$BBC35", None],
        ["WWW$918:WWW$919", None],
        ["$III$305:$IIT$503", None],
    ])
    def test_row_range_re(self, Translator, test_str, groups):
        match = Translator.ROW_RANGE_RE.match(test_str)
        if groups is None:
            assert match is None
        else:
            assert match.groups() == groups

    @pytest.mark.parametrize("test_str, groups", [
        ["1:1", None],
        ["1234:5678", None],
        ["$1234:78910", None],
        ["$12321:$23432", None],
        ["112233:$445566", None],
        ["A:A", ("A", "A")],
        ["$ABC:AZZ", ("$ABC", "AZZ")],
        ["$DEF:$FOV", ("$DEF", "$FOV")],
        ["HA:$JA", ("HA", "$JA")],
        ["named1", None],
        ["A15", None],
        ["$AB303", None],
        ["YY$101", None],
        ["$ZZ$99", None],
        ["B2:C3", None],
        ["$ATV25:$BBC35", None],
        ["WWW$918:WWW$919", None],
        ["$III$305:$IIT$503", None],
    ])
    def test_col_range_re(self, Translator, test_str, groups):
        match = Translator.COL_RANGE_RE.match(test_str)
        if groups is None:
            assert match is None
        else:
            assert match.groups() == groups

    @pytest.mark.parametrize("test_str, groups", [
        ["1:1", None],
        ["1234:5678", None],
        ["$1234:78910", None],
        ["$12321:$23432", None],
        ["112233:$445566", None],
        ["A:A", None],
        ["$ABC:AZZ", None],
        ["$DEF:$FOV", None],
        ["HA:$JA", None],
        ["named1", None],
        ["A15", ("A", "15")],
        ["$AB303", ("$AB", "303")],
        ["YY$101", ("YY", "$101")],
        ["$ZZ$99", ("$ZZ", "$99")],
        ["B2:C3", None],
        ["$ATV25:$BBC35", None],
        ["WWW$918:WWW$919", None],
        ["$III$305:$IIT$503", None],
    ])
    def test_cell_ref_re(self, Translator, test_str, groups):
        match = Translator.CELL_REF_RE.match(test_str)
        if groups is None:
            assert match is None
        else:
            assert match.groups() == groups

    @pytest.mark.parametrize("test_str, rdelta, value", [
        ('1', 1, '2'),
        ('$222333', 1, '$222333'),
        ('1048576', -100, '1048476'),
        ('$1012023', -100, '$1012023'),
        ('101', 0, '101'),
        ('$101', 0, '$101'),
        ('12', -15, None),
        ('$12', -15, '$12'),
    ])
    def test_translate_row(self, Translator, TranslatorError,
                           test_str, rdelta, value):
        if value is None:
            with pytest.raises(TranslatorError):
                Translator.translate_row(test_str, rdelta)
        else:
            assert Translator.translate_row(test_str, rdelta) == value

    @pytest.mark.parametrize("test_str, cdelta, value", [
        ('A', 1, 'B'),
        ('XED', 26, 'XFD'),
        ('$XED', 26, '$XED'),
        ('WWW', -52, 'WUW'),
        ('$WWW', -52, '$WWW'),
        ('ABC', 0, 'ABC'),
        ('$ABC', 0, '$ABC'),
        ('AA', -100, None),
        ('$AA', -100, '$AA'),
    ])
    def test_translate_col(self, Translator, TranslatorError,
                           test_str, cdelta, value):
        if value is None:
            with pytest.raises(TranslatorError):
                Translator.translate_col(test_str, cdelta)
        else:
            assert Translator.translate_col(test_str, cdelta) == value

    @pytest.mark.parametrize("test_str, value", [
        ('A$3', ("", 'A$3')),
        ('Pipeline!B$4:B$138', ('Pipeline!', 'B$4:B$138')),
        ("'Summary slices'!$C$3", ("'Summary slices'!", '$C$3')),
        ("'Lions! Tigers! Bears!'!$OM$1", ("'Lions! Tigers! Bears!'!", '$OM$1')),
        ('named_range_1', ("", 'named_range_1')),
        ('Sheet-2!named_range_2', ("Sheet-2!", 'named_range_2')),
    ])
    def test_strip_ws_name(self, Translator, test_str, value):
        assert Translator.strip_ws_name(test_str) == value

    @pytest.mark.parametrize("test_str, rdelta, cdelta, value", [
        ("1:1", 2, 1, "3:3"),
        ("$1234:78910", 1, 10, "$1234:78911"),
        ("$12321:$23432", 3, 5, "$12321:$23432"),
        ("112233:$445566", -3, -20, "112230:$445566"),
        ("987:999", 0, 12, "987:999"),
        ("1:5", -2, 3, None),
        ("A:A", 0, 1, "B:B"),
        ("$ABC:AZZ", 1, 3, "$ABC:BAC"),
        ("$DEF:$FOV", 25, 25, "$DEF:$FOV"),
        ("HA:$JA", -5, -15, "GL:$JA"),
        ("named1", -33, 33, "named1"),
        ("A15", -3, 4, "E12"),
        ("$AB303", 3, 2, "$AB306"),
        ("YY$101", 4, 2, "ZA$101"),
        ("$ZZ$99", 5, 2, "$ZZ$99"),
        ("B2:C3", 4, 3, "E6:F7"),
        ("$ATV25:$BBC35", 5, 3, "$ATV30:$BBC40"),
        ("WWW$918:WWW$919", 5, 4, "WXA$918:WXA$919"),
        ("$III$305:$IIT$503", 25, 35, "$III$305:$IIT$503"),
    ])
    def test_translate_range(self, Translator, TranslatorError, test_str,
                             rdelta, cdelta, value):
        if value is None:
            with pytest.raises(TranslatorError):
                Translator.translate_range(test_str, rdelta, cdelta)
        else:
            assert value == Translator.translate_range(test_str,
                                                       rdelta, cdelta)

    @pytest.mark.parametrize("formula, origin, dest, result", [
        ('=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))', "A1", "B2",
         '=IF(B$3<40%,"",INDEX(Pipeline!C$4:C$138,#REF!))'),
        ("='Summary slices'!$C$3", "A1", "B2", "='Summary slices'!$C$3"),
        ('=-MAX(Pipeline!AA4:AA138)', "A1", "B2",
         '=-MAX(Pipeline!AB5:AB139)'),
        ('=TEXT(-\'External Ref\'!K7/DENOMINATOR,"$#,##0""M""")', "A1", "B2",
         '=TEXT(-\'External Ref\'!L8/DENOMINATOR,"$#,##0""M""")'),
        ("=ROWS('Sh 1'!$1:3)+COLUMNS('Sh 2'!$A:C)", "A1", "B2",
         "=ROWS('Sh 1'!$1:4)+COLUMNS('Sh 2'!$A:D)"),
        ("Just text", "A1", "B2", "Just text"),
        ("123.456", "A1", "B2", "123.456"),
        ("31/12/1999", "A1", "B2", "31/12/1999"),
        ("", "A1", "B2", ""),
    ])
    def test_translate_formula_range(self, Translator, formula,
                               origin, dest, result):
        trans = Translator(formula, origin)
        assert trans.translate_formula(dest) == result


    def test_translate_formula_coordinates(self, Translator):
        trans = Translator("='Summary slices'!C3", "A1")
        result = trans.translate_formula(row_delta=2, col_delta=3)
        assert result == "='Summary slices'!F5"
