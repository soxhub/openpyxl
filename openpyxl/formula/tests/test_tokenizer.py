import pytest

@pytest.fixture
def tokenizer():
    from .. import tokenizer
    return tokenizer

# Constants from tokenizer.Token:
LITERAL = "LITERAL"
OPERAND = "OPERAND"
FUNC = "FUNC"
ARRAY = "ARRAY"
PAREN = "PAREN"
SEP = "SEP"
OP_PRE = "OPERATOR-PREFIX"
OP_IN = "OPERATOR-INFIX"
OP_POST = "OPERATOR-POSTFIX"
WSPACE = "WHITE-SPACE"
TEXT = 'TEXT'
NUMBER = 'NUMBER'
LOGICAL = 'LOGICAL'
ERROR = 'ERROR'
RANGE = 'RANGE'
OPEN = "OPEN"
CLOSE = "CLOSE"
ARG = "ARG"
ROW = "ROW"


class TestTokenizerRegexes(object):

    @pytest.mark.parametrize("string, success", [
        ('1.0E', True),
        ('1.53321E', True),
        ('9.999E', True),
        ('3E', True),
        ('12E', False),
        ('0.1E', False),
        ('0E', False),
        ('', False),
        ('E', False),
    ])
    def test_scientific_re(self, tokenizer, string, success):
        regex = tokenizer.Tokenizer.SN_RE
        assert bool(regex.match(string)) is success

    @pytest.mark.parametrize('string, expected', [
        (' ', ' '),
        (' *', ' '),
        ('     ', '     '),
        ('     a', '     '),
        ('   ', '   '),
        ('   +', '   '),
        ('', None),
        ('*', None),
    ])
    def test_whitespace_re(self, tokenizer, string, expected):
        if expected is None:
            assert not tokenizer.Tokenizer.WSPACE_RE.match(string)
        else:
            assert tokenizer.Tokenizer.WSPACE_RE.match(string)
            assert tokenizer.Tokenizer.WSPACE_RE.match(string).group(0) == expected

    @pytest.mark.parametrize('string, expected', [
        ('"spamspamspam"', '"spamspamspam"'),
        ('"this is "" a test "" "', '"this is "" a test "" "'),
        ('""', '""'),
        ('"spam and ""cheese"""+"ignore"', ('"spam and ""cheese"""')),
        ('\'"spam and ""cheese"""+"ignore"', None),
        ('"oops ""', None),
    ])
    def test_string_re(self, tokenizer, string, expected):
        regex = tokenizer.Tokenizer.STRING_REGEXES['"']
        if expected is None:
            assert not regex.match(string)
        else:
            assert regex.match(string)
            assert regex.match(string).group(0) == expected

    @pytest.mark.parametrize('string, expected', [
        ("'spam and ham'", "'spam and ham'"),
        ("'double'' triple''' quadruple ''''", "'double'' triple'''"),
        ("'sextuple '''''' and septuple''''''' and more",
         "'sextuple '''''' and septuple'''''''",),
         ("''", "''"),
         ("'oops ''", None),
         ("gunk'hello world'", None),
    ])
    def test_link_re(self, tokenizer, string, expected):
        regex = tokenizer.Tokenizer.STRING_REGEXES["'"]
        if expected is None:
            assert not regex.match(string)
        else:
            assert regex.match(string)
            assert regex.match(string).group(0) == expected


class TestTokenizer(object):

    def test_init(self, tokenizer):
        tok = tokenizer.Tokenizer("abcdefg")
        assert tok.formula == "abcdefg"
        tok = tokenizer.Tokenizer("=abcdefg")
        assert tok.formula == "=abcdefg"

    @pytest.mark.parametrize('formula, tokens', [
        ('=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
         [('IF(', FUNC, OPEN),
          ('A$3', OPERAND, RANGE),
          ('<', OP_IN, ""),
          ('40', OPERAND, NUMBER),
          ('%', OP_POST, ""),
          (',', SEP, ARG),
          ('""', OPERAND, TEXT),
          (',', SEP, ARG),
          ('INDEX(', FUNC, OPEN),
          ('Pipeline!B$4:B$138', OPERAND, RANGE),
          (',', SEP, ARG),
          ('#REF!', OPERAND, ERROR),
          (')', FUNC, CLOSE),
          (')', FUNC, CLOSE)]),

        ("='Summary slices'!$C$3",
         [("'Summary slices'!$C$3", OPERAND, RANGE)]),

        ('=-MAX(Pipeline!AA4:AA138)',
         [("-", OP_PRE, ""),
          ('MAX(', FUNC, OPEN),
          ('Pipeline!AA4:AA138', OPERAND, RANGE),
          (')', FUNC, CLOSE)]),

        ('=TEXT(-S7/1000,"$#,##0""M""")',
         [('TEXT(', FUNC, OPEN),
          ('-', OP_PRE, ""),
          ('S7', OPERAND, RANGE),
          ('/', OP_IN, ""),
          ('1000', OPERAND, NUMBER),
          (',', SEP, ARG),
          ('"$#,##0""M"""', OPERAND, TEXT),
          (')', FUNC, CLOSE)]),

        ("=IF(A$3<1.3E-8,\"\",IF(ISNA('External Ref'!K7)," +
         '"N/A",TEXT(K7*1E+12,"0")&"bp"',
         [('IF(', FUNC, OPEN),
          ('A$3', OPERAND, RANGE),
          ('<', OP_IN, ""),
          ('1.3E-8', OPERAND, NUMBER),
          (',', SEP, ARG),
          ('""', OPERAND, TEXT),
          (',', SEP, ARG),
          ('IF(', FUNC, OPEN),
          ('ISNA(', FUNC, OPEN),
          ("'External Ref'!K7", OPERAND, RANGE),
          (')', FUNC, CLOSE),
          (',', SEP, ARG),
          ('"N/A"', OPERAND, TEXT),
          (',', SEP, ARG),
          ('TEXT(', FUNC, OPEN),
          ('K7', OPERAND, RANGE),
          ('*', OP_IN, ""),
          ('1E+12', OPERAND, NUMBER),
          (',', SEP, ARG),
          ('"0"', OPERAND, TEXT),
          (')', FUNC, CLOSE),
          ('&', OP_IN, ""),
          ('"bp"', OPERAND, TEXT)]),

        ('=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=' +
         'LOWER_BOUND)*($BR6>$S72123))',
         [("+", OP_PRE, ""),
          ('IF(', FUNC, OPEN),
          ('A$3', OPERAND, RANGE),
          ('<>', OP_IN, ""),
          ('$B7', OPERAND, RANGE),
          (',', SEP, ARG),
          ('""', OPERAND, TEXT),
          (',', SEP, ARG),
          ('(', PAREN, OPEN),
          ('MIN(', FUNC, OPEN),
          ('IF(', FUNC, OPEN),
          ('{', ARRAY, OPEN),
          ('TRUE', OPERAND, LOGICAL),
          (',', SEP, ARG),
          (' ', WSPACE, ''),
          ('FALSE', OPERAND, LOGICAL),
          (';', SEP, ROW),
          ('1', OPERAND, NUMBER),
          (',', SEP, ARG),
          ('2', OPERAND, NUMBER),
          ('}', ARRAY, CLOSE),
          (',', SEP, ARG),
          ('A6:B6', OPERAND, RANGE),
          (',', SEP, ARG),
          ('$S7', OPERAND, RANGE ),
          (')', FUNC, CLOSE),
          (')', FUNC, CLOSE),
          ('>=', OP_IN, ''),
          ('LOWER_BOUND', OPERAND, RANGE),
          (')', PAREN, CLOSE),
          ('*', OP_IN, ''),
          ('(', PAREN, OPEN),
          ('$BR6', OPERAND, RANGE),
          ('>', OP_IN, ''),
          ('$S72123', OPERAND, RANGE),
          (')', PAREN, CLOSE),
          (')', FUNC, CLOSE)]),

        ('=(AW$4=$D7)+0%',
         [('(', PAREN, OPEN),
          ('AW$4', OPERAND, RANGE),
          ('=', OP_IN, ''),
          ('$D7', OPERAND, RANGE),
          (')', PAREN, CLOSE),
          ('+', OP_IN, ''),
          ('0', OPERAND, NUMBER),
          ('%', OP_POST, '')]),

        ('=$A:$A,$C:$C',
         [('$A:$A', OPERAND, RANGE),
          (',', OP_IN, ""),
          ('$C:$C', OPERAND, RANGE)]),

        ('=3 +1-5',
         [('3', 'OPERAND', 'NUMBER'),
          (' ', 'WHITE-SPACE', ''),
          ('+', 'OPERATOR-INFIX', ''),
          ('1', 'OPERAND', 'NUMBER'),
          ('-', 'OPERATOR-INFIX', ''),
          ('5', 'OPERAND', 'NUMBER')]),

        ("Just text", [("Just text", LITERAL, "")]),
        ("123.456", [("123.456", LITERAL, "")]),
        ("31/12/1999", [("31/12/1999", LITERAL, "")]),
        ("", []),

        ('=A1+\nA2',
         [('A1', OPERAND, RANGE),
          ('+', OP_IN, ''),
          ('\n', WSPACE, ''),
          ('A2', OPERAND, RANGE)]),

        ('=R[41]C[2]',
         [('R[41]C[2]', 'OPERAND', 'RANGE')]),
    ])
    def test_parse(self, tokenizer, formula, tokens):
        tok = tokenizer.Tokenizer(formula)
        result = [(token.value, token.type, token.subtype)
                  for token in tok.items]
        assert result == tokens
        assert tok.render() == formula

    @pytest.mark.parametrize('formula, offset, result', [
        ('"spamspamspam"spam', 0, '"spamspamspam"'),
        ('"this is "" a test "" "test', 0, '"this is "" a test "" "'),
        ('""', 0, '""'),
        ('a"bcd""efg"hijk', 1, '"bcd""efg"'),
        ('"oops ""', 0, None),
        ("'spam and ham'", 0, "'spam and ham'"),
        ("'double'' triple''' quad ''''", 0, "'double'' triple'''"),
        ("123'sextuple '''''' and septuple''''''' and more", 3,
         "'sextuple '''''' and septuple'''''''"),
         ("''", 0, "''"),
         ("'oops ''", 0, None),
    ])
    def test_parse_string(self, tokenizer, formula, offset, result):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        if result is None:
            with pytest.raises(tokenizer.TokenizerError):
                tok._parse_string()
            return
        assert tok._parse_string() == len(result)
        if formula[offset] == '"':
            token = tok.items[0]
            assert token.value == result
            assert token.type == OPERAND
            assert token.subtype == TEXT
            assert not tok.token
        else:
            assert not tok.items
            assert tok.token[0] == result
            assert len(tok.token) == 1

    @pytest.mark.parametrize('formula, offset, result', [
        ('[abc]def', 0, '[abc]'),
        ('[]abcdef', 0, '[]'),
        ('[abcdef]', 0, '[abcdef]'),
        ('a[bcd]ef', 1, '[bcd]'),
        ('ab[cde]f', 2, '[cde]'),
        ('R[1]C[2]', 1, '[1]'),
        ('R[1]C[2]', 5, '[2]'),
    ])
    def test_parse_brackets(self, tokenizer, formula, offset, result):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        tok._parse_brackets()
        assert tok.token[-1] == result

    @pytest.mark.parametrize('formula, offset, result', [
        ('[a[b]c]def', 0, '[a[b]c]'),
        ('[[]]abcdef', 0, '[[]]'),
        ('[[abc]def]', 0, '[[abc]def]'),
        ('a[b[c]d]e[f]', 1, '[b[c]d]'),
        ('ab[c[d][e]][f]', 2, '[c[d][e]]'),
        ('TableX[[#Data],[COL1]]', 6, '[[#Data],[COL1]]'),
        ('TableX[[#Data],[COL1]:[COL2]]', 6, '[[#Data],[COL1]:[COL2]]'),
    ])
    def test_parse_nested_brackets(self, tokenizer, formula, offset, result):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        tok._parse_brackets()
        assert tok.token[0] == result

    @pytest.mark.parametrize('formula, offset', [
        ('[unfinished business', 0),
        ('unfinished [business', 11),
        ('[un[finished business]', 0),
        ('un[finished [business]', 2),
    ])
    def test_parse_brackets_error(self, tokenizer, formula, offset):
        tok = tokenizer.Tokenizer(formula)
        tok.offset = offset
        with pytest.raises(tokenizer.TokenizerError):
            tok._parse_brackets()

    @pytest.mark.parametrize('error', [
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
    ])
    def test_parse_error(self, tokenizer, error):
        tok = tokenizer.Tokenizer(error)
        tok.offset = 0
        del tok.items[:]
        assert tok._parse_error() == len(error)
        assert len(tok.items) == 1
        assert not tok.token
        token = tok.items[0]
        assert token.value == error
        assert token.type == OPERAND
        assert token.subtype == ERROR

    def test_parse_defined_name_reference_error(self, tokenizer):
        formula = "=SUM(MyTable!#REF!)"
        tok = tokenizer.Tokenizer("=SUM(MyTable!#REF!)")
        result = [(token.value, token.type, token.subtype)
                  for token in tok.items]
        tokens = [
            ('SUM(', FUNC, OPEN),
            ('MyTable!#REF!', OPERAND, RANGE),
            (')', FUNC, CLOSE),
        ]

        assert result == tokens
        assert tok.render() == formula

    def test_parse_error_error(self, tokenizer):
        tok = tokenizer.Tokenizer("#NotAnError")
        tok.offset = 0
        del tok.items[:]
        with pytest.raises(tokenizer.TokenizerError):
            tok._parse_error()

    @pytest.mark.parametrize('formula, value',
        [(' ' * i, ' ') for i in range(1, 10)] + [('\n', '\n')])
    def test_parse_whitespace(self, tokenizer, formula, value):
        tok = tokenizer.Tokenizer(formula)
        tok.offset = 0
        del tok.items[:]
        assert tok._parse_whitespace() == len(formula)
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == value
        assert token.type == WSPACE
        assert token.subtype == ""
        assert not tok.token

    @pytest.mark.parametrize('formula, result, type_', [
        ('>=', '>=', OP_IN),
        ('<=', '<=', OP_IN),
        ('<>', '<>', OP_IN),
        ('%', '%', OP_POST),
        ('*', '*', OP_IN),
        ('/', '/', OP_IN),
        ('^', '^', OP_IN),
        ('&', '&', OP_IN),
        ('=', '=', OP_IN),
        ('>', '>', OP_IN),
        ('<', '<', OP_IN),
        ('+', '+', OP_PRE),
        ('-', '-', OP_PRE),
        ('=<', '=', OP_IN),
        ('><', '>', OP_IN),
        ('<<', '<', OP_IN),
        ('>>', '>', OP_IN),
    ])
    def test_parse_operator(self, tokenizer, formula, result, type_):
        tok = tokenizer.Tokenizer(formula)
        tok.offset = 0
        del tok.items[:]
        assert tok._parse_operator() == len(result)
        assert len(tok.items) == 1
        assert not tok.token
        token = tok.items[0]
        assert token.value == result
        assert token.type == type_
        assert token.subtype == ''

    @pytest.mark.parametrize('prefix, char, type_', [
        ('name', '(', FUNC),
        ('', '(', PAREN),
        ('', '{', ARRAY),
    ])
    def test_parse_opener(self, tokenizer, prefix, char, type_):
        tok = tokenizer.Tokenizer(prefix + char)
        del tok.items[:]
        tok.offset = len(prefix)
        if prefix:
            tok.token.append(prefix)
        assert tok._parse_opener() == 1
        assert not tok.token
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == prefix + char
        assert token.type == type_
        assert token.subtype == OPEN
        assert len(tok.token_stack) == 1
        assert tok.token_stack[0] is token

    def test_parse_opener_error(self, tokenizer):
        tok = tokenizer.Tokenizer('name{')
        tok.offset = 4
        tok.token[:] = ('name',)
        with pytest.raises(tokenizer.TokenizerError):
            tok._parse_opener()

    @pytest.mark.parametrize('formula, offset, opener', [
        ('func(a)', 6, ('func(', FUNC, OPEN)),
        ('(a)', 2, ('(', PAREN, OPEN)),
        ('{a,b,c}', 6, ('{', ARRAY, OPEN)),
    ])
    def test_parse_closer(self, tokenizer, formula, offset, opener):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        tok.token_stack.append(tokenizer.Token(*opener))
        assert tok._parse_closer() == 1
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == formula[offset]
        assert token.type == opener[1]
        assert token.subtype == CLOSE

    @pytest.mark.parametrize('formula, offset, opener', [
        ('func(a}', 6, ('func(', FUNC, OPEN)),
        ('(a}', 2, ('(', PAREN, OPEN)),
        ('{a,b,c)', 6, ('{', ARRAY, OPEN)),
    ])
    def test_parse_closer_error(self, tokenizer, formula, offset, opener):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        tok.token_stack.append(tokenizer.Token(*opener))
        with pytest.raises(tokenizer.TokenizerError):
            tok._parse_closer()

    @pytest.mark.parametrize('formula, offset, opener, type_, subtype', [
        ("{a;b}", 2, ('{', ARRAY, OPEN), SEP, ROW),
        ("{a,b}", 2, ('{', ARRAY, OPEN), SEP, ARG),
        ("(a,b)", 2, ('(', PAREN, OPEN), OP_IN, ''),
        ("FUNC(a,b)", 6, ('FUNC(', FUNC, OPEN), SEP, ARG),
        ("$A$15:$B$20,$A$1:$B$5", 11, None, OP_IN, "")
    ])
    def test_parse_separator(self, tokenizer, formula, offset, opener, type_, subtype):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        if opener:
            tok.token_stack.append(tokenizer.Token(*opener))
        assert tok._parse_separator() == 1
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == formula[offset]
        assert token.type == type_
        assert token.subtype == subtype

    @pytest.mark.parametrize('formula, offset, token, ret', [
        ('1.0E-5', 4, ['1', '.', '0', 'E'], True),
        ('1.53321E+3', 8, ['1.53321', 'E'], True),
        ('9.9E+12', 4, ['9.', '9E'], True),
        ('3E+155', 2, ['9.', '9', 'E'], True),
        ('12E+15', 3, ['12', 'E'], False),
        ('0.1E-5', 4, ['0', '.1', 'E'], False),
        ('0E+7', 2, ['0', 'E'], False),
        ('12+', 2, ['1', '2'], False),
        ('13-E+', 4, ['E'], False),
        ('+', 0, [], False),
        ('1.0e-5', 4, ['1', '.', '0', 'e'], True),
        ('1.53321e+3', 8, ['1.53321', 'e'], True),
        ('9.9e+12', 4, ['9.', '9e'], True),
        ('3e+155', 2, ['9.', '9', 'e'], True),
        ('12e+15', 3, ['12', 'e'], False),
        ('0.1e-5', 4, ['0', '.1', 'e'], False),
        ('0e+7', 2, ['0', 'e'], False),
        ('12+', 2, ['1', '2'], False),
        ('13-e+', 4, ['e'], False),
        ('+', 0, [], False),
    ])
    def test_check_scientific_notation(self, tokenizer, formula, offset, token, ret):
        tok = tokenizer.Tokenizer(formula)
        del tok.items[:]
        tok.offset = offset
        tok.token[:] = token
        assert ret is tok.check_scientific_notation()
        if ret:
            assert offset + 1 == tok.offset
            assert token == tok.token[:-1]
            assert tok.token[-1] == formula[offset]
        else:
            assert offset == tok.offset
            assert token == tok.token

    @pytest.mark.parametrize('token, can_follow, raises', [
        ('', {}, False),
        ('test', {}, True),
        ('test:', {'can_follow': ':'}, False),  # sheetname in range
        ('test', {'can_follow': ':'}, True),
        ('test!', {'can_follow': '!'}, False),  # #ERR! in defined name
        ('test', {'can_follow': '!'}, True),
    ])
    def test_assert_empty_token(self, tokenizer, token, can_follow, raises):
        tok = tokenizer.Tokenizer('')
        tok.token.extend(list(token))
        if not raises:
            try:
                tok.assert_empty_token(**can_follow)
            except tokenizer.TokenizerError:
                pytest.fail(
                    "assert_empty_token raised TokenizerError incorrectly")
        else:
            with pytest.raises(tokenizer.TokenizerError):
                tok.assert_empty_token(**can_follow)

    def test_save_token(self, tokenizer):
        tok = tokenizer.Tokenizer("")
        tok.save_token()
        assert not tok.items
        tok.token.append("test")
        tok.save_token()
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == "test"
        assert token.type == OPERAND

    @pytest.mark.parametrize('formula', [
        '=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
        "='Summary slices'!$C$3",
        '=-MAX(Pipeline!AA4:AA138)',
        '=TEXT(-S7/1000,"$#,##0""M""")',
        ("=IF(A$3<1.3E-8,\"\",IF(ISNA('External Ref'!K7),"
         '"N/A",TEXT(K7*1E+12,"0")&"bp"'),
        ('=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=' +
         'LOWER_BOUND)*($BR6>$S72123))'),
        '=(AW$4=$D7)+0%',
        "Just text",
        "123.456",
        "31/12/1999",
        "",
    ])
    def test_render(self, tokenizer, formula):
        tok = tokenizer.Tokenizer(formula)
        assert tok.render() == formula


class TestToken(object):

    def test_init(self, tokenizer):
        tokenizer.Token('val', 'type', 'subtype')

    @pytest.mark.parametrize('value, subtype', [
        ('"text"', TEXT),
        ('#REF!', ERROR),
        ('123', NUMBER),
        ('0', NUMBER),
        ('0.123', NUMBER),
        ('.123', NUMBER),
        ('1.234E5', NUMBER),
        ('1E+5', NUMBER),
        ('1.13E-55', NUMBER),
        ('TRUE', LOGICAL),
        ('FALSE', LOGICAL),
        ('A1', RANGE),
        ('ABCD12345', RANGE),
        ("'Hello world'!R123C[-12]", RANGE),
        ("[outside-workbook.xlsx]'A sheet name'!$AB$122", RANGE),
    ])
    def test_make_operand(self, tokenizer, value, subtype):
        tok = tokenizer.Token.make_operand(value)
        assert tok.value == value
        assert tok.type == OPERAND
        assert tok.subtype == subtype

    @pytest.mark.parametrize('value, type_, subtype', [
        ('{', ARRAY, OPEN),
        ('}', ARRAY, CLOSE),
        ('(', PAREN, OPEN),
        (')', PAREN, CLOSE),
        ('FUNC(', FUNC, OPEN),
    ])
    def test_make_subexp(self, tokenizer, value, type_, subtype):
        tok = tokenizer.Token.make_subexp(value)
        assert tok.value == value
        assert tok.type == type_
        assert tok.subtype == subtype

    def test_make_subexp_func(self, tokenizer):
        tok = tokenizer.Token.make_subexp(')', True)
        assert tok.value == ')'
        assert tok.type == FUNC
        assert tok.subtype == CLOSE

        tok = tokenizer.Token.make_subexp('TEST(', True)
        assert tok.value == 'TEST('
        assert tok.type == FUNC
        assert tok.subtype == OPEN

    @pytest.mark.parametrize('token, close_val', [
        (('(', PAREN, OPEN), ')'),
        (('{', ARRAY, OPEN), '}'),
        (('FUNC(', FUNC, OPEN), ')'),
    ])
    def test_get_closer(self, tokenizer, token, close_val):
        closer = tokenizer.Token(*token).get_closer()
        assert closer.value == close_val
        assert closer.type == token[1]
        assert closer.subtype == CLOSE

    def test_make_separator(self, tokenizer):
        token = tokenizer.Token.make_separator(',')
        assert token.value == ','
        assert token.type == SEP
        assert token.subtype == ARG

        token = tokenizer.Token.make_separator(';')
        assert token.value == ';'
        assert token.type == SEP
        assert token.subtype == ROW

    @pytest.mark.parametrize('formula, tokens', [
        ("SUM(Inputs!$W$111:'Input 1'!W111)",
         [("SUM(Inputs!$W$111:'Input 1'!W111)", 'LITERAL', '')]),

        ("=SUM('Inputs 1'!$W$111:'Input 1'!W111)",
         [('SUM(', 'FUNC', 'OPEN'),
          ("'Inputs 1'!$W$111:'Input 1'!W111", 'OPERAND', 'RANGE'),
          (')', 'FUNC', 'CLOSE')]),

        ("=SUM(Inputs!$W$111:'Input 1'!W111)",
         [('SUM(', 'FUNC', 'OPEN'),
          ("Inputs!$W$111:'Input 1'!W111", 'OPERAND', 'RANGE'),
          (')', 'FUNC', 'CLOSE')]),

        ("=SUM(Inputs!$W$111:'Input ''\"1'!W111)",
         [('SUM(', 'FUNC', 'OPEN'),
          ("Inputs!$W$111:'Input ''\"1'!W111", 'OPERAND', 'RANGE'),
          (')', 'FUNC', 'CLOSE')]),

        ("=SUM(Inputs!$W$111:Input1!W111)",
         [('SUM(', 'FUNC', 'OPEN'),
          ('Inputs!$W$111:Input1!W111', 'OPERAND', 'RANGE'),
          (')', 'FUNC', 'CLOSE')]),
    ])
    def test_parse_quoted_sheet_name_in_range(self, tokenizer, formula, tokens):
        tok = tokenizer.Tokenizer(formula)
        result = [(token.value, token.type, token.subtype)
                  for token in tok.items]
        assert result == tokens
        assert tok.render() == formula
