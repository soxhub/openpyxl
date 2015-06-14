from __future__ import absolute_import

import pytest

@pytest.fixture
def Tokenizer():
    from ..tokenizer import Tokenizer
    return Tokenizer

@pytest.fixture
def Token():
    from ..tokenizer import Token
    return Token


class TestTokenizerRegexes(object):

    def check_regex(self, regex, cases):
        for string, expected in cases:
            if expected is None:
                assert not regex.match(string)
            else:
                assert regex.match(string)
                assert regex.match(string).group(0) == expected

    def test_scientific_re(self, Tokenizer):
        positive = [
            '1.0E',
            '1.53321E',
            '9.999E',
            '3E',
        ]
        negative = [
            '12E',
            '0.1E',
            '0E',
            '',
            'E',
        ]
        regex = Tokenizer.SN_RE
        for string in positive:
            assert bool(regex.match(string))
        for string in negative:
            assert not bool(regex.match(string))

    def test_whitespace_re(self, Tokenizer):
        cases = [
            (' ', ' '),
            (' *', ' '),
            ('     ', '     '),
            ('     a', '     '),
            ('   ', '   '),
            ('   +', '   '),
            ('', None),
            ('*', None),
        ]
        self.check_regex(Tokenizer.WSPACE_RE, cases)

    def test_string_re(self, Tokenizer):
        cases = [
            ('"spamspamspam"', '"spamspamspam"'),
            ('"this is "" a test "" "', '"this is "" a test "" "'),
            ('""', '""'),
            ('"spam and ""cheese"""+"ignore"', ('"spam and ""cheese"""')),
            ('\'"spam and ""cheese"""+"ignore"', None),
            ('"oops ""', None),
        ]
        regex = Tokenizer.STRING_REGEXES['"']
        self.check_regex(regex, cases)

    def test_link_re(self, Tokenizer):
        cases = [
            ("'spam and ham'", "'spam and ham'"),
            ("'double'' triple''' quadruple ''''", "'double'' triple'''"),
            ("'sextuple '''''' and septuple''''''' and more",
             "'sextuple '''''' and septuple'''''''",),
            ("''", "''"),
            ("'oops ''", None),
            ("gunk'hello world'", None),
        ]
        regex = Tokenizer.STRING_REGEXES["'"]
        self.check_regex(regex, cases)


class TestTokenizer(object):

    def test_init(self, Tokenizer):
        tok = Tokenizer("abcdefg")
        assert tok.offset == 0
        tok = Tokenizer("=abcdefg")
        assert tok.offset == 0

    def test_parse(self, Tokenizer, Token):
        cases = [
            (u'=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
             [(u'IF(', Token.FUNC, Token.OPEN),
             (u'A$3', Token.OPERAND, Token.RANGE),
             (u'<', Token.OP_IN, u""),
             (u'40', Token.OPERAND, Token.NUMBER),
             (u'%', Token.OP_POST, u""),
             (u',', Token.SEP, Token.ARG),
             (u'""', Token.OPERAND, Token.TEXT),
             (u',', Token.SEP, Token.ARG),
             (u'INDEX(', Token.FUNC, Token.OPEN),
             (u'Pipeline!B$4:B$138', Token.OPERAND, Token.RANGE),
             (u',', Token.SEP, Token.ARG),
             (u'#REF!', Token.OPERAND, Token.ERROR),
             (u')', Token.FUNC, Token.CLOSE),
             (u')', Token.FUNC, Token.CLOSE)]),

            (u"='Summary slices'!$C$3",
             [(u"'Summary slices'!$C$3", Token.OPERAND, Token.RANGE)]),

            (u'=-MAX(Pipeline!AA4:AA138)',
             [(u"-", Token.OP_PRE, u""),
              (u'MAX(', Token.FUNC, Token.OPEN),
              (u'Pipeline!AA4:AA138', Token.OPERAND, Token.RANGE),
              (u')', Token.FUNC, Token.CLOSE)]),

            (u'=TEXT(-S7/1000,"$#,##0""M""")',
             [(u'TEXT(', Token.FUNC, Token.OPEN),
              (u'-', Token.OP_PRE, u""),
              (u'S7', Token.OPERAND, Token.RANGE),
              (u'/', Token.OP_IN, u""),
              (u'1000', Token.OPERAND, Token.NUMBER),
              (u',', Token.SEP, Token.ARG),
              (u'"$#,##0""M"""', Token.OPERAND, Token.TEXT),
              (u')', Token.FUNC, Token.CLOSE)]),

            (u"=IF(A$3<1.3E-8,\"\",IF(ISNA('External Ref'!K7)," +
             u'"N/A",TEXT(K7*1E+12,"0")&"bp"',
             [(u'IF(', Token.FUNC, Token.OPEN),
              (u'A$3', Token.OPERAND, Token.RANGE),
              (u'<', Token.OP_IN, u""),
              (u'1.3E-8', Token.OPERAND, Token.NUMBER),
              (u',', Token.SEP, Token.ARG),
              (u'""', Token.OPERAND, Token.TEXT),
              (u',', Token.SEP, Token.ARG),
              (u'IF(', Token.FUNC, Token.OPEN),
              (u'ISNA(', Token.FUNC, Token.OPEN),
              (u"'External Ref'!K7", Token.OPERAND, Token.RANGE),
              (u')', Token.FUNC, Token.CLOSE),
              (u',', Token.SEP, Token.ARG),
              (u'"N/A"', Token.OPERAND, Token.TEXT),
              (u',', Token.SEP, Token.ARG),
              (u'TEXT(', Token.FUNC, Token.OPEN),
              (u'K7', Token.OPERAND, Token.RANGE),
              (u'*', Token.OP_IN, u""),
              (u'1E+12', Token.OPERAND, Token.NUMBER),
              (u',', Token.SEP, Token.ARG),
              (u'"0"', Token.OPERAND, Token.TEXT),
              (u')', Token.FUNC, Token.CLOSE),
              (u'&', Token.OP_IN, u""),
              (u'"bp"', Token.OPERAND, Token.TEXT)]),

            (u'=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=' +
             u'LOWER_BOUND)*($BR6>$S72123))',
             [(u"+", Token.OP_PRE, u""),
              (u'IF(', Token.FUNC, Token.OPEN),
              (u'A$3', Token.OPERAND, Token.RANGE),
              (u'<>', Token.OP_IN, u""),
              (u'$B7', Token.OPERAND, Token.RANGE),
              (u',', Token.SEP, Token.ARG),
              (u'""', Token.OPERAND, Token.TEXT),
              (u',', Token.SEP, Token.ARG),
              (u'(', Token.PAREN, Token.OPEN),
              (u'MIN(', Token.FUNC, Token.OPEN),
              (u'IF(', Token.FUNC, Token.OPEN),
              (u'{', Token.ARRAY, Token.OPEN),
              (u'TRUE', Token.OPERAND, Token.LOGICAL),
              (u',', Token.SEP, Token.ARG),
              (u' ', Token.WSPACE, u''),
              (u'FALSE', Token.OPERAND, Token.LOGICAL),
              (u';', Token.SEP, Token.ROW),
              (u'1', Token.OPERAND, Token.NUMBER),
              (u',', Token.SEP, Token.ARG),
              (u'2', Token.OPERAND, Token.NUMBER),
              (u'}', Token.ARRAY, Token.CLOSE),
              (u',', Token.SEP, Token.ARG),
              (u'A6:B6', Token.OPERAND, Token.RANGE),
              (u',', Token.SEP, Token.ARG),
              (u'$S7', Token.OPERAND, Token.RANGE ),
              (u')', Token.FUNC, Token.CLOSE),
              (u')', Token.FUNC, Token.CLOSE),
              (u'>=', Token.OP_IN, u''),
              (u'LOWER_BOUND', Token.OPERAND, Token.RANGE),
              (u')', Token.PAREN, Token.CLOSE),
              (u'*', Token.OP_IN, u''),
              (u'(', Token.PAREN, Token.OPEN),
              (u'$BR6', Token.OPERAND, Token.RANGE),
              (u'>', Token.OP_IN, u''),
              (u'$S72123', Token.OPERAND, Token.RANGE),
              (u')', Token.PAREN, Token.CLOSE),
              (u')', Token.FUNC, Token.CLOSE)]),

            (u'=(AW$4=$D7)+0%',
             [(u'(', Token.PAREN, Token.OPEN),
              (u'AW$4', Token.OPERAND, Token.RANGE),
              (u'=', Token.OP_IN, u''),
              (u'$D7', Token.OPERAND, Token.RANGE),
              (u')', Token.PAREN, Token.CLOSE),
              (u'+', Token.OP_IN, u''),
              (u'0', Token.OPERAND, Token.NUMBER),
              (u'%', Token.OP_POST, u'')]),

            (u'=$A:$A,$C:$C',
             [(u'$A:$A', Token.OPERAND, Token.RANGE),
              (u',', Token.OP_IN, u""),
              (u'$C:$C', Token.OPERAND, Token.RANGE)]),

            (u"Just text", [(u"Just text", Token.LITERAL, u"")]),
            (u"123.456", [(u"123.456", Token.LITERAL, u"")]),
            (u"31/12/1999", [(u"31/12/1999", Token.LITERAL, u"")]),
            (u"", []),
        ]
        for formula, tokens in cases:
            tok = Tokenizer(formula)
            tok.parse()
            result = [(token.value, token.type, token.subtype)
                      for token in tok.items]
            assert result == tokens

    def test_parse_string(self, Tokenizer, Token):
        cases = [
            (u'"spamspamspam"spam', 0, u'"spamspamspam"'),
            (u'"this is "" a test "" "test', 0, u'"this is "" a test "" "'),
            (u'""', 0, u'""'),
            (u'a"bcd""efg"hijk', 1, u'"bcd""efg"'),
            (u'"oops ""', 0, None),
            (u"'spam and ham'", 0, u"'spam and ham'"),
            (u"'double'' triple''' quad ''''", 0, u"'double'' triple'''"),
            (u"123'sextuple '''''' and septuple''''''' and more", 3,
             u"'sextuple '''''' and septuple'''''''"),
            (u"''", 0, u"''"),
            (u"'oops ''", 0, None),
        ]
        tok = Tokenizer(u'')
        for formula, offset, result in cases:
            tok.offset = offset
            tok.formula = formula
            if result is None:
                with pytest.raises(TokenizerError):
                    tok.parse_string()
                continue
            assert tok.parse_string() == len(result)
            if formula[offset] == '"':
                token = tok.items[0]
                assert token.value == result
                assert token.type == Token.OPERAND
                assert token.subtype == Token.TEXT
                assert not tok.token
            else:
                assert not tok.items
                assert tok.token[0] == result
                assert len(tok.token) == 1
            del tok.items[:], tok.token[:], tok.token_stack[:]

    def test_parse_brackets(self, Tokenizer):
        cases = [
            ('[abc]def', 0, '[abc]'),
            ('[]abcdef', 0, '[]'),
            ('[abcdef]', 0, '[abcdef]'),
            ('a[bcd]ef', 1, '[bcd]'),
            ('ab[cde]f', 2, '[cde]'),
        ]
        tok = Tokenizer('')
        for formula, offset, result in cases:
            tok.offset = offset
            tok.formula = formula
            assert tok.parse_brackets() == len(result)
            assert not tok.items
            assert tok.token[0] == result
            assert len(tok.token) == 1
            del tok.items[:], tok.token[:], tok.token_stack[:]
        with pytest.raises(TokenizerError):
            tok.formula = '[unfinished business'
            tok.offset = 0
            tok.parse_brackets()

    def test_parse_error(self, Tokenizer, Token):
        errors = (u"#NULL!", u"#DIV/0!", u"#VALUE!", u"#REF!", u"#NAME?",
                  u"#NUM!", u"#N/A")
        for error in errors:
            tok = Tokenizer(error)
            assert tok.parse_error() == len(error)
            assert len(tok.items) == 1
            assert not tok.token
            token = tok.items[0]
            assert token.value == error
            assert token.type == Token.OPERAND
            assert token.subtype == Token.ERROR

        with pytest.raises(TokenizerError):
            tok = Tokenizer(u"#NotAnError")
            tok.parse_error()

    def test_parse_whitespace(self, Tokenizer, Token):
        for i in range(1, 10):
            tok = Tokenizer(u" " * i)
            assert tok.parse_whitespace() == i
            assert len(tok.items) == 1
            token = tok.items[0]
            assert token.value == u" "
            assert token.type == Token.WSPACE
            assert token.subtype == u""
            assert not tok.token

    def test_parse_operator(self, Tokenizer, Token):
        cases = [
            (u'>=', u'>=', Token.OP_IN),
            (u'<=', u'<=', Token.OP_IN),
            (u'<>', u'<>', Token.OP_IN),
            (u'%', u'%', Token.OP_POST),
            (u'*', u'*', Token.OP_IN),
            (u'/', u'/', Token.OP_IN),
            (u'^', u'^', Token.OP_IN),
            (u'&', u'&', Token.OP_IN),
            (u'=', u'=', Token.OP_IN),
            (u'>', u'>', Token.OP_IN),
            (u'<', u'<', Token.OP_IN),
            (u'+', u'+', Token.OP_PRE),
            (u'-', u'-', Token.OP_PRE),
            (u'=<', u'=', Token.OP_IN),
            (u'><', u'>', Token.OP_IN),
            (u'<<', u'<', Token.OP_IN),
            (u'>>', u'>', Token.OP_IN),
        ]
        for formula, result, type_ in cases:
            tok = Tokenizer(formula)
            assert tok.parse_operator() == len(result)
            assert len(tok.items) == 1
            assert not tok.token
            token = tok.items[0]
            assert token.value == result
            assert token.type == type_
            assert token.subtype == u''

    def test_parse_opener(self, Tokenizer, Token):
        cases = [
            (u'name', u'(', Token.FUNC),
            (u'', u'(', Token.PAREN),
            (u'', u'{', Token.ARRAY),
        ]
        for prefix, char, type_ in cases:
            tok = Tokenizer(prefix + char)
            tok.offset = len(prefix)
            if prefix:
                tok.token.append(prefix)
            assert tok.parse_opener() == 1
            assert not tok.token
            assert len(tok.items) == 1
            token = tok.items[0]
            assert token.value == prefix + char
            assert token.type == type_
            assert token.subtype == Token.OPEN
            assert len(tok.token_stack) == 1
            assert tok.token_stack[0] is token
        with pytest.raises(TokenizerError):
            tok = Tokenizer('name{')
            tok.offset = 4
            tok.token.append('name')
            tok.parse_opener()

    def test_parse_closer(self, Tokenizer, Token):
        cases = [
            #  formula offset top of token_stack
            (u'func(a)', 6, Token('func(', Token.FUNC, Token.OPEN)),
            (u'(a)', 2, Token('(', Token.PAREN, Token.OPEN)),
            (u'{a,b,c}', 6, Token('{', Token.ARRAY, Token.OPEN)),
        ]
        for formula, offset, opener in cases:
            tok = Tokenizer(formula)
            tok.offset = offset
            tok.token_stack.append(opener)
            assert tok.parse_closer() == 1
            assert len(tok.items) == 1
            token = tok.items[0]
            assert token.value == formula[offset]
            assert token.type == opener.type
            assert token.subtype == Token.CLOSE
        cases = [
            (u'func(a}', 6, Token('func(', Token.FUNC, Token.OPEN)),
            (u'(a}', 2, Token('(', Token.PAREN, Token.OPEN)),
            (u'{a,b,c)', 6, Token('{', Token.ARRAY, Token.OPEN)),
        ]
        for formula, offset, opener in cases:
            tok = Tokenizer(formula)
            tok.offset = offset
            tok.token_stack.append(opener)
            with pytest.raises(TokenizerError):
                tok.parse_closer()

    def test_parse_separator(self, Tokenizer, Token):
        cases = [
            (u"{a;b}", 2, Token('{', Token.ARRAY, Token.OPEN),
             Token.SEP, Token.ROW),
            (u"{a,b}", 2, Token('{', Token.ARRAY, Token.OPEN),
             Token.SEP, Token.ARG),
            (u"(a,b)", 2, Token('(', Token.PAREN, Token.OPEN),
             Token.OP_IN, u''),
            (u"FUNC(a,b)", 6, Token('FUNC(', Token.FUNC, Token.OPEN),
             Token.SEP, Token.ARG),
            (u"$A$15:$B$20,$A$1:$B$5", 11, None, Token.OP_IN, u"")
        ]
        for formula, offset, opener, type_, subtype in cases:
            tok = Tokenizer(formula)
            tok.offset = offset
            if opener:
                tok.token_stack.append(opener)
            assert tok.parse_separator() == 1
            assert len(tok.items) == 1
            token = tok.items[0]
            assert token.value == formula[offset]
            assert token.type == type_
            assert token.subtype == subtype

    def test_check_scientific_notation(self, Tokenizer):
        cases = [
            # formula offset  token-pre         retval
            (u'1.0E-5', 4, ['1', '.', '0', 'E'], True),
            (u'1.53321E+3', 8, ['1.53321', 'E'], True),
            (u'9.9E+12', 4, ['9.', '9E'], True),
            (u'3E+155', 2, ['9.', '9', 'E'], True),
            (u'12E+15', 3, ['12', 'E'], False),
            (u'0.1E-5', 4, ['0', '.1', 'E'], False),
            (u'0E+7', 2, ['0', 'E'], False),
            (u'12+', 2, ['1', '2'], False),
            (u'13-E+', 4, ['E'], False),
            (u'+', 0, [], False),
        ]
        for formula, offset, token, ret in cases:
            tok = Tokenizer(formula)
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

    def test_assert_empty_token(self, Tokenizer):
        tok = Tokenizer(u"")
        try:
            tok.assert_empty_token()
        except TokenizerError:
            pytest.fail(
                u"assert_empty_token raised TokenizerError incorrectly")
        tok.token.append(u"test")
        with pytest.raises(TokenizerError):
            tok.assert_empty_token()

    def test_save_token(self, Tokenizer, Token):
        tok = Tokenizer(u"")
        tok.save_token()
        assert not tok.items
        tok.token.append(u"test")
        tok.save_token()
        assert len(tok.items) == 1
        token = tok.items[0]
        assert token.value == u"test"
        assert token.type == Token.OPERAND

    def test_render(self, Tokenizer):
        cases = [
            u'=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
            u"='Summary slices'!$C$3",
            u'=-MAX(Pipeline!AA4:AA138)',
            u'=TEXT(-S7/1000,"$#,##0""M""")',
            (u"=IF(A$3<1.3E-8,\"\",IF(ISNA('External Ref'!K7),"
             u'"N/A",TEXT(K7*1E+12,"0")&"bp"'),
            (u'=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=' +
             u'LOWER_BOUND)*($BR6>$S72123))'),
            u'=(AW$4=$D7)+0%',
            u"Just text",
            u"123.456",
            u"31/12/1999",
            u"",
        ]
        for formula in cases:
            tok = Tokenizer(formula)
            tok.parse()
            assert tok.render() == formula


class TestToken(object):

    def test_init(self, Token):
        Token(u'val', u'type', u'subtype')

    def test_make_operand(self, Token):
        cases = [
            (u'"text"', Token.TEXT),
            (u'#REF!', Token.ERROR),
            (u'123', Token.NUMBER),
            (u'0', Token.NUMBER),
            (u'0.123', Token.NUMBER),
            (u'.123', Token.NUMBER),
            (u'1.234E5', Token.NUMBER),
            (u'1E+5', Token.NUMBER),
            (u'1.13E-55', Token.NUMBER),
            (u'TRUE', Token.LOGICAL),
            (u'FALSE', Token.LOGICAL),
            (u'A1', Token.RANGE),
            (u'ABCD12345', Token.RANGE),
            (u"'Hello world'!R123C[-12]", Token.RANGE),
            (u"[outside-workbook.xlsx]'A sheet name'!$AB$122", Token.RANGE),
        ]
        for value, subtype in cases:
            tok = Token.make_operand(value)
            assert tok.value == value
            assert tok.type == Token.OPERAND
            assert tok.subtype == subtype

    def test_make_subexp(self, Token):
        cases = [
            (u'{', Token.ARRAY, Token.OPEN),
            (u'}', Token.ARRAY, Token.CLOSE),
            (u'(', Token.PAREN, Token.OPEN),
            (u')', Token.PAREN, Token.CLOSE),
            (u'FUNC(', Token.FUNC, Token.OPEN),
        ]
        for value, type_, subtype in cases:
            tok = Token.make_subexp(value)
            assert tok.value == value
            assert tok.type == type_
            assert tok.subtype == subtype

        tok = Token.make_subexp(')', True)
        assert tok.value == ')'
        assert tok.type == Token.FUNC
        assert tok.subtype == Token.CLOSE

        tok = Token.make_subexp('TEST(', True)
        assert tok.value == 'TEST('
        assert tok.type == Token.FUNC
        assert tok.subtype == Token.OPEN

    def test_get_closer(self, Token):
        cases = [
            (Token(u'(', Token.PAREN, Token.OPEN), u')'),
            (Token(u'{', Token.ARRAY, Token.OPEN), u'}'),
            (Token(u'FUNC(', Token.FUNC, Token.OPEN), u')'),
        ]
        for token, close_val in cases:
            closer = token.get_closer()
            assert closer.value == close_val
            assert closer.type == token.type
            assert closer.subtype == Token.CLOSE

    def test_make_separator(self, Token):
        token = Token.make_separator(u',')
        assert token.value == u','
        assert token.type == Token.SEP
        assert token.subtype == Token.ARG

        token = Token.make_separator(u';')
        assert token.value == u';'
        assert token.type == Token.SEP
        assert token.subtype == Token.ROW
