u"""
This module contains a tokenizer for Excel formulae.

The tokenizer is based on the Javascript tokenizer found at
http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html written by Eric
Bachtal (ewbachtal -AT- ewbi -DOT- com), who provided the software AS IS,
without warranty of any kind.

"""

import re


class TokenizerError(Exception):
    u"Base class for all Tokenizer errors."


class Tokenizer(object):

    u"""
    A tokenizer for Excel worksheet formulae.

    Converts a unicode string representing an Excel formula (in A1 notation)
    into a sequence of `Token` objects.

    `formula`: The unicode string to tokenize

    Tokenizer defines a method `.parse()` to parse the formula into tokens,
    which can then be accessed through the `.items` attribute.

    """

    SN_RE = re.compile(u"^[1-9](\\.[0-9]+)?E$")  # Scientific notation
    WSPACE_RE = re.compile(u" +")
    STRING_REGEXES = {
        # Inside a string, all characters are treated as literals, except for
        # the quote character used to start the string. That character, when
        # doubled is treated as a single character in the string. If an
        # unmatched quote appears, the string is terminated.
        u'"': re.compile(u'"(?:[^"]*"")*[^"]*"(?!")'),
        u"'": re.compile(u"'(?:[^']*'')*[^']*'(?!')"),
    }
    ERROR_CODES = (u"#NULL!", u"#DIV/0!", u"#VALUE!", u"#REF!", u"#NAME?",
                   u"#NUM!", u"#N/A")
    TOKEN_ENDERS = u',;}) +-*/^&=><%'  # Each of these characters, marks the
                                       # end of an operand token

    def __init__(self, formula):
        self.formula = formula
        self.items = []
        self.token_stack = []  # Used to keep track of arrays, functions, and
                               # parentheses
        self.offset = 0  # How many chars have we read
        self.token = []  # Used to build up token values char by char

    def parse(self):
        u"Populate self.items with the tokens from the formula."
        if not self.formula:
            return
        elif self.formula[0] == '=':
            self.offset += 1
        else:
            self.items.append(Token(self.formula, Token.LITERAL))
            return
        consumers = (
            (u'"\'', self.parse_string),
            (u'[', self.parse_brackets),
            (u'#', self.parse_error),
            (u' ', self.parse_whitespace),
            (u'+-*/^&=><%', self.parse_operator),
            (u'{(', self.parse_opener),
            (u')}', self.parse_closer),
            (u';,', self.parse_separator),
        )
        dispatcher = {}  # maps chars to the specific parsing function
        for chars, consumer in consumers:
            dispatcher.update(dict.fromkeys(chars, consumer))
        while self.offset < len(self.formula):
            if self.check_scientific_notation():  # May consume one character
                continue
            curr_char = self.formula[self.offset]
            if curr_char in self.TOKEN_ENDERS:
                self.save_token()
            if curr_char in dispatcher:
                self.offset += dispatcher[curr_char]()
            else:
                # TODO: this can probably be sped up using a regex to get to
                # the next interesting character
                self.token.append(curr_char)
                self.offset += 1
        self.save_token()

    def parse_string(self):
        u"""
        Parse a "-delimited string or '-delimited link.

        The offset must be pointing to either a single quote ("'") or double
        quote ('"') character. The strings are parsed according to Excel
        rules where to escape the delimiter you just double it up. E.g.,
        "abc""def" in Excel is parsed as 'abc"def' in Python.

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        self.assert_empty_token()
        delim = self.formula[self.offset]
        assert delim in (u'"', u"'")
        regex = self.STRING_REGEXES[delim]
        match = regex.match(self.formula[self.offset:])
        if match is None:
            subtype = u"string" if delim == u'"' else u'link'
            raise TokenizerError(
                u"Reached end of formula while parsing %s in %s" %
                (subtype, self.formula))
        match = match.group(0)
        if delim == u'"':
            self.items.append(Token.make_operand(match))
        else:
            self.token.append(match)
        return len(match)

    def parse_brackets(self):
        u"""
        Consume all the text between square brackets [].

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] == u'['
        right = self.formula.find(u']', self.offset) + 1
        if right == 0:
            raise TokenizerError(
                "Encountered unmatched '[' in %s" % self.formula)
        self.token.append(self.formula[self.offset: right])
        return right - self.offset

    def parse_error(self):
        u"""
        Consume the text following a '#' as an error.

        Looks for a match in self.ERROR_CODES and returns the number of
        characters matched. (Does not update self.offset)

        """
        self.assert_empty_token()
        assert self.formula[self.offset] == u'#'
        subformula = self.formula[self.offset:]
        for err in self.ERROR_CODES:
            if subformula.startswith(err):
                self.items.append(Token.make_operand(err))
                return len(err)
        raise TokenizerError(
            u"Invalid error code at position %d in '%s'" %
            (self.offset, self.formula))

    def parse_whitespace(self):
        u"""
        Consume a string of consecutive spaces.

        Returns the number of spaces found. (Does not update self.offset).

        """
        assert self.formula[self.offset] == u' '
        self.items.append(Token(u' ', Token.WSPACE))
        return self.WSPACE_RE.match(self.formula[self.offset:]).end()

    def parse_operator(self):
        u"""
        Consume the characters constituting an operator.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        if self.formula[self.offset:self.offset + 2] in (u'>=', u'<=', u'<>'):
            self.items.append(Token(
                self.formula[self.offset:self.offset + 2],
                Token.OP_IN
            ))
            return 2
        curr_char = self.formula[self.offset]  # guaranteed to be 1 char
        assert curr_char in u'%*/^&=><+-'
        if curr_char == u'%':
            token = Token(u'%', Token.OP_POST)
        elif curr_char in u"*/^&=><":
            token = Token(curr_char, Token.OP_IN)
        # From here on, curr_char is guaranteed to be in '+-'
        elif not self.items:
            token = Token(curr_char, Token.OP_PRE)
        else:
            prev = self.items[-1]
            is_infix = (
                prev.subtype == Token.CLOSE
                or prev.type == Token.OP_POST
                or prev.type == Token.OPERAND
            )
            if is_infix:
                token = Token(curr_char, Token.OP_IN)
            else:
                token = Token(curr_char, Token.OP_PRE)
        self.items.append(token)
        return 1

    def parse_opener(self):
        u"""
        Consumes a ( or { character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] in (u'(', u'{')
        if self.formula[self.offset] == u'{':
            self.assert_empty_token()
            token = Token.make_subexp(u"{")
        elif self.token:
            token_value = u"".join(self.token) + u'('
            del self.token[:]
            token = Token.make_subexp(token_value)
        else:
            token = Token.make_subexp(u"(")
        self.items.append(token)
        self.token_stack.append(token)
        return 1

    def parse_closer(self):
        u"""
        Consumes a } or ) character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] in (u')', u'}')
        token = self.token_stack.pop().get_closer()
        if token.value != self.formula[self.offset]:
            raise TokenizerError(
                u"Mismatched ( and { pair in '%s'" % self.formula)
        self.items.append(token)
        return 1

    def parse_separator(self):
        u"""
        Consumes a ; or , character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        curr_char = self.formula[self.offset]
        assert curr_char in (u';', u',')
        if curr_char == u';':
            token = Token.make_separator(u";")
        else:
            try:
                top_type = self.token_stack[-1].type
            except IndexError:
                token = Token(u",", Token.OP_IN)  # Range Union operator
            else:
                if top_type == Token.PAREN:
                    token = Token(u",", Token.OP_IN)  # Range Union operator
                else:
                    token = Token.make_separator(u",")
        self.items.append(token)
        return 1

    def check_scientific_notation(self):
        """
        Consumes a + or - character if part of a number in sci. notation.

        Returns True if the character was consumed and self.offset was
        updated, False otherwise.

        """
        curr_char = self.formula[self.offset]
        if (curr_char in u'+-'
                and len(self.token) >= 1
                and self.SN_RE.match(u"".join(self.token))):
            self.token.append(curr_char)
            self.offset += 1
            return True
        return False

    def assert_empty_token(self):
        """
        Ensure that there's no token currently being parsed.

        If there are unconsumed token contents, it means we hit an unexpected
        token transition. In this case, we raise a TokenizerError

        """
        if self.token:
            raise TokenizerError(
                u"Unexpected character at position %d in '%s'" %
                (self.offset, self.formula))

    def save_token(self):
        u"""If there's a token being parsed, add it to the item list."""
        if self.token:
            self.items.append(Token.make_operand(u"".join(self.token)))
            del self.token[:]

    def render(self):
        u"Convert the parsed tokens back to a string."
        if not self.items:
            return u""
        elif self.items[0].type == Token.LITERAL:
            return self.items[0].value
        return u"=" + u"".join(token.value for token in self.items)


class Token(object):

    u"""
    A token in an Excel formula.

    Tokens have three attributes:

    * `value`: The string value parsed that led to this token
    * `type`: A string identifying the type of token
    * `subtype`: A string identifying subtype of the token (optional, and
                 defaults to u"")

    """

    __slots__ = ['value', 'type', 'subtype']

    LITERAL = u"LITERAL"
    OPERAND = u"OPERAND"
    FUNC = u"FUNC"
    ARRAY = u"ARRAY"
    PAREN = u"PAREN"
    SEP = u"SEP"
    OP_PRE = u"OPERATOR-PREFIX"
    OP_IN = u"OPERATOR-INFIX"
    OP_POST = u"OPERATOR-POSTFIX"
    WSPACE = u"WHITE-SPACE"

    def __init__(self, value, type_, subtype=u""):
        self.value = value
        self.type = type_
        self.subtype = subtype

    def __repr__(self):  # pragma: no cover
        if self.subtype:
            subtype = ", subtype=%s" % self.subtype
        else:
            subtype = ""
        return "<%s: value=%s, type=%s%s>" % (
            self.__class__.__name__, self.value, self.type, subtype)

    # Literal operands:
    #
    # Literal operands are always of type 'OPERAND' and can be of subtype
    # 'TEXT' (for text strings), 'NUMBER' (for all numeric types), 'LOGICAL'
    # (for TRUE and FALSE), 'ERROR' (for literal error values), or 'RANGE'
    # (for all range references).

    TEXT = u'TEXT'
    NUMBER = u'NUMBER'
    LOGICAL = u'LOGICAL'
    ERROR = u'ERROR'
    RANGE = u'RANGE'

    @classmethod
    def make_operand(cls, value):
        u"Create an operand token."
        if value.startswith('"'):
            subtype = cls.TEXT
        elif value.startswith('#'):
            subtype = cls.ERROR
        elif value in (u'TRUE', u'FALSE'):
            subtype = cls.LOGICAL
        else:
            try:
                float(value)
                subtype = cls.NUMBER
            except ValueError:
                subtype = cls.RANGE
        return cls(value, cls.OPERAND, subtype)


    # Subexpresssions
    #
    # There are 3 types of `Subexpressions`: functions, array literals, and
    # parentheticals. Subexpressions have 'OPEN' and 'CLOSE' tokens. 'OPEN'
    # is used when parsing the initital expression token (i.e., '(' or '{')
    # and 'CLOSE' is used when parsing the closing expression token ('}' or
    # ')').

    OPEN = u"OPEN"
    CLOSE = u"CLOSE"

    @classmethod
    def make_subexp(cls, value, func=False):
        """
        Create a subexpression token.

        `value`: The value of the token
        `func`: If True, force the token to be of type FUNC

        """
        assert value[-1] in (u'{', u'}', u'(', u')')
        if func:
            assert re.match(u'.+\\(|\\)', value)
            type_ = Token.FUNC
        elif value in u'{}':
            type_ = Token.ARRAY
        elif value in u'()':
            type_ = Token.PAREN
        else:
            type_ = Token.FUNC
        subtype = cls.CLOSE if value in u')}' else cls.OPEN
        return cls(value, type_, subtype)

    def get_closer(self):
        u"Return a closing token that matches this token's type."
        assert self.type in (self.FUNC, self.ARRAY, self.PAREN)
        assert self.subtype == self.OPEN
        value = "}" if self.type == self.ARRAY else ")"
        return self.make_subexp(value, func=self.type == self.FUNC)

    # Separator tokens
    #
    # Argument separators always have type 'SEP' and can have one of two
    # subtypes: 'ARG', 'ROW'. 'ARG' is used for the ',' token, when used to
    # delimit either function arguments or array elements. 'ROW' is used for
    # the ';' token, which is always used to delimit rows in an array
    # literal.

    ARG = u"ARG"
    ROW = u"ROW"

    @classmethod
    def make_separator(cls, value):
        u"Create a separator token"
        assert value in (u',', u';')
        subtype = cls.ARG if value == u',' else cls.ROW
        return cls(value, cls.SEP, subtype)
