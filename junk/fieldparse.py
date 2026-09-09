"""
fieldparse - declarative parsing of Word/RTF field instruction text.

An ``ArgumentParser``-alike for the contents of ``{\\*\\fldinst ...}`` groups.
You declare a field's positionals and switches once; the parser handles
tokenizing, quoting, switch arity, case folding and errors.

    >>> p = FieldParser('INCLUDEPICTURE')
    >>> p.add_argument('path')
    >>> p.add_switch('d', 'store_true', help="don't store image data")
    >>> p.add_switch('c', metavar='converter')
    >>> v = p.parse(r'INCLUDEPICTURE "C:\\Images\\logo.png" \\d \\* MERGEFORMAT')
    >>> v.path
    'C:\\\\Images\\\\logo.png'
    >>> v.d, v.format
    (True, ['MERGEFORMAT'])

Ready-made parsers for the common fields live in the companion module
``fields``, one module-level constant per field::

    >>> from fields import INCLUDEPICTURE
    >>> INCLUDEPICTURE.parse(r'INCLUDEPICTURE x \\d').d
    True

Unlike ``argparse`` this never touches ``sys.argv``, never writes to stderr
and never calls ``sys.exit``; malformed input raises ``FieldSyntaxError``.
"""

from __future__ import annotations

import re
from enum import Enum
from typing import Any, Iterator

__all__ = [
    'FieldError', 'FieldSyntaxError', 'FieldDefinitionError',
    'TokenKind', 'ArgKind', 'Action',
    'NestedField', 'Token', 'tokenize', 'unescape_rtf',
    'FieldParser', 'FieldValues', 'SWITCH_CHARS', 'DEFAULT_NESTED_RE',
]


# --------------------------------------------------------------------------
# enums
# --------------------------------------------------------------------------

class _Named(Enum):
    """Enum base whose ``str()`` is the bare member name, for error messages."""

    def __str__(self) -> str:
        return self.name.lower()


class TokenKind(_Named):
    """What a :class:`Token` lexed to."""

    #: A bare or quoted run of text: a positional, or a switch's value.
    WORD = 'word'
    #: A backslash followed by one switch character, e.g. ``\\d``.
    SWITCH = 'switch'
    #: A placeholder standing in for a flattened nested field group.
    NESTED = 'nested'


class ArgKind(_Named):
    """Whether a :attr:`FieldValues.raw` entry came from a switch or a positional."""

    POSITIONAL = 'positional'
    SWITCH = 'switch'


class Action(_Named):
    """What a switch does with the token that follows it."""

    #: Consume the next token as this switch's value; last occurrence wins.
    STORE = 'store'
    #: Consume the next token and append it to a list.
    APPEND = 'append'
    #: Consume nothing; set the destination true.
    STORE_TRUE = 'store_true'
    #: Consume nothing; set the destination false.
    STORE_FALSE = 'store_false'
    #: Consume nothing; tally how many times the switch appeared.
    COUNT = 'count'

    @property
    def arity(self) -> int:
        """Number of following tokens this action consumes: 0 or 1."""
        return 1 if self in (Action.STORE, Action.APPEND) else 0


# --------------------------------------------------------------------------
# errors
# --------------------------------------------------------------------------

class FieldError(Exception):
    """Base class for everything this module raises."""


class FieldDefinitionError(FieldError):
    """The parser was declared incorrectly (a programming error)."""


class FieldSyntaxError(FieldError):
    """The instruction text could not be parsed (a data error)."""

    def __init__(self, message: str, instruction: str = '', pos: int | None = None):
        self.instruction = instruction
        self.pos = pos
        super().__init__(message)

    def __str__(self) -> str:
        base = super().__str__()
        if self.pos is None:
            return base
        return f'{base} (at offset {self.pos}: {self.instruction[self.pos:self.pos + 30]!r})'


# --------------------------------------------------------------------------
# RTF-level unescaping
# --------------------------------------------------------------------------

_RTF_ESCAPE_RE = re.compile(r"\\(?:'([0-9a-fA-F]{2})|u(-?\d+) ?|([\\{}]))")


def unescape_rtf(text: str, encoding: str = 'cp1252', uc: int = 1) -> str:
    r"""Undo RTF-level escaping, leaving raw field instruction text.

    Handles ``\\``, ``\{``, ``\}``, ``\'xx`` hex bytes and ``\uN`` code points.
    Run this on the contents of the ``\fldinst`` group *before* parsing, so
    that ``\\d`` in the file becomes the ``\d`` switch.

    ``uc`` is the number of fallback characters that follow a ``\uN`` and must
    be discarded; pass whatever the enclosing group's ``\ucN`` set it to.
    """
    out, pos = [], 0
    for m in _RTF_ESCAPE_RE.finditer(text):
        if m.start() < pos:          # inside a skipped fallback run
            continue
        out.append(text[pos:m.start()])
        hexbyte, uni, literal = m.groups()
        end = m.end()
        if hexbyte is not None:
            out.append(bytes([int(hexbyte, 16)]).decode(encoding, errors='replace'))
        elif uni is not None:
            n = int(uni)
            out.append(chr(n + 65536 if n < 0 else n))
            end = min(end + uc, len(text))
        else:
            out.append(literal)
        pos = end
    out.append(text[pos:])
    return ''.join(out)


# --------------------------------------------------------------------------
# tokenizer
# --------------------------------------------------------------------------

class NestedField:
    """Placeholder for a nested ``{\\field}`` group flattened out of the text."""

    __slots__ = ('index',)

    def __init__(self, index: int):
        self.index = index

    def __repr__(self) -> str:
        return f'NestedField({self.index})'

    def __eq__(self, other: object) -> bool:
        return isinstance(other, NestedField) and other.index == self.index

    def __hash__(self) -> int:
        return hash(('NestedField', self.index))


class Token:
    """A single lexical unit produced by :func:`tokenize`."""

    __slots__ = ('kind', 'value', 'pos', 'quoted')

    def __init__(self, kind: TokenKind, value: Any, pos: int, quoted: bool = False):
        self.kind = kind
        self.value = value
        self.pos = pos
        self.quoted = quoted

    def __repr__(self) -> str:
        return f'Token({self.kind}, {self.value!r}, pos={self.pos})'


#: Characters Word accepts after the backslash in a field switch.
SWITCH_CHARS = r'A-Za-z0-9*@#!&'

#: Default marker for a nested field, e.g. ``"\x00" "0" "\x00"``.
DEFAULT_NESTED_RE = re.compile('\x00(\\d+)\x00')

_TOKEN_RE = re.compile(
    r'''
      "(?P<quoted>(?:\\.|[^"\\])*)"          # "a quoted string"
    | (?P<switch>\\[''' + SWITCH_CHARS + r'''])(?=\s|$|")
    | (?P<word>[^\s"]+)                      # bare run of non-space
    ''',
    re.VERBOSE,
)

# Inside quotes only \" and \\ are escapes. A backslash before anything else
# is literal, so "C:\Images\logo.png" keeps its separators.
_QUOTED_ESCAPE_RE = re.compile(r'\\(["\\])')


def tokenize(instruction: str, nested_re: re.Pattern | None = DEFAULT_NESTED_RE) -> list[Token]:
    r"""Split instruction text into tokens.

    A backslash only starts a switch when followed by a single switch
    character *and* a word boundary, so UNC paths like ``\\server\share``
    tokenize as one word rather than as a run of bogus switches.
    """
    tokens: list[Token] = []
    pos, n = 0, len(instruction)
    while pos < n:
        if instruction[pos].isspace():
            pos += 1
            continue
        m = _TOKEN_RE.match(instruction, pos)
        if m is None:
            raise FieldSyntaxError('unterminated quoted string', instruction, pos)
        group = m.lastgroup
        raw = m.group(group)
        if group == 'quoted':
            tokens.append(
                Token(TokenKind.WORD, _QUOTED_ESCAPE_RE.sub(r'\1', raw), pos, quoted=True))
        elif group == 'switch':
            tokens.append(Token(TokenKind.SWITCH, raw, pos))
        else:
            hit = nested_re.fullmatch(raw) if nested_re else None
            if hit:
                tokens.append(Token(TokenKind.NESTED, NestedField(int(hit.group(1))), pos))
            else:
                tokens.append(Token(TokenKind.WORD, raw, pos))
        pos = m.end()
    return tokens


# --------------------------------------------------------------------------
# declarations
# --------------------------------------------------------------------------

_IDENT_RE = re.compile(r'[A-Za-z_][A-Za-z0-9_]*\Z')

_ACTION_DEFAULTS = {
    Action.STORE: None,
    Action.APPEND: None,
    Action.STORE_TRUE: False,
    Action.STORE_FALSE: True,
    Action.COUNT: 0,
}


class _Switch:
    def __init__(self, letter, action, dest, default, metavar, choices, help):
        self.letter = letter
        self.action = action
        self.dest = dest
        self.default = default
        self.metavar = metavar
        self.choices = choices
        self.help = help

    @property
    def arity(self) -> int:
        return self.action.arity


class _Positional:
    def __init__(self, name, required, default, nargs, help):
        self.name = name
        self.required = required
        self.default = default
        self.nargs = nargs
        self.help = help


class FieldValues:
    """Parse result. Attribute access by ``dest``, plus the ordered raw list."""

    def __init__(self, field: str, values: dict, raw: list):
        self.__dict__.update(values)
        self._field = field
        self._raw = raw

    @property
    def field(self) -> str:
        """The field name as declared (canonical case)."""
        return self._field

    @property
    def raw(self) -> list[tuple[ArgKind, str, Any]]:
        """Ordered ``(ArgKind, name, value)`` triples, preserving source order."""
        return self._raw

    def get(self, name: str, default: Any = None) -> Any:
        return self.__dict__.get(name, default)

    def __getitem__(self, name: str) -> Any:
        try:
            return self.__dict__[name]
        except KeyError:
            raise KeyError(name) from None

    def __contains__(self, name: str) -> bool:
        return name in self.__dict__

    def __iter__(self) -> Iterator[tuple[str, Any]]:
        for k, v in self.__dict__.items():
            if not k.startswith('_'):
                yield k, v

    def __repr__(self) -> str:
        body = ', '.join(f'{k}={v!r}' for k, v in self)
        return f'<{self._field} {body}>'


# --------------------------------------------------------------------------
# the parser
# --------------------------------------------------------------------------

class FieldParser:
    r"""Declarative parser for one field type.

    :param name: canonical field name, e.g. ``'INCLUDEPICTURE'``.
    :param aliases: other spellings accepted in place of ``name``.
    :param general_switches: register the universal ``\*`` formatting switch
        (``dest='format'``, repeatable). On by default since it may appear on
        essentially any field.
    :param case_sensitive: if false (the default, matching Word) ``\D`` and
        ``\d`` are the same switch.
    :param allow_unknown_switches: collect unrecognised switches into
        ``values.unknown`` instead of raising.
    """

    def __init__(self, name: str, *, aliases: tuple[str, ...] = (),
                 general_switches: bool = True, case_sensitive: bool = False,
                 allow_unknown_switches: bool = False, help: str = ''):
        self.name = name
        self.aliases = aliases
        self.case_sensitive = case_sensitive
        self.allow_unknown_switches = allow_unknown_switches
        self.help = help
        self._switches: dict[str, _Switch] = {}
        self._positionals: list[_Positional] = []
        if general_switches:
            self.add_switch('*', Action.APPEND, dest='format', metavar='FORMAT',
                            help='general formatting switch')

    # -- declaration -------------------------------------------------------

    def _key(self, letter: str) -> str:
        return letter if self.case_sensitive else letter.lower()

    def add_switch(self, letter: str, action: Action | str = Action.STORE, *,
                   dest: str | None = None, default: Any = None,
                   metavar: str | None = None, choices: tuple[str, ...] | None = None,
                   help: str = '') -> None:
        r"""Declare a switch. ``letter`` is given without its backslash.

        The default action is :attr:`Action.STORE`: the switch consumes the
        following token as its value. Use :attr:`Action.APPEND` for switches
        that may repeat and should collect their values into a list.

        For switches that take no argument, pass :attr:`Action.STORE_TRUE`
        (or ``STORE_FALSE``, or ``COUNT`` to tally repeats). This mirrors
        ``argparse``, so a bare ``add_switch('c')`` takes a value and a flag
        must say so explicitly. Actions may be given as :class:`Action`
        members or as their lowercase string names.
        """
        letter = letter.lstrip('\\')
        if len(letter) != 1:
            raise FieldDefinitionError(f'switch must be one character, got {letter!r}')
        try:
            action = Action(action)
        except ValueError:
            raise FieldDefinitionError(
                f'unknown action {action!r}; expected one of '
                f'{", ".join(a.value for a in Action)}') from None
        if dest is None:
            if not _IDENT_RE.match(letter):
                raise FieldDefinitionError(
                    f'switch {letter!r} is not a valid identifier; pass dest= explicitly')
            dest = letter
        if default is None:
            default = _ACTION_DEFAULTS[action]
        key = self._key(letter)
        if key in self._switches:
            raise FieldDefinitionError(f'switch {letter!r} already declared')
        self._switches[key] = _Switch(letter, action, dest, default,
                                      metavar or dest.upper(), choices, help)

    def add_argument(self, name: str, *, required: bool = False, default: Any = None,
                     nargs: str | None = None, help: str = '') -> None:
        """Declare a positional. ``nargs='*'`` makes it a trailing catch-all."""
        if not _IDENT_RE.match(name):
            raise FieldDefinitionError(f'{name!r} is not a valid identifier')
        if nargs not in (None, '*'):
            raise FieldDefinitionError("nargs must be None or '*'")
        if self._positionals and self._positionals[-1].nargs == '*':
            raise FieldDefinitionError("no positional may follow an nargs='*' positional")
        if nargs == '*' and default is None:
            default = []
        self._positionals.append(_Positional(name, required, default, nargs, help))

    # -- parsing -----------------------------------------------------------

    def parse(self, instruction: str, *, expect_name: bool = True,
              nested_re: re.Pattern | None = DEFAULT_NESTED_RE) -> FieldValues:
        """Parse instruction text and return a :class:`FieldValues`.

        Set ``expect_name=False`` if the leading field name has already been
        consumed by whatever chose this parser.
        """
        tokens = tokenize(instruction, nested_re)
        return self.parse_tokens(tokens, instruction=instruction, expect_name=expect_name)

    def parse_tokens(self, tokens: list[Token], *, instruction: str = '',
                     expect_name: bool = True) -> FieldValues:
        i = 0
        if expect_name:
            if not tokens:
                raise FieldSyntaxError('empty field instruction', instruction, 0)
            got = tokens[0]
            if got.kind is not TokenKind.WORD or not self.matches_name(str(got.value)):
                raise FieldSyntaxError(
                    f'expected field name {self.name!r}, got {got.value!r}',
                    instruction, got.pos)
            i = 1

        values: dict[str, Any] = {}
        raw: list[tuple[ArgKind, str, Any]] = []
        seen: set[str] = set()
        unknown: list[tuple[str, Any]] = []
        pending = list(self._positionals)

        for sw in self._switches.values():
            if sw.action is Action.APPEND and sw.default is None:
                values[sw.dest] = []
            else:
                values[sw.dest] = sw.default
        for p in self._positionals:
            values[p.name] = list(p.default) if p.nargs == '*' else p.default

        while i < len(tokens):
            tok = tokens[i]

            if tok.kind is TokenKind.SWITCH:
                letter = tok.value[1:]
                sw = self._switches.get(self._key(letter))
                if sw is None:
                    if not self.allow_unknown_switches:
                        raise FieldSyntaxError(
                            f'{self.name}: unknown switch {tok.value!r}', instruction, tok.pos)
                    nxt = tokens[i + 1] if i + 1 < len(tokens) else None
                    if nxt is not None and nxt.kind is not TokenKind.SWITCH:
                        unknown.append((letter, nxt.value))
                        i += 2
                    else:
                        unknown.append((letter, None))
                        i += 1
                    continue

                if sw.arity == 0:
                    if sw.action is Action.COUNT:
                        values[sw.dest] = (values[sw.dest] or 0) + 1
                    else:
                        values[sw.dest] = (sw.action is Action.STORE_TRUE)
                    raw.append((ArgKind.SWITCH, sw.dest, None))
                    seen.add(sw.dest)
                    i += 1
                    continue

                if i + 1 >= len(tokens) or tokens[i + 1].kind is TokenKind.SWITCH:
                    raise FieldSyntaxError(
                        f'{self.name}: switch {tok.value!r} requires an argument',
                        instruction, tok.pos)
                arg = tokens[i + 1].value
                if sw.choices is not None and arg not in sw.choices:
                    raise FieldSyntaxError(
                        f'{self.name}: {tok.value} expects one of '
                        f'{", ".join(sw.choices)}, got {arg!r}',
                        instruction, tokens[i + 1].pos)
                if sw.action is Action.APPEND:
                    if values.get(sw.dest) is None:
                        values[sw.dest] = []
                    values[sw.dest].append(arg)
                else:
                    values[sw.dest] = arg
                raw.append((ArgKind.SWITCH, sw.dest, arg))
                seen.add(sw.dest)
                i += 2
                continue

            # positional
            if not pending:
                raise FieldSyntaxError(
                    f'{self.name}: unexpected argument {tok.value!r}', instruction, tok.pos)
            target = pending[0]
            if target.nargs == '*':
                values[target.name].append(tok.value)
            else:
                values[target.name] = tok.value
                pending.pop(0)
            raw.append((ArgKind.POSITIONAL, target.name, tok.value))
            seen.add(target.name)
            i += 1

        for p in self._positionals:
            if p.required and p.name not in seen:
                raise FieldSyntaxError(
                    f'{self.name}: missing required argument {p.name!r}', instruction, None)

        if self.allow_unknown_switches:
            values['unknown'] = unknown
        return FieldValues(self.name, values, raw)

    def matches_name(self, token: str) -> bool:
        """True if ``token`` is this field's name or one of its aliases."""
        candidates = (self.name,) + self.aliases
        if self.case_sensitive:
            return token in candidates
        return token.upper() in {c.upper() for c in candidates}

    # -- serialisation -----------------------------------------------------

    def unparse(self, values: FieldValues, *, quote_all: bool = False) -> str:
        """Rebuild instruction text from a parse result, preserving order."""
        out = [self.name]
        by_dest = {sw.dest: sw for sw in self._switches.values()}
        for kind, name, val in values.raw:
            if kind is ArgKind.POSITIONAL:
                out.append(_quote(val, quote_all))
            else:
                sw = by_dest[name]
                out.append('\\' + sw.letter)
                if sw.arity:
                    out.append(_quote(val, quote_all))
        return ' '.join(out)

    # -- introspection -----------------------------------------------------

    def format_help(self) -> str:
        lines = [f'{self.name} field' + (f' - {self.help}' if self.help else '')]
        if self._positionals:
            lines.append('  arguments:')
            for p in self._positionals:
                tag = ' (required)' if p.required else ''
                star = '...' if p.nargs == '*' else ''
                lines.append(f'    {p.name}{star:<12}{p.help}{tag}')
        if self._switches:
            lines.append('  switches:')
            for sw in self._switches.values():
                spec = f'\\{sw.letter}' + (f' {sw.metavar}' if sw.arity else '')
                lines.append(f'    {spec:<18}{sw.help}')
        return '\n'.join(lines)

    def __repr__(self) -> str:
        return f'<FieldParser {self.name} switches={len(self._switches)}>'


_NEEDS_QUOTE_RE = re.compile(r'[\s"]|\A\\')


def _quote(value: Any, always: bool = False) -> str:
    if isinstance(value, NestedField):
        return f'\x00{value.index}\x00'
    s = str(value)
    if always or s == '' or _NEEDS_QUOTE_RE.search(s):
        return '"' + s.replace('\\', '\\\\').replace('"', '\\"') + '"'
    return s
