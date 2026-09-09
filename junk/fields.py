r"""
fields - ready-made :class:`~fieldparse.FieldParser` constants.

One module-level constant per field instruction, each already declared with
its positionals and switches::

    >>> from fields import INCLUDEPICTURE
    >>> v = INCLUDEPICTURE.parse(r'INCLUDEPICTURE "logo.png" \d')
    >>> v.path, v.d
    ('logo.png', True)

Pick the constant for the instruction you have. The same switch letter means
different things in different fields, so each parser carries its own arity::

    >>> INCLUDEPICTURE.parse(r'INCLUDEPICTURE x \c GIF').c
    'GIF'
    >>> SEQ.parse(r'SEQ figure \c').repeat_nearest
    True

Every parser accepts the universal ``\*`` formatting switch, collected into
``values.format``. Adjust or extend a constant in place if your documents use
switches these definitions do not model, or build your own with
:class:`~fieldparse.FieldParser`.
"""

from __future__ import annotations

from fieldparse import Action, FieldParser

__all__ = [
    'INCLUDEPICTURE', 'INCLUDETEXT', 'HYPERLINK', 'MERGEFIELD', 'TOC',
    'REF', 'PAGEREF', 'NOTEREF', 'SEQ', 'STYLEREF', 'SYMBOL', 'XE',
    'LISTNUM', 'ASK', 'FILLIN', 'FORMULA',
    'DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE', 'EDITTIME',
]


# --------------------------------------------------------------------------
# INCLUDEPICTURE
# --------------------------------------------------------------------------

def _includepicture() -> FieldParser:
    p = FieldParser('INCLUDEPICTURE', help='insert an image by path or URL')
    p.add_argument('path', required=True, help='file path or URL')
    p.add_switch('d', Action.STORE_TRUE, help='do not store image data in the document')
    p.add_switch('c', metavar='CONVERTER', help='graphics filter to use')
    return p


#: ``INCLUDEPICTURE "path" [\d] [\c converter]``
INCLUDEPICTURE = _includepicture()


# --------------------------------------------------------------------------
# INCLUDETEXT
# --------------------------------------------------------------------------

def _includetext() -> FieldParser:
    p = FieldParser('INCLUDETEXT', help='insert text from another document')
    p.add_argument('path', required=True, help='file path')
    p.add_argument('bookmark', help='range within the source document')
    p.add_switch('c', metavar='CONVERTER', help='file converter to use')
    p.add_switch('!', Action.STORE_TRUE, dest='no_update_nested',
                 help='do not update fields in the inserted text')
    return p


#: ``INCLUDETEXT "path" [bookmark] [\c converter] [\!]``
INCLUDETEXT = _includetext()


# --------------------------------------------------------------------------
# HYPERLINK
# --------------------------------------------------------------------------

def _hyperlink() -> FieldParser:
    p = FieldParser('HYPERLINK', help='link to a document, anchor or URL')
    p.add_argument('target', help='URL or file path')
    p.add_switch('l', dest='anchor', metavar='BOOKMARK', help='anchor within the target')
    p.add_switch('o', dest='tooltip', metavar='TEXT', help='screen tip')
    p.add_switch('t', dest='frame', metavar='FRAME', help='target frame')
    p.add_switch('n', Action.STORE_TRUE, dest='new_window', help='open in a new window')
    p.add_switch('m', Action.STORE_TRUE, dest='image_map', help='coordinates from an image map')
    return p


#: ``HYPERLINK "target" [\l anchor] [\o tip] [\t frame] [\n] [\m]``
HYPERLINK = _hyperlink()


# --------------------------------------------------------------------------
# MERGEFIELD
# --------------------------------------------------------------------------

def _mergefield() -> FieldParser:
    p = FieldParser('MERGEFIELD', help='mail-merge data field')
    p.add_argument('name', required=True, help='data source column')
    p.add_switch('b', dest='before', metavar='TEXT', help='text to insert before non-empty data')
    p.add_switch('f', dest='after', metavar='TEXT', help='text to insert after non-empty data')
    p.add_switch('m', Action.STORE_TRUE, dest='mapped', help='this is a mapped field')
    p.add_switch('v', Action.STORE_TRUE, dest='vertical', help='enable vertical formatting')
    return p


#: ``MERGEFIELD name [\b before] [\f after] [\m] [\v]``
MERGEFIELD = _mergefield()


# --------------------------------------------------------------------------
# TOC
# --------------------------------------------------------------------------

def _toc() -> FieldParser:
    p = FieldParser('TOC', help='table of contents')
    p.add_switch('o', dest='levels', metavar='"1-3"', help='build from heading levels')
    p.add_switch('t', Action.APPEND, dest='styles', metavar='"STYLE,LEVEL"',
                 help='build from custom styles; repeatable')
    p.add_switch('h', Action.STORE_TRUE, dest='hyperlinks', help='entries are hyperlinks')
    p.add_switch('n', dest='no_page_numbers', metavar='"1-3"',
                 help='omit page numbers for these levels')
    p.add_switch('p', dest='separator', metavar='SEP', help='entry/page-number separator')
    p.add_switch('w', Action.STORE_TRUE, dest='preserve_tabs', help='keep tab entries')
    p.add_switch('x', Action.STORE_TRUE, dest='preserve_newlines', help='keep newlines')
    p.add_switch('z', Action.STORE_TRUE, dest='hide_leaders', help='hide leaders in web layout')
    p.add_switch('b', dest='bookmark', metavar='NAME', help='restrict to a bookmarked range')
    p.add_switch('f', dest='identifier', metavar='ID', help='build from TC entries of this type')
    p.add_switch('l', dest='tc_levels', metavar='"1-3"', help='TC entry levels to include')
    p.add_switch('u', Action.STORE_TRUE, dest='use_outline', help='use applied outline levels')
    return p


#: ``TOC [\o "1-3"] [\t "Style,Level"] [\h] ...``
TOC = _toc()


# --------------------------------------------------------------------------
# REF / PAGEREF / NOTEREF
# --------------------------------------------------------------------------

def _ref(name: str, help: str) -> FieldParser:
    p = FieldParser(name, help=help)
    p.add_argument('bookmark', required=True, help='bookmark name')
    p.add_switch('h', Action.STORE_TRUE, dest='hyperlink', help='insert as a hyperlink')
    p.add_switch('p', Action.STORE_TRUE, dest='relative_position',
                 help='report position relative to the source')
    p.add_switch('n', Action.STORE_TRUE, dest='paragraph_number',
                 help='insert the paragraph number')
    p.add_switch('r', Action.STORE_TRUE, dest='relative_number',
                 help='insert the relative paragraph number')
    p.add_switch('w', Action.STORE_TRUE, dest='full_number',
                 help='insert the full context paragraph number')
    p.add_switch('d', dest='separator', metavar='SEP', help='separator for sequence numbers')
    p.add_switch('f', Action.STORE_TRUE, dest='note_number',
                 help='insert the footnote or endnote number')
    p.add_switch('t', Action.STORE_TRUE, dest='suppress_non_delimiter',
                 help='suppress non-delimiter text')
    return p


#: ``REF bookmark [\h] [\p] [\d sep] ...``
REF = _ref('REF', 'cross-reference to a bookmark')

#: ``PAGEREF bookmark [\h] [\p] ...``
PAGEREF = _ref('PAGEREF', 'page number of a bookmark')

#: ``NOTEREF bookmark [\h] [\p] [\f]``
NOTEREF = _ref('NOTEREF', 'footnote or endnote reference')


# --------------------------------------------------------------------------
# SEQ
# --------------------------------------------------------------------------

def _seq() -> FieldParser:
    p = FieldParser('SEQ', help='sequential numbering (figures, tables)')
    p.add_argument('identifier', required=True, help='sequence name')
    p.add_argument('bookmark', help='repeat the number at this bookmark')
    p.add_switch('c', Action.STORE_TRUE, dest='repeat_nearest',
                 help='repeat the closest preceding number')
    p.add_switch('h', Action.STORE_TRUE, dest='hidden', help='hide the result')
    p.add_switch('n', Action.STORE_TRUE, dest='next', help='insert the next number (default)')
    p.add_switch('r', dest='reset_to', metavar='N', help='reset the sequence to N')
    p.add_switch('s', dest='reset_level', metavar='LEVEL', help='reset at this heading level')
    return p


#: ``SEQ identifier [bookmark] [\c] [\h] [\n] [\r N] [\s level]``
SEQ = _seq()


# --------------------------------------------------------------------------
# STYLEREF
# --------------------------------------------------------------------------

def _styleref() -> FieldParser:
    p = FieldParser('STYLEREF', help='text of the nearest paragraph in a style')
    p.add_argument('style', required=True, help='style name or number')
    p.add_switch('l', Action.STORE_TRUE, dest='search_up',
                 help='search from the bottom of the page upward')
    p.add_switch('n', Action.STORE_TRUE, dest='paragraph_number',
                 help='insert the paragraph number')
    p.add_switch('p', Action.STORE_TRUE, dest='relative_position',
                 help='report position relative to the source')
    p.add_switch('r', Action.STORE_TRUE, dest='relative_number',
                 help='insert the relative paragraph number')
    p.add_switch('w', Action.STORE_TRUE, dest='full_number',
                 help='insert the full context paragraph number')
    p.add_switch('t', Action.STORE_TRUE, dest='suppress_non_delimiter',
                 help='suppress non-delimiter text')
    return p


#: ``STYLEREF style [\l] [\n] [\p] ...``
STYLEREF = _styleref()


# --------------------------------------------------------------------------
# SYMBOL
# --------------------------------------------------------------------------

def _symbol() -> FieldParser:
    p = FieldParser('SYMBOL', help='insert a character by code')
    p.add_argument('code', required=True, help='character code')
    p.add_switch('f', dest='font', metavar='NAME', help='font to take the glyph from')
    p.add_switch('s', dest='size', metavar='POINTS', help='point size')
    p.add_switch('a', Action.STORE_TRUE, dest='ansi', help='code is ANSI')
    p.add_switch('u', Action.STORE_TRUE, dest='unicode', help='code is Unicode')
    p.add_switch('h', Action.STORE_TRUE, dest='no_line_height',
                 help='do not affect line height')
    p.add_switch('j', Action.STORE_TRUE, dest='shift_jis', help='code is Shift-JIS')
    return p


#: ``SYMBOL code [\f font] [\s size] [\a] [\u] [\h] [\j]``
SYMBOL = _symbol()


# --------------------------------------------------------------------------
# XE
# --------------------------------------------------------------------------

def _xe() -> FieldParser:
    p = FieldParser('XE', help='index entry')
    p.add_argument('entry', required=True, help='entry text')
    p.add_switch('b', Action.STORE_TRUE, dest='bold', help='bold the page number')
    p.add_switch('i', Action.STORE_TRUE, dest='italic', help='italicise the page number')
    p.add_switch('f', dest='entry_type', metavar='TYPE', help='index type')
    p.add_switch('r', dest='bookmark', metavar='NAME', help='page range from a bookmark')
    p.add_switch('t', dest='cross_ref', metavar='TEXT', help='cross-reference text')
    p.add_switch('y', dest='yomi', metavar='TEXT', help='yomi for sorting')
    return p


#: ``XE "entry" [\b] [\i] [\f type] [\r bookmark] [\t text] [\y yomi]``
XE = _xe()


# --------------------------------------------------------------------------
# LISTNUM
# --------------------------------------------------------------------------

def _listnum() -> FieldParser:
    p = FieldParser('LISTNUM', help='inline list number')
    p.add_argument('name', help='list name')
    p.add_switch('l', dest='level', metavar='N', help='level within the list')
    p.add_switch('s', dest='start', metavar='N', help='starting value')
    return p


#: ``LISTNUM [name] [\l level] [\s start]``
LISTNUM = _listnum()


# --------------------------------------------------------------------------
# ASK / FILLIN
# --------------------------------------------------------------------------

def _ask() -> FieldParser:
    p = FieldParser('ASK', help='prompt the user and store the answer in a bookmark')
    p.add_argument('bookmark', required=True, help='bookmark to assign')
    p.add_argument('prompt', required=True, help='prompt text')
    p.add_switch('d', dest='default', metavar='TEXT', help='default response')
    p.add_switch('o', Action.STORE_TRUE, dest='once_per_merge', help='prompt once per merge')
    return p


def _fillin() -> FieldParser:
    p = FieldParser('FILLIN', help='prompt the user and insert the answer')
    p.add_argument('prompt', required=True, help='prompt text')
    p.add_switch('d', dest='default', metavar='TEXT', help='default response')
    p.add_switch('o', Action.STORE_TRUE, dest='once_per_merge', help='prompt once per merge')
    return p


#: ``ASK bookmark "prompt" [\d default] [\o]``
ASK = _ask()

#: ``FILLIN "prompt" [\d default] [\o]``
FILLIN = _fillin()


# --------------------------------------------------------------------------
# = (formula)
# --------------------------------------------------------------------------

def _formula() -> FieldParser:
    p = FieldParser('=', aliases=('FORMULA',), help='evaluate an expression')
    p.add_argument('expression', nargs='*', help='the expression to evaluate')
    p.add_switch('#', dest='numeric_picture', metavar='"#,##0.00"',
                 help='numeric formatting picture')
    return p


#: ``= expression [\# picture]`` — the field name is a bare ``=``.
FORMULA = _formula()


# --------------------------------------------------------------------------
# date and time
# --------------------------------------------------------------------------

def _datetime(name: str, help: str) -> FieldParser:
    p = FieldParser(name, help=help)
    p.add_switch('@', dest='picture', metavar='"dd MMMM yyyy"', help='date-time picture')
    p.add_switch('h', Action.STORE_TRUE, dest='hijri', help='use the Hijri calendar')
    p.add_switch('s', Action.STORE_TRUE, dest='saka', help='use the Saka Era calendar')
    return p


#: ``DATE [\@ "picture"] [\h] [\s]``
DATE = _datetime('DATE', 'current date')

#: ``TIME [\@ "picture"]``
TIME = _datetime('TIME', 'current time')

#: ``CREATEDATE [\@ "picture"]``
CREATEDATE = _datetime('CREATEDATE', 'document creation date')

#: ``SAVEDATE [\@ "picture"]``
SAVEDATE = _datetime('SAVEDATE', 'date last saved')

#: ``PRINTDATE [\@ "picture"]``
PRINTDATE = _datetime('PRINTDATE', 'date last printed')

#: ``EDITTIME [\@ "picture"]``
EDITTIME = _datetime('EDITTIME', 'total editing time')
