import re
import unittest

from fieldparse import (
    FieldParser, FieldRegistry, FieldSyntaxError, FieldDefinitionError,
    UnknownFieldError, NestedField, STANDARD_FIELDS, tokenize, unescape_rtf,
)


def ip():
    p = FieldParser('INCLUDEPICTURE')
    p.add_argument('path', required=True)
    p.add_switch('d', 'store_true')
    p.add_switch('c', 'store')
    return p


class TestTokenizer(unittest.TestCase):
    def test_quoted_path_keeps_backslashes(self):
        toks = tokenize(r'INCLUDEPICTURE "C:\Images\logo.png" \d')
        self.assertEqual([t.value for t in toks],
                         ['INCLUDEPICTURE', r'C:\Images\logo.png', '\\d'])

    def test_unc_path_is_one_word(self):
        toks = tokenize(r'INCLUDEPICTURE \\server\share\logo.png \d')
        self.assertEqual(toks[1].kind, 'word')
        self.assertEqual(toks[1].value, r'\\server\share\logo.png')
        self.assertEqual(toks[2].kind, 'switch')

    def test_escapes_inside_quotes(self):
        toks = tokenize(r'HYPERLINK "a \"b\" c"')
        self.assertEqual(toks[1].value, 'a "b" c')

    def test_unterminated_quote_raises(self):
        with self.assertRaises(FieldSyntaxError):
            tokenize('HYPERLINK "unclosed')

    def test_nested_placeholder(self):
        toks = tokenize('INCLUDEPICTURE \x000\x00 \\d')
        self.assertEqual(toks[1].kind, 'nested')
        self.assertEqual(toks[1].value, NestedField(0))

    def test_switch_before_quote(self):
        toks = tokenize(r'TOC \o"1-3"')
        self.assertEqual([t.kind for t in toks], ['word', 'switch', 'word'])
        self.assertEqual(toks[2].value, '1-3')


class TestUnescape(unittest.TestCase):
    def test_backslash_and_braces(self):
        self.assertEqual(unescape_rtf(r'\\d \{x\}'), r'\d {x}')

    def test_hex_and_unicode(self):
        self.assertEqual(unescape_rtf(r"caf\'e9"), 'caf\u00e9')
        self.assertEqual(unescape_rtf(r'\u8364 ?'), '\u20ac')
        self.assertEqual(unescape_rtf(r'a\u8364 ?b'), 'a\u20acb')
        self.assertEqual(unescape_rtf(r'\u8364 ??', uc=2), '\u20ac')


class TestParsing(unittest.TestCase):
    def test_basic(self):
        v = ip().parse(r'INCLUDEPICTURE "C:\Images\logo.png" \d \* MERGEFORMAT')
        self.assertEqual(v.path, r'C:\Images\logo.png')
        self.assertTrue(v.d)
        self.assertIsNone(v.c)
        self.assertEqual(v.format, ['MERGEFORMAT'])
        self.assertEqual(v.field, 'INCLUDEPICTURE')

    def test_switch_case_insensitive(self):
        self.assertTrue(ip().parse(r'INCLUDEPICTURE x \D').d)

    def test_field_name_case_insensitive(self):
        self.assertEqual(ip().parse(r'includepicture x').path, 'x')

    def test_value_switch(self):
        v = ip().parse(r'INCLUDEPICTURE x \c "PNG Filter"')
        self.assertEqual(v.c, 'PNG Filter')

    def test_missing_switch_arg(self):
        with self.assertRaises(FieldSyntaxError) as cm:
            ip().parse(r'INCLUDEPICTURE x \c \d')
        self.assertIn('requires an argument', str(cm.exception))

    def test_unknown_switch(self):
        with self.assertRaises(FieldSyntaxError) as cm:
            ip().parse(r'INCLUDEPICTURE x \q')
        self.assertIn('unknown switch', str(cm.exception))

    def test_allow_unknown(self):
        p = FieldParser('INCLUDEPICTURE', allow_unknown_switches=True)
        p.add_argument('path')
        v = p.parse(r'INCLUDEPICTURE x \q 7 \z')
        self.assertEqual(v.unknown, [('q', '7'), ('z', None)])

    def test_missing_required_positional(self):
        with self.assertRaises(FieldSyntaxError) as cm:
            ip().parse('INCLUDEPICTURE \\d')
        self.assertIn('missing required', str(cm.exception))

    def test_extra_positional(self):
        with self.assertRaises(FieldSyntaxError):
            ip().parse('INCLUDEPICTURE a b')

    def test_wrong_field_name(self):
        with self.assertRaises(FieldSyntaxError):
            ip().parse('HYPERLINK x')

    def test_nested_value_survives(self):
        v = ip().parse('INCLUDEPICTURE \x003\x00 \\d')
        self.assertEqual(v.path, NestedField(3))

    def test_append_repeats(self):
        v = STANDARD_FIELDS.parse(r'TOC \t "Head1,1" \t "Head2,2" \h')
        self.assertEqual(v.styles, ['Head1,1', 'Head2,2'])
        self.assertTrue(v.hyperlinks)

    def test_count_action(self):
        p = FieldParser('X', general_switches=False)
        p.add_switch('v', 'count')
        self.assertEqual(p.parse(r'X \v \v \v').v, 3)

    def test_choices(self):
        p = FieldParser('X', general_switches=False)
        p.add_switch('m', 'store', choices=('a', 'b'))
        self.assertEqual(p.parse(r'X \m a').m, 'a')
        with self.assertRaises(FieldSyntaxError):
            p.parse(r'X \m c')

    def test_nargs_star(self):
        v = STANDARD_FIELDS.parse(r'= SUM(A1:A5) * 2 \# "#,##0.00"')
        self.assertEqual(v.expression, ['SUM(A1:A5)', '*', '2'])
        self.assertEqual(v.numeric_picture, '#,##0.00')

    def test_raw_order_preserved(self):
        v = ip().parse(r'INCLUDEPICTURE \d x \c f')
        self.assertEqual([r[1] for r in v.raw], ['d', 'path', 'c'])

    def test_mapping_access(self):
        v = ip().parse('INCLUDEPICTURE x')
        self.assertEqual(v['path'], 'x')
        self.assertIn('d', v)
        self.assertEqual(v.get('nope', 'fallback'), 'fallback')


class TestSharedLetterArity(unittest.TestCase):
    """The whole point: \\c means different things in different fields."""

    def test_c_takes_value_in_includepicture(self):
        self.assertEqual(STANDARD_FIELDS.parse(r'INCLUDEPICTURE x \c GIF').c, 'GIF')

    def test_c_takes_nothing_in_seq(self):
        v = STANDARD_FIELDS.parse(r'SEQ figure \c')
        self.assertTrue(v.repeat_nearest)
        self.assertEqual(v.identifier, 'figure')

    def test_d_differs_between_fields(self):
        self.assertTrue(STANDARD_FIELDS.parse(r'INCLUDEPICTURE x \d').d)
        self.assertEqual(STANDARD_FIELDS.parse(r'REF bm \d ", "').separator, ', ')

    def test_h_differs(self):
        self.assertTrue(STANDARD_FIELDS.parse(r'TOC \h').hyperlinks)
        self.assertTrue(STANDARD_FIELDS.parse(r'REF bm \h').hyperlink)

    def test_l_differs(self):
        self.assertEqual(STANDARD_FIELDS.parse(r'HYPERLINK "u" \l "bm"').anchor, 'bm')
        self.assertEqual(STANDARD_FIELDS.parse(r'TOC \l "1-3"').tc_levels, '1-3')


class TestRegistry(unittest.TestCase):
    def test_dispatch(self):
        v = STANDARD_FIELDS.parse(r'HYPERLINK "https://x.test" \o "tip" \n')
        self.assertEqual(v.target, 'https://x.test')
        self.assertEqual(v.tooltip, 'tip')
        self.assertTrue(v.new_window)

    def test_unknown_field(self):
        with self.assertRaises(UnknownFieldError):
            STANDARD_FIELDS.parse('NOSUCHFIELD x')

    def test_alias(self):
        self.assertEqual(STANDARD_FIELDS.parse('= 1+1').expression, ['1+1'])

    def test_default_parser_fallback(self):
        fallback = FieldParser('UNKNOWN', allow_unknown_switches=True)
        fallback.add_argument('args', nargs='*')
        reg = FieldRegistry(default=fallback)
        v = reg.parse(r'WEIRDFIELD a b \q')
        self.assertEqual(v.args, ['a', 'b'])
        self.assertEqual(v.unknown, [('q', None)])

    def test_contains(self):
        self.assertIn('toc', STANDARD_FIELDS)
        self.assertNotIn('nope', STANDARD_FIELDS)


class TestUnparse(unittest.TestCase):
    def test_roundtrip(self):
        p = ip()
        src = r'INCLUDEPICTURE "C:\Images\logo.png" \d \* MERGEFORMAT'
        out = p.unparse(p.parse(src))
        self.assertEqual(p.parse(out).path, r'C:\Images\logo.png')
        self.assertIn(r'\d', out)
        self.assertIn(r'\* MERGEFORMAT', out)

    def test_quotes_when_needed(self):
        p = ip()
        out = p.unparse(p.parse('INCLUDEPICTURE "a b.png"'))
        self.assertEqual(out, 'INCLUDEPICTURE "a b.png"')

    def test_unc_gets_quoted(self):
        p = ip()
        out = p.unparse(p.parse(r'INCLUDEPICTURE \\srv\s\l.png'))
        self.assertEqual(p.parse(out).path, r'\\srv\s\l.png')


class TestDefinitionErrors(unittest.TestCase):
    def test_bad_action(self):
        with self.assertRaises(FieldDefinitionError):
            FieldParser('X').add_switch('a', 'frobnicate')

    def test_duplicate_switch(self):
        p = FieldParser('X', general_switches=False)
        p.add_switch('a')
        with self.assertRaises(FieldDefinitionError):
            p.add_switch('A')

    def test_non_identifier_needs_dest(self):
        p = FieldParser('X', general_switches=False)
        with self.assertRaises(FieldDefinitionError):
            p.add_switch('#')
        p.add_switch('#', 'store', dest='picture')

    def test_positional_after_star(self):
        p = FieldParser('X')
        p.add_argument('a', nargs='*')
        with self.assertRaises(FieldDefinitionError):
            p.add_argument('b')


class TestHelp(unittest.TestCase):
    def test_format_help(self):
        text = STANDARD_FIELDS.get('INCLUDEPICTURE').format_help()
        self.assertIn('INCLUDEPICTURE', text)
        self.assertIn('\\d', text)
        self.assertIn('\\c CONVERTER', text)


if __name__ == '__main__':
    unittest.main(verbosity=2)
