import unittest

import fields
from fieldparse import FieldParser, FieldSyntaxError
from fields import (
    ASK, DATE, FILLIN, FORMULA, HYPERLINK, INCLUDEPICTURE, INCLUDETEXT,
    LISTNUM, MERGEFIELD, PAGEREF, REF, SEQ, STYLEREF, SYMBOL, TOC, XE,
)


class TestSharedLetterArity(unittest.TestCase):
    r"""The point of per-field parsers: \c and friends differ by field."""

    def test_c_takes_a_value_in_includepicture(self):
        self.assertEqual(INCLUDEPICTURE.parse(r'INCLUDEPICTURE x \c GIF').c, 'GIF')

    def test_c_takes_nothing_in_seq(self):
        v = SEQ.parse(r'SEQ figure \c')
        self.assertTrue(v.repeat_nearest)
        self.assertEqual(v.identifier, 'figure')

    def test_d_differs_between_fields(self):
        self.assertTrue(INCLUDEPICTURE.parse(r'INCLUDEPICTURE x \d').d)
        self.assertEqual(REF.parse(r'REF bm \d ", "').separator, ', ')
        self.assertEqual(ASK.parse(r'ASK bm "Name?" \d "none"').default, 'none')

    def test_h_differs_between_fields(self):
        self.assertTrue(TOC.parse(r'TOC \h').hyperlinks)
        self.assertTrue(REF.parse(r'REF bm \h').hyperlink)
        self.assertTrue(SEQ.parse(r'SEQ fig \h').hidden)
        self.assertTrue(DATE.parse(r'DATE \h').hijri)

    def test_l_differs_between_fields(self):
        self.assertEqual(HYPERLINK.parse(r'HYPERLINK "u" \l "bm"').anchor, 'bm')
        self.assertEqual(TOC.parse(r'TOC \l "1-3"').tc_levels, '1-3')
        self.assertEqual(LISTNUM.parse(r'LISTNUM \l 2').level, '2')
        self.assertTrue(STYLEREF.parse(r'STYLEREF Heading1 \l').search_up)


class TestIndividualFields(unittest.TestCase):
    def test_includepicture(self):
        v = INCLUDEPICTURE.parse(r'INCLUDEPICTURE "C:\Images\logo.png" \d \* MERGEFORMAT')
        self.assertEqual(v.path, r'C:\Images\logo.png')
        self.assertTrue(v.d)
        self.assertEqual(v.format, ['MERGEFORMAT'])

    def test_includepicture_requires_path(self):
        with self.assertRaises(FieldSyntaxError):
            INCLUDEPICTURE.parse(r'INCLUDEPICTURE \d')

    def test_hyperlink(self):
        v = HYPERLINK.parse(r'HYPERLINK "https://x.test/a b" \l "sec2" \o "Go" \n')
        self.assertEqual(v.target, 'https://x.test/a b')
        self.assertEqual(v.anchor, 'sec2')
        self.assertEqual(v.tooltip, 'Go')
        self.assertTrue(v.new_window)
        self.assertFalse(v.image_map)

    def test_toc_repeats_styles(self):
        v = TOC.parse(r'TOC \o "1-3" \t "Caption,4" \t "Quote,5" \h \z \u')
        self.assertEqual(v.levels, '1-3')
        self.assertEqual(v.styles, ['Caption,4', 'Quote,5'])
        self.assertTrue(v.hyperlinks and v.hide_leaders and v.use_outline)

    def test_mergefield(self):
        v = MERGEFIELD.parse(r'MERGEFIELD ImagePath \b "[" \f "]"')
        self.assertEqual((v.name, v.before, v.after), ('ImagePath', '[', ']'))

    def test_pageref_shares_ref_shape(self):
        self.assertTrue(PAGEREF.parse(r'PAGEREF _Ref1 \h').hyperlink)
        self.assertEqual(PAGEREF.name, 'PAGEREF')

    def test_seq_two_positionals(self):
        v = SEQ.parse(r'SEQ Figure fig12 \* ARABIC')
        self.assertEqual((v.identifier, v.bookmark), ('Figure', 'fig12'))
        self.assertEqual(v.format, ['ARABIC'])

    def test_includetext_bang_switch(self):
        v = INCLUDETEXT.parse(r'INCLUDETEXT "boiler.docx" Intro \!')
        self.assertEqual((v.path, v.bookmark), ('boiler.docx', 'Intro'))
        self.assertTrue(v.no_update_nested)

    def test_symbol(self):
        v = SYMBOL.parse(r'SYMBOL 61548 \f "Wingdings" \s 10 \h')
        self.assertEqual((v.code, v.font, v.size), ('61548', 'Wingdings', '10'))
        self.assertTrue(v.no_line_height)

    def test_xe(self):
        v = XE.parse(r'XE "RTF:fields" \b \t "See fields"')
        self.assertEqual(v.entry, 'RTF:fields')
        self.assertTrue(v.bold)
        self.assertEqual(v.cross_ref, 'See fields')

    def test_date_picture(self):
        self.assertEqual(DATE.parse(r'DATE \@ "dd MMMM yyyy"').picture, 'dd MMMM yyyy')

    def test_fillin(self):
        v = FILLIN.parse(r'FILLIN "Your name?" \d "Anon" \o')
        self.assertEqual((v.prompt, v.default), ('Your name?', 'Anon'))
        self.assertTrue(v.once_per_merge)

    def test_formula_bare_equals(self):
        v = FORMULA.parse(r'= SUM(A1:A5) * 2 \# "#,##0.00"')
        self.assertEqual(v.expression, ['SUM(A1:A5)', '*', '2'])
        self.assertEqual(v.numeric_picture, '#,##0.00')

    def test_formula_alias(self):
        self.assertEqual(FORMULA.parse('FORMULA 1+1').expression, ['1+1'])


class TestConstantsAreWellFormed(unittest.TestCase):
    def test_all_exports_exist_and_are_parsers(self):
        for name in fields.__all__:
            with self.subTest(field=name):
                obj = getattr(fields, name)
                self.assertIsInstance(obj, FieldParser)

    def test_every_parser_accepts_general_format_switch(self):
        for name in fields.__all__:
            parser = getattr(fields, name)
            with self.subTest(field=name):
                self.assertIn('*', parser._switches)

    def test_constant_name_matches_its_field(self):
        for name in fields.__all__:
            parser = getattr(fields, name)
            with self.subTest(field=name):
                self.assertTrue(parser.matches_name(parser.name))

    def test_constants_are_distinct_objects(self):
        self.assertIsNot(REF, PAGEREF)
        self.assertIsNot(ASK, FILLIN)

    def test_wrong_parser_rejects_the_instruction(self):
        with self.assertRaises(FieldSyntaxError):
            HYPERLINK.parse(r'INCLUDEPICTURE x \d')


if __name__ == '__main__':
    unittest.main(verbosity=2)
