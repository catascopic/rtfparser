"""Regression tests for the crash bugs fixed so far.

Run from the repository root with: python3 -m unittest discover

Every case here is a document that used to abort the parse. The parser takes a path rather than a
stream, so each test writes its source to a temp file first -- letting Parser.parse() accept an
open binary stream would make this a lot less ceremonious.
"""

import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import rtfcharset

from rtfparser import Destination, Handler, NullDevice, NULL_DEVICE, Parser, RtfWarning

# a minimal but complete preamble, so each case only has to supply the part under test
HEADER = rb"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil Arial;}}"


class Recorder(Handler):
	# records everything the parser hands us, so a test can assert on the document's content

	def __init__(self, doc):
		super().__init__(doc)
		self.text = []
		self.links = []
		self.pictures = []
		self.lists = []
		self.warnings = []

	def write(self, text):
		self.text.append(text)

	def par(self):
		pass

	def page_break(self):
		pass

	def hyperlink(self, text, args):
		self.links.append((text, args.url))

	def picture(self, pic):
		self.pictures.append((pic.format, bytes(pic.data)))

	def numbering_on(self, info):
		self.lists.append(info)

	def warning(self, message, position=None):
		self.warnings.append(message)


def parse(source: bytes, strict: bool = False) -> Recorder:
	fd, path = tempfile.mkstemp(suffix='.rtf')
	try:
		with os.fdopen(fd, 'wb') as f:
			f.write(source)
		doc = Parser(Recorder, strict=strict)
		doc.parse(path)
		return doc.output
	finally:
		os.remove(path)


def text_of(source: bytes) -> str:
	return ''.join(parse(source).text)


class TestUnicodeAndBinary(unittest.TestCase):
	# \u and \bin called module-level functions that don't exist

	def test_unicode_escape(self):
		self.assertEqual(text_of(HEADER + rb"a\u945 ?b}"), 'aαb')

	def test_surrogate_pair(self):
		self.assertEqual(text_of(HEADER + rb"a\u-10179 ?\u-8477 ?b}"), 'a\U0001f6e3b')

	def test_bin_keeps_bytes_undecoded(self):
		out = parse(HEADER + b"{\\pict\\pngblip\\picw1\\pich1\\bin4 \x00\x01\xfe\xff}}")
		self.assertEqual(out.pictures, [('png', b'\x00\x01\xfe\xff')])


class TestFields(unittest.TestCase):

	def test_hyperlink(self):
		# the field's switches used to be splatted in as keyword arguments the handler never declared
		out = parse(HEADER + b'{\\field{\\*\\fldinst HYPERLINK "http://x.com"}{\\fldrslt Click}}}')
		self.assertEqual(out.links, [('Click', 'http://x.com')])

	def test_hyperlink_with_switch(self):
		out = parse(HEADER + b'{\\field{\\*\\fldinst HYPERLINK "http://x.com" \\\\o "a tip"}{\\fldrslt C}}}')
		self.assertEqual(out.links, [('C', 'http://x.com')])

	def test_unmodelled_field_passes_its_result_through(self):
		# PAGE/DATE/NUMPAGES are everywhere in real documents and have no parser of their own.
		# The result is what the writer last rendered for the field, so it's what we show.
		for name in (b'PAGE', b'DATE', b'NUMPAGES'):
			with self.subTest(field=name.decode()):
				out = parse(HEADER + b"{\\field{\\*\\fldinst " + name + b"}{\\fldrslt 7}}}")
				self.assertEqual(''.join(out.text), '7')
				self.assertEqual(len(out.warnings), 1)
				self.assertIn(name.decode(), out.warnings[0])

	def test_unknown_instruction_reports_the_name(self):
		# this path used to raise NameError, because InstructionError isn't defined in rtfparser
		out = parse(HEADER + rb"{\field{\*\fldinst BOGUS x}{\fldrslt 1}}}")
		self.assertIn('BOGUS', out.warnings[0])

	def test_unknown_instruction_raises_when_strict(self):
		with self.assertRaises(RtfWarning) as caught:
			parse(HEADER + rb"{\field{\*\fldinst BOGUS x}{\fldrslt 1}}}", strict=True)
		self.assertIn('BOGUS', str(caught.exception))


class TestCharset(unittest.TestCase):

	def test_ansi_without_codepage(self):
		# CHARSETS mapped \ansi to 'ansi', which is not a codec Python knows
		self.assertEqual(text_of(rb"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil A;}}\f0 caf\'e9}"), 'café')

	def test_no_charset_declared_at_all(self):
		self.assertEqual(text_of(rb"{\rtf1\deff0{\fonttbl{\f0\fnil A;}}\f0 caf\'e9}"), 'café')

	def test_unknown_fcharset_falls_back(self):
		source = rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset199 A;}}\f0 caf\'e9}"
		self.assertEqual(text_of(source), 'café')

	def test_get_encoding_precedence(self):
		self.assertEqual(rtfcharset.get_encoding(204, 'cp850'), 'cp1251')  # the font wins
		self.assertEqual(rtfcharset.get_encoding(1, 'cp850'), 'cp850')     # \fcharset1 defers
		self.assertEqual(rtfcharset.get_encoding(None, 'cp850'), 'cp850')
		self.assertEqual(rtfcharset.get_encoding(999, None), rtfcharset.DEFAULT_ENCODING)


class TestFontFallback(unittest.TestCase):

	def test_no_font_table(self):
		self.assertEqual(text_of(rb"{\rtf1\ansi\ansicpg1252 caf\'e9}"), 'café')

	def test_undefined_font_index(self):
		self.assertEqual(text_of(HEADER + rb"\f9 caf\'e9}"), 'café')

	def test_plain_without_deff(self):
		self.assertEqual(text_of(rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil A;}}a\plain b}"), 'ab')


class TestNumberingOutsideGroup(unittest.TestCase):
	# every \pn* word assumed a {\*\pn} group had already created a Numbering to write to

	def test_pnseclvl_at_top_level(self):
		self.assertEqual(text_of(HEADER + rb"\pnseclvl1\pndec x}"), 'x')

	def test_pn_properties_at_top_level(self):
		self.assertEqual(text_of(HEADER + rb"\pnstart5\pnlvl2\pnf1 x}"), 'x')

	def test_pntxtb_without_numbering(self):
		self.assertEqual(text_of(HEADER + rb"{\pntxtb -}x}"), 'x')

	def test_pnseclvl_destination_group(self):
		self.assertEqual(text_of(HEADER + rb"{\*\pnseclvl3\pndec{\pntxta .}}x}"), 'x')


class TestStarDestination(unittest.TestCase):
	# try_read_dest didn't consume the control word's parameter or its trailing delimiter

	def test_space_after_destination_word(self):
		out = parse(HEADER + rb"{\*\pn \pnlvlblt\pndec{\pntxtb -}}\pard x\par}")
		self.assertEqual(''.join(out.text), 'x')
		self.assertEqual([n.style for n in out.lists], ['pndec'])

	def test_no_space_after_destination_word(self):
		out = parse(HEADER + rb"{\*\pn\pnlvlblt\pndec{\pntxtb -}}\pard x\par}")
		self.assertEqual(''.join(out.text), 'x')
		self.assertEqual([n.before for n in out.lists], ['-'])

	def test_unknown_destination_is_skipped_whole(self):
		self.assertEqual(text_of(HEADER + rb"a{\*\bkmkstart foo}b}"), 'ab')

	def test_group_immediately_after_destination_word(self):
		out = parse(HEADER + rb"{\*\shppict{\pict\pngblip\picw1\pich1 00ff}}x}")
		self.assertEqual(out.pictures, [('png', b'\x00\xff')])
		self.assertEqual(''.join(out.text), 'x')

	def test_nonshppict_swallows_the_picture_it_wraps(self):
		out = parse(HEADER + rb"{\nonshppict{\pict\pngblip 00ff}}x}")
		self.assertEqual(out.pictures, [])
		self.assertEqual(''.join(out.text), 'x')


class TestNullDevice(unittest.TestCase):

	def test_overrides_the_whole_destination_protocol(self):
		# __getattr__ can't cover these: they resolve on Destination, whose defaults raise, so it
		# never fires for them. Any new protocol method has to be answered here explicitly.
		for name, member in vars(Destination).items():
			if callable(member) and not name.startswith('_'):
				with self.subTest(method=name):
					self.assertIn(name, vars(NullDevice), f"NullDevice must override {name}()")

	def test_paragraph_inside_a_skipped_destination(self):
		# \par used to reach Destination.par and raise "NullDevice can't handle paragraphs"
		out = parse(HEADER + rb"a{\*\bkmkstart foo\par}b}")
		self.assertEqual(''.join(out.text), 'ab')

	def test_page_break_inside_a_skipped_destination(self):
		out = parse(HEADER + rb"a{\*\bkmkstart foo\page}b}")
		self.assertEqual(''.join(out.text), 'ab')

	def test_attributes_do_not_stick_to_the_singleton(self):
		# a \pict inside a group we're skipping used to write format onto the shared singleton,
		# where it outlived the document that set it
		parse(HEADER + rb"{\nonshppict{\pict\pngblip 00ff}}x}")
		self.assertEqual(vars(NULL_DEVICE), {})

	def test_sub_destinations_still_resolve_to_the_null_device(self):
		self.assertIs(NULL_DEVICE.instruction, NULL_DEVICE)
		self.assertIs(NULL_DEVICE.result, NULL_DEVICE)

	def test_field_nested_in_a_skipped_destination(self):
		# the document is well formed; the field is only unreachable because we chose to skip
		# its container, so it has to be swallowed rather than raise
		field = b'{\\field{\\*\\fldinst HYPERLINK "http://x.com"}{\\fldrslt Click}}'
		out = parse(HEADER + b'a{\\*\\annotation ' + field + b'}b}')
		self.assertEqual(''.join(out.text), 'ab')
		self.assertEqual(out.links, [])


class TestWarnings(unittest.TestCase):

	def test_stray_pn_word_warns(self):
		out = parse(HEADER + rb"\pnstart5 x}")
		self.assertEqual(''.join(out.text), 'x')
		self.assertEqual(len(out.warnings), 1)
		self.assertIn('start', out.warnings[0])

	def test_stray_pn_word_raises_when_strict(self):
		with self.assertRaises(RtfWarning):
			parse(HEADER + rb"\pnstart5 x}", strict=True)

	def test_pn_word_inside_a_skipped_destination_is_silent(self):
		# {\*\pnseclvl3 ...} is valid RTF we simply don't model, so it isn't worth reporting
		out = parse(HEADER + rb"{\*\pnseclvl3\pndec{\pntxta .}}x}")
		self.assertEqual(''.join(out.text), 'x')
		self.assertEqual(out.warnings, [])

	def test_unknown_font_charset_warns_once_per_font(self):
		source = (rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset199 Weird;}}"
		          rb"\f0 caf\'e9 caf\'e9 caf\'e9}")
		out = parse(source)
		self.assertEqual(''.join(out.text), 'café café café')
		self.assertEqual(len(out.warnings), 1)
		self.assertIn('199', out.warnings[0])

	def test_unexpected_text_in_color_table_warns(self):
		out = parse(rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil A;}}"
		            rb"{\colortbl;\red255\green0\blue0 oops;}\f0 x}")
		self.assertTrue(any('color table' in w for w in out.warnings))

	def test_warning_carries_a_byte_offset(self):
		positions = []

		class Positional(Recorder):
			def warning(self, message, position=None):
				positions.append(position)

		fd, path = tempfile.mkstemp(suffix='.rtf')
		try:
			with os.fdopen(fd, 'wb') as f:
				f.write(HEADER + rb"\pnstart5 x}")
			Parser(Positional).parse(path)
		finally:
			os.remove(path)
		self.assertEqual(len(positions), 1)
		self.assertIsInstance(positions[0], int)
		self.assertGreater(positions[0], 0)

	def test_strict_error_reports_the_position(self):
		with self.assertRaises(RtfWarning) as caught:
			parse(HEADER + rb"\pnstart5 x}", strict=True)
		self.assertIsNotNone(caught.exception.position)
		self.assertIn(str(caught.exception.position), str(caught.exception))

	def test_warnings_are_silent_by_default(self):
		# the base Output.warning is a no-op, so a plain handler never has to care
		class Quiet(Handler):
			def write(self, text):
				pass

			def par(self):
				pass

		fd, path = tempfile.mkstemp(suffix='.rtf')
		try:
			with os.fdopen(fd, 'wb') as f:
				f.write(HEADER + rb"\pnstart5 x}")
			Parser(Quiet).parse(path)
		finally:
			os.remove(path)


class TestSampleDocuments(unittest.TestCase):
	# the files in testdocs/ parse end to end without raising

	def documents(self):
		testdocs = os.path.join(ROOT, 'testdocs')
		for name in sorted(os.listdir(testdocs)):
			if name.endswith('.rtf'):
				yield os.path.join(testdocs, name)
		yield os.path.join(ROOT, 'imagetest.rtf')

	def test_all_parse(self):
		for path in self.documents():
			with self.subTest(document=os.path.basename(path)):
				Parser(Recorder).parse(path)

	def test_hyperlinks_are_extracted(self):
		doc = Parser(Recorder)
		doc.parse(os.path.join(ROOT, 'testdocs', 'hyper.rtf'))
		self.assertEqual(doc.output.links, [('Flaming Text', 'http://flamingtext.com')])


if __name__ == '__main__':
	unittest.main()
