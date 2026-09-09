"""Regression tests for the crash bugs fixed so far.

Run with: python3 -m unittest test_rtfparser

Every case here is a document that used to abort the parse. The parser takes a path rather than a
stream, so each test writes its source to a temp file first -- letting Parser.parse() accept an
open binary stream would make this a lot less ceremonious.
"""

import os
import tempfile
import unittest

import rtfcharset

from rtfparser import Handler, Parser

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


def parse(source: bytes) -> Recorder:
	fd, path = tempfile.mkstemp(suffix='.rtf')
	try:
		with os.fdopen(fd, 'wb') as f:
			f.write(source)
		doc = Parser(Recorder)
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

	def test_instruction_without_arguments(self):
		# `name, args = ...split(maxsplit=1)` used to raise ValueError before we ever looked the name up
		with self.assertRaises(ValueError) as caught:
			parse(HEADER + rb"{\field{\*\fldinst PAGE}{\fldrslt 1}}}")
		self.assertIn('PAGE', str(caught.exception))

	def test_unknown_instruction_reports_the_name(self):
		# this path used to raise NameError, because InstructionError isn't defined in rtfparser
		with self.assertRaises(ValueError) as caught:
			parse(HEADER + rb"{\field{\*\fldinst BOGUS x}{\fldrslt 1}}}")
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


class TestSampleDocuments(unittest.TestCase):
	# the files in testdocs/ parse end to end without raising

	def documents(self):
		here = os.path.dirname(os.path.abspath(__file__))
		testdocs = os.path.join(here, 'testdocs')
		for name in sorted(os.listdir(testdocs)):
			if name.endswith('.rtf'):
				yield os.path.join(testdocs, name)
		yield os.path.join(here, 'imagetest.rtf')

	def test_all_parse(self):
		for path in self.documents():
			with self.subTest(document=os.path.basename(path)):
				Parser(Recorder).parse(path)

	def test_hyperlinks_are_extracted(self):
		here = os.path.dirname(os.path.abspath(__file__))
		doc = Parser(Recorder)
		doc.parse(os.path.join(here, 'testdocs', 'hyper.rtf'))
		self.assertEqual(doc.output.links, [('Flaming Text', 'http://flamingtext.com')])


if __name__ == '__main__':
	unittest.main()
