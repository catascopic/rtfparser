"""Tests for the document tree prototype.

Run from the repository root with: python3 test/test_rtfdoc.py
"""

import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import rtfdoc

from rtfdoc import Image, LineBreak, Link, Run

HEADER = rb"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil Arial;}}"


def read(source: bytes):
	return rtfdoc.read_bytes(source)


def document(name: str):
	return rtfdoc.read(os.path.join(ROOT, name))


class TestRuns(unittest.TestCase):

	def test_runs_sharing_a_style_are_merged(self):
		# the parser splits at every control word, so this arrives as three separate writes
		doc = document('imagetest.rtf')
		self.assertEqual([r.text for r in doc.paragraphs[0].runs], ['Image Test'])

	def test_runs_with_different_styles_stay_separate(self):
		doc = document('testdocs/test.rtf')
		runs = doc.paragraphs[-1].runs
		self.assertEqual([r.text for r in runs], ['Heli', 'co', 'pter'])
		self.assertEqual(
			[(r.style.italic, bool(r.style.underline)) for r in runs],
			[(False, True), (True, True), (True, False)])

	def test_style_records_points_not_half_points(self):
		out = read(HEADER + rb"\f0\fs24 twelve point}")
		self.assertEqual(out.paragraphs[0].runs[0].style.size, 12.0)

	def test_line_becomes_a_node_not_a_newline(self):
		out = read(HEADER + rb"\f0 first\line second\par}")
		content = out.paragraphs[0].content
		self.assertEqual([type(c) for c in content], [Run, LineBreak, Run])
		self.assertEqual(out.paragraphs[0].text, 'firstsecond')

	def test_paragraph_text_joins_its_runs(self):
		out = read(HEADER + rb"\f0 plain \b bold\b0  plain\par}")
		self.assertEqual(out.paragraphs[0].text, 'plain bold plain')


class TestLinks(unittest.TestCase):

	def test_a_link_split_across_fields_is_condensed(self):
		# Word emits this url as eight separate HYPERLINK fields
		doc = document('imagetest.rtf')
		self.assertEqual(len(doc.links), 1)
		self.assertEqual(doc.links[0].text, 'https://www.flamingtext.com/')
		self.assertEqual(doc.links[0].url, 'https://www.flamingtext.com/')

	def test_links_with_different_urls_stay_separate(self):
		source = (HEADER + b'{\\field{\\*\\fldinst HYPERLINK "http://a.com"}{\\fldrslt A}}'
		          b'{\\field{\\*\\fldinst HYPERLINK "http://b.com"}{\\fldrslt B}}\\par}')
		out = read(source)
		self.assertEqual([(l.text, l.url) for l in out.links],
			[('A', 'http://a.com'), ('B', 'http://b.com')])

	def test_a_link_is_inline_beside_other_runs(self):
		doc = document('testdocs/cont.rtf')
		para = next(p for p in doc.paragraphs if p.content and isinstance(p.content[0], Link))
		self.assertEqual([type(c) for c in para.content], [Link, Run])
		self.assertEqual(para.text, 'www.bzpower.com testing')


class TestImages(unittest.TestCase):

	def test_an_image_is_inline_content(self):
		doc = document('imagetest.rtf')
		self.assertEqual(len(doc.images), 1)
		image = doc.images[0]
		self.assertEqual(image.format, 'png')
		self.assertEqual(image.data[:4], b'\x89PNG')
		self.assertEqual((image.width, image.height), (1746, 1323))

	def test_an_image_contributes_no_text(self):
		doc = document('imagetest.rtf')
		para = next(p for p in doc.paragraphs if any(isinstance(c, Image) for c in p.content))
		self.assertEqual(para.text, '')


class TestLists(unittest.TestCase):

	def test_each_item_carries_its_own_marker(self):
		doc = document('testdocs/test.rtf')
		items = [p for p in doc.paragraphs if p.numbering]
		self.assertEqual([p.numbering.marker for p in items], ['I.', 'II.', 'III.'])
		self.assertEqual([p.text for p in items], ['One', 'Two', 'Three!'])

	def test_the_marker_stays_out_of_the_body_text(self):
		# \pntext is decoration for readers that don't number; putting it in the text is what
		# produced "I.\t                    - One" in the markdown output
		doc = document('testdocs/test.rtf')
		self.assertNotIn('I.', doc.paragraphs[1].text)

	def test_list_style_is_recorded(self):
		doc = document('testdocs/test.rtf')
		self.assertEqual({p.numbering.style for p in doc.paragraphs if p.numbering}, {'pnucrm'})

	def test_paragraphs_outside_the_list_have_no_numbering(self):
		doc = document('testdocs/test.rtf')
		self.assertIsNone(doc.paragraphs[0].numbering)
		self.assertIsNone(doc.paragraphs[-1].numbering)


class TestParagraphs(unittest.TestCase):

	def test_a_trailing_par_does_not_leave_an_empty_paragraph(self):
		# \par terminates a paragraph where <p> contains one
		out = read(HEADER + rb"\f0 only\par}")
		self.assertEqual([p.text for p in out.paragraphs], ['only'])

	def test_an_intentional_blank_paragraph_is_kept(self):
		out = read(HEADER + rb"\f0 one\par\par two\par}")
		self.assertEqual([p.text for p in out.paragraphs], ['one', '', 'two'])

	def test_alignment_is_recorded(self):
		out = read(HEADER + rb"\f0\qc centred\par\pard\f0 left\par}")
		self.assertEqual([p.alignment for p in out.paragraphs], ['c', 'l'])

	def test_a_page_break_marks_the_following_paragraph(self):
		out = read(HEADER + rb"\f0 before\par\page after\par}")
		self.assertEqual([p.page_break_before for p in out.paragraphs], [False, True])


class TestDocument(unittest.TestCase):

	def test_metadata_is_carried_through(self):
		out = read(HEADER + rb"{\info{\title A Title}{\author Max}}\f0 body\par}")
		self.assertEqual(out.info.title, 'A Title')
		self.assertEqual(out.info.author, 'Max')

	def test_tables_are_captured_after_the_parse(self):
		# fonts and colours fill up as the document is read, so a builder that snapshots them
		# in its constructor gets empty ones
		out = read(HEADER + rb"{\colortbl;\red255\green0\blue0;}\f0 body\par}")
		self.assertEqual(out.fonts[0].name, 'Arial')
		self.assertTrue(out.colors)

	def test_document_text_joins_paragraphs(self):
		out = read(HEADER + rb"\f0 one\par two\par}")
		self.assertEqual(out.text, 'one\ntwo')

	def test_warnings_are_collected(self):
		out = read(HEADER + rb"\f0\pnstart5 body\par}")
		self.assertTrue(out.warnings)
		self.assertIn('numbering', out.warnings[0])

	def test_page_furniture_stays_out_of_the_body(self):
		out = read(HEADER + rb"{\header a page header}{\footer a page footer}\f0 body\par}")
		self.assertEqual(out.text, 'body')

	def test_every_sample_document_builds(self):
		names = ['imagetest.rtf'] + [
			os.path.join('testdocs', n)
			for n in sorted(os.listdir(os.path.join(ROOT, 'testdocs'))) if n.endswith('.rtf')]
		for name in names:
			with self.subTest(document=name):
				self.assertTrue(document(name).paragraphs)


if __name__ == '__main__':
	unittest.main()
