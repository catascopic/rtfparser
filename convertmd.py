#!/usr/bin/env python3
"""
rtf_to_md.py - Convert an RTF file to Markdown using rtfparser.py.

Usage:
	python rtf_to_md.py input.rtf [-o output.md]

If -o is omitted, the output is written next to the input file with a
.md extension, and also printed to stdout if you pass '-' as the output.

Supported RTF -> Markdown mapping:
	\\b			-> **bold**
	\\i			-> *italic*
	\\strike		-> ~~strikethrough~~
	\\ul			-> <u>underline</u>	(no native markdown equivalent)
	\\super/\\sub	-> <sup>/<sub>
	\\par		  -> paragraph break
	\\page		 -> horizontal rule (---)
	HYPERLINK field -> [text](url)
	\\pn (lists)	-> "- " / "1." list items, using the list's own marker text
					  and nesting level (\\pnlvl) for indentation

Anything not listed above (fonts, colors, tables, exact spacing, etc.) is
not represented in Markdown and is dropped, which is expected for this
kind of conversion.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from rtfparser import Parser, Handler, Numbering

# Characters that have special meaning in Markdown and need escaping when they
# appear in plain document text (so they don't get misinterpreted downstream).
_MD_SPECIAL = re.compile(r'([\\`*_\[\]])')


def escape_markdown(text: str) -> str:
	return _MD_SPECIAL.sub(r'\\\1', text)


class MarkdownOutput(Handler):
	"""
	Output/Recorder implementation for rtfparser.Parser that renders the
	document as Markdown text.
	"""

	def __init__(self, doc):
		super().__init__(doc)
		self.parts: list[str] = []
		self._list_prefix: str | None = None
		self._list_indent: int = 0
		self._at_line_start = True

	# --- helpers -----------------------------------------------------

	def _emit(self, text: str):
		if not text:
			return
		if self._list_prefix is not None:
			self.parts.append('  ' * self._list_indent + self._list_prefix)
			self._list_prefix = None
		self.parts.append(text)
		self._at_line_start = False

	# --- Destination interface (called by the parser) ----------------

	def write(self, text: str):
		if not text:
			return

		# RTF's \line becomes a literal '\n' inside a text run (a soft
		# line break within a paragraph), so translate that into a
		# markdown hard line break rather than escaping it as text.
		if '\n' in text:
			for i, chunk in enumerate(text.split('\n')):
				if i > 0:
					self.parts.append('  \n')
					self._at_line_start = True
				self._write_run(chunk)
			return

		self._write_run(text)

	def _write_run(self, text: str):
		if not text:
			return

		text = escape_markdown(text)
		if self.bold:
			text = f"**{text}**"
		if self.italic:
			text = f"*{text}*"
		if self.prop.get('strike'):
			text = f"~~{text}~~"
		if self.underline:
			text = f"<u>{text}</u>"
		if self.prop.get('super'):
			text = f"<sup>{text}</sup>"
		elif self.prop.get('sub'):
			text = f"<sub>{text}</sub>"

		self._emit(text)

	def par(self):
		self.parts.append('\n\n')
		self._at_line_start = True

	def page_break(self):
		self.parts.append('\n\n---\n\n')
		self._at_line_start = True

	def close(self):
		pass

	# --- Output interface ---------------------------------------------

	def plain_text(self, text: str):
		# Used for destinations like \pntext that carry plain, unformatted text.
		self._emit(escape_markdown(text))

	def hyperlink(self, text: str, url: str):
		self._emit(f'[{escape_markdown(text)}]({url})')

	def numbering_on(self, info: Numbering):
		marker = (info.before or '').strip()
		if not marker:
			# No literal marker text was captured; fall back to a generic bullet.
			marker = '-'
		elif marker[-1].isdigit() is False and marker[-1] not in '.):':
			# e.g. a raw bullet glyph -- turn it into a plain markdown bullet
			marker = '-'
		self._list_prefix = marker + ' '
		self._list_indent = max(0, info.level)

	def numbering_off(self, info: Numbering):
		self._list_prefix = None

	def end_doc(self):
		pass

	# --- result --------------------------------------------------------

	@property
	def markdown(self) -> str:
		text = ''.join(self.parts)
		# Collapse runs of 3+ blank lines down to a single blank line.
		text = re.sub(r'\n{3,}', '\n\n', text)
		return text.strip() + '\n'


def convert(rtf_path: str | Path) -> str:
	parser = Parser(MarkdownOutput)
	parser.parse(rtf_path)
	return parser.output.markdown


def main():
	ap = argparse.ArgumentParser(description='Convert an RTF file to Markdown.')
	ap.add_argument('input', help='Path to the .rtf file to convert')
	ap.add_argument('-o', '--output', help="Output path (default: input path with .md extension). Use '-' to print to stdout instead of writing a file.")
	args = ap.parse_args()

	input_path = Path(args.input)
	if not input_path.exists():
		print(f'error: no such file: {input_path}', file=sys.stderr)
		sys.exit(1)

	markdown = convert(input_path)

	if args.output == '-':
		sys.stdout.write(markdown)
		return

	output_path = Path(args.output) if args.output else input_path.with_suffix('.md')
	output_path.write_text(markdown, encoding='utf-8')
	print(f'Wrote {output_path}')


if __name__ == '__main__':
	main()