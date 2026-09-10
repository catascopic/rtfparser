"""
A document tree built on top of the streaming parser.

This is a prototype. It's an ordinary Output -- nothing in rtfparser.py knows it exists -- which is
the point: the event API stays the engine, and this is the friendly layer over it. Everything the
handlers in test.py and convertmd.py currently do by hand (coalescing runs that share a style,
condensing hyperlinks Word split across several fields, keeping a marker out of the body text)
happens here once, so a consumer walks a tree instead of tracking state.

	doc = read('report.rtf')
	for para in doc.paragraphs:
		print(para.text)

Two deliberate omissions, both explained at their call sites below: formatting inside a hyperlink is
flattened (the parser hands us the result as one string), and list nesting is not inferred.
"""

from __future__ import annotations

import os

import rtfparser

from dataclasses import dataclass, field
from typing import BinaryIO, Union

from rtfparser import Color, Font, Handler, Info, Numbering, Picture

# an rtf property dict holds ints, strings and bools; we keep the raw one on every style
Properties = dict[str, Union[str, int, bool]]


@dataclass(frozen=True)
class Style:
	# the character formatting in effect for a run. Frozen so runs can be compared and merged,
	# which is the whole reason this exists rather than handing out the parser's live prop dict.
	bold: bool = False
	italic: bool = False
	underline: Union[bool, str] = False  # \ul is True, \uldb and friends give their variant
	strike: bool = False
	small_caps: bool = False
	caps: bool = False
	hidden: bool = False  # \v -- usually shouldn't be rendered at all
	superscript: bool = False
	subscript: bool = False
	font: Font | None = None
	size: float | None = None  # points; rtf counts in half-points
	color: Color | None = None

	@classmethod
	def current(cls, handler: Handler) -> Style:
		prop = handler.prop
		size = handler.font_size
		return cls(
			bold=bool(handler.bold),
			italic=bool(handler.italic),
			underline=handler.underline,
			strike=bool(prop.get('strike')),
			small_caps=bool(prop.get('scaps')),
			caps=bool(prop.get('caps')),
			hidden=bool(prop.get('v')),
			superscript=bool(prop.get('super')),
			subscript=bool(prop.get('sub')),
			font=handler.font,
			size=None if size is None else size / 2,
			color=handler.color_foreground)


@dataclass
class Run:
	text: str
	style: Style

	def __repr__(self):
		return f"Run({self.text!r})"


@dataclass
class LineBreak:
	# \line: a break within a paragraph, not between paragraphs. A node rather than a newline
	# inside a run's text, so nobody downstream has to go looking for one.
	text = ''


@dataclass
class Image:
	# an image embedded in the document. Sizes are in the source's own units, except the goal
	# sizes, which are twips (1/1440 inch) -- what the document wants it displayed at.
	format: str | None
	data: bytes
	width: int | None = None
	height: int | None = None
	goal_width: int | None = None
	goal_height: int | None = None
	scale_x: int = 100
	scale_y: int = 100
	text = ''

	@classmethod
	def of(cls, pic: Picture) -> Image:
		return cls(pic.format, bytes(pic.data), pic.width, pic.height,
			pic.goal_width, pic.goal_height, pic.scale_x, pic.scale_y)

	def __repr__(self):
		return f"Image({self.format}, {len(self.data)} bytes, {self.width}x{self.height})"


@dataclass
class Link:
	url: str
	content: list[Run]

	@property
	def text(self):
		return ''.join(run.text for run in self.content)

	def __repr__(self):
		return f"Link({self.text!r} -> {self.url})"


Inline = Union[Run, LineBreak, Image, Link]


@dataclass
class ListInfo:
	# what a paragraph knows about its own place in a list. Deliberately flat: rtf doesn't nest
	# lists, it tags each paragraph with a level, so any hierarchy is inferred rather than read.
	# Inferring it is an interpretation that can be wrong, so it belongs in a pass of its own
	# rather than baked in here where a caller couldn't get back to the truth.
	style: str | None = None  # pndec, pnucrm, ...
	level: int = 0
	marker: str = ''  # the literal text rtf supplies for readers that don't do numbering
	start: int = 1
	indent: int = 0


@dataclass
class Paragraph:
	content: list[Inline] = field(default_factory=list)
	alignment: str = 'l'
	indent: int = 0  # \li, in twips
	first_line_indent: int = 0  # \fi
	numbering: ListInfo | None = None
	page_break_before: bool = False
	style_index: int | None = None  # \s -- an index into a stylesheet we don't read yet

	@property
	def text(self) -> str:
		# just the words: images and breaks contribute nothing
		return ''.join(item.text for item in self.content)

	@property
	def runs(self):
		return [item for item in self.content if isinstance(item, Run)]

	def __repr__(self):
		return f"Paragraph({self.text!r})"


@dataclass
class Document:
	info: Info
	paragraphs: list[Paragraph] = field(default_factory=list)
	fonts: dict[int, Font] = field(default_factory=dict)
	colors: list[Color] = field(default_factory=list)
	warnings: list[str] = field(default_factory=list)

	@property
	def text(self) -> str:
		return '\n'.join(p.text for p in self.paragraphs)

	@property
	def images(self):
		return [item for p in self.paragraphs for item in p.content if isinstance(item, Image)]

	@property
	def links(self):
		return [item for p in self.paragraphs for item in p.content if isinstance(item, Link)]


class DocumentBuilder(Handler):

	def __init__(self):
		self.document = Document(Info())
		self.content: list[Inline] = []
		self.marker: list[str] = []
		self.current_list: Numbering | None = None
		self.pending_break = False

	# --- inline content ------------------------------------------------

	def add(self, item: Inline):
		# runs that share a style are one run. The parser splits text at every control word, so a
		# document says 'Image', ' ', 'Test' where it means 'Image Test' -- left alone, every
		# consumer re-derives this, and the ones that don't emit <u>Heli</u><u>co</u>.
		if isinstance(item, Run) and self.content:
			last = self.content[-1]
			if isinstance(last, Run) and last.style == item.style:
				last.text += item.text
				return
		self.content.append(item)

	def write(self, text: str):
		# \line arrives inside a run as a newline; it's a break, so it becomes one
		style = Style.current(self)
		for i, chunk in enumerate(text.split('\n')):
			if i:
				self.content.append(LineBreak())
			if chunk:
				self.add(Run(chunk, style))

	def plain_text(self, text: str):
		# \pntext: the literal marker rtf provides so readers that ignore numbering still show
		# something. It's decoration, not content, so it goes to the list info and not the body.
		self.marker.append(text)

	def picture(self, pic: Picture):
		self.content.append(Image.of(pic))

	def hyperlink(self, text: str, args):
		# Word splits one visible link across several fields -- eight of them for a single url in
		# imagetest.rtf -- so a link that continues the previous one just extends it.
		# The parser flattens \fldrslt to a string, so a link's own formatting is lost here; giving
		# Field.result a destination that keeps runs would fix it, and nothing else has to change.
		run = Run(text, Style.current(self))
		if self.content:
			last = self.content[-1]
			if isinstance(last, Link) and last.url == args.url:
				last.content.append(run)
				return
		self.content.append(Link(args.url, [run]))

	# --- structure -----------------------------------------------------

	def numbering_on(self, info: Numbering):
		self.current_list = info

	def numbering_off(self, info: Numbering):
		self.current_list = None

	def page_break(self):
		# a break belongs to the paragraph that follows it, so it only ends the current one
		# if that paragraph has anything in it
		if self.content:
			self.flush()
		self.pending_break = True

	def par(self):
		self.flush()

	def list_info(self) -> ListInfo | None:
		marker = ''.join(self.marker).strip()
		if self.current_list is None:
			return ListInfo(marker=marker) if marker else None
		return ListInfo(
			style=self.current_list.style,
			level=self.current_list.level,
			marker=marker or self.current_list.before,
			start=self.current_list.start,
			indent=self.current_list.indent)

	def flush(self):
		# always emits: \par means "a paragraph ended here", and an empty one between two others
		# is a deliberate blank line. Only a paragraph left pending at the end of the document is
		# an artefact, and end_doc is where that's decided.
		numbering = self.list_info()
		self.document.paragraphs.append(Paragraph(
			content=self.content,
			alignment=self.alignment,
			indent=self.prop.get('li', 0),
			first_line_indent=self.prop.get('fi', 0),
			numbering=numbering,
			page_break_before=self.pending_break,
			style_index=self.prop.get('s')))
		self.content = []
		self.marker = []
		self.pending_break = False

	def warning(self, message: str, position: int | None = None):
		self.document.warnings.append(message if position is None else f"{message} at {position}")

	def end_doc(self):
		# \par terminates a paragraph where <p> contains one, so a document ending in \par has
		# nothing pending here -- flushing anyway would invent an empty paragraph at the end
		if self.content or self.marker or self.pending_break:
			self.flush()
		# the tables fill up as we go, so they're copied at the end rather than aliased
		self.document.fonts = dict(self._doc.fonts)
		self.document.colors = list(self._doc.colors)
		self.document.info = self._doc.info


def read(file: str | bytes | os.PathLike, *, strict: bool = False) -> Document:
	return rtfparser.parse(file, DocumentBuilder(), strict=strict).document


def read_bytes(data: bytes, *, strict: bool = False) -> Document:
	return rtfparser.parse_bytes(data, DocumentBuilder(), strict=strict).document


def read_stream(f: BinaryIO, *, strict: bool = False) -> Document:
	return rtfparser.parse_stream(f, DocumentBuilder(), strict=strict).document
