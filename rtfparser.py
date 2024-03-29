from __future__ import annotations

import os
import struct
from collections import deque

import rtfcharset

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, BinaryIO, Iterable, Callable

# This parser makes a questionable, but I believe justified, decision to parse files in binary mode.
# RTF files are pure ascii, so it sort of doesn't matter. Strings are generally easier to work with in python,
# but using binary mode lets us call seek() on the reader, which is useful when we want to "peek" at the next token in
# the stream. (We also get to use bytebuffer a few times and avoid having to do ''.join on lists of chars.)
# The one awkward part is, in python, characters are 1-char strings, while individual bytes are just ints.
# We don't want to use ints, because they make the code hard to read, so we instead use byte sequences of length 1.
# This isn't ideal, but it's really the only downside to this paradigm (other than having to remember to start your
# strings with the b prefix in that section of the code).

# TODO: \upr, \ud (these are only used for not-output destinations)

ASCII = 'ascii'

CHARSETS = {
	'ansi': 'ansi',
	'pc':   'cp437',
	'pca':  'cp850',
	'mac':  'macintosh'
}

FONT_FAMILIES = frozenset({'fnil', 'froman', 'fswiss', 'fmodern', 'fscript', 'fdecor', 'ftech', 'fbidi'})

# Paragraph formatting (reset with \pard)
PARFMT = frozenset({
	's', 'hyphpar', 'intbl', 'keep', 'nowidctlpar', 'widctlpar',
	'keepn', 'level', 'noline', 'outlinelevel', 'pagebb', 'sbys',
	'q',  # use q to keep track of alignment
	'fi', 'li', 'ri',  # indentation
	'sb', 'sa', 'sl', 'slmult',  # spacing
	'subdocument',
	'rtlpar', 'ltrpar',  # direction
})

TOGGLE = frozenset({'b', 'caps', 'deleted', 'i', 'outl', 'scaps', 'shad', 'strike', 'ul', 'v'})  # hyphpar
# Character formatting (reset with \plain)
CHRFMT = frozenset({
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'expnd', 'expndtw',
	'kerning', 'f', 'fs', 'strikedl', 'up', 'super', 'cf', 'cb', 'rtlch',
	'ltrch', 'cs', 'cchs', 'lang'} | TOGGLE)

# TABS = { ... }

INFO_PROPS = frozenset({'version', 'edmins', 'nofpages', 'nofwords', 'word_count', 'nofchars', 'nofcharsws'})
TEXT_INFO = frozenset({'title', 'subject', 'author', 'manager', 'company', 'operator',
                       'category', 'keywords', 'comment', 'doccomm', 'hlinkbase'})
DATE_INFO = frozenset({'creatim', 'revtim', 'printim', 'buptim'})
NUMBERING_STYLES = frozenset({'pncard', 'pndec', 'pnucltr', 'pnucrm', 'pnlcltr', 'pnlcrm', 'pnord', 'pnordt'})

ESCAPE = {
	'line':      '\n',
	'tab':       '\t',
	'emdash':    '—',
	'endash':    '-',
	'lquote':    '‘',
	'rquote':    '’',
	'ldblquote': '“',
	'rdblquote': '”',
	'bullet':    '•',
}

SPECIAL = {
	b'~': '\u00A0',  # nonbreaking space
	b'-': '\u00AD',  # optional hyphen
	b'_': '\u2011',  # nonbreaking hyphen
}

# can't do frozenset(b'\\{}') because iterating bytes gives you ints
META_CHARS = frozenset({b'\\', b'{', b'}'})

IGNORE_WORDS = frozenset({'nouicompat', 'viewkind'})
UNSUPPORTED_DEST = frozenset({'filetbl', 'stylesheet', 'listtables', 'revtbl'})

BytePredicate = Callable[[bytes], bool]


class Destination(ABC):

	def write(self, text: str):
		raise ValueError(f"{type(self)} can't handle text: {text}")

	def par(self):
		raise ValueError(f"{type(self)} can't handle paragraphs")

	def page_break(self):
		raise ValueError(f"{type(self)} can't handle page breaks")

	def close(self):
		pass


class PlainText(Destination):

	def __init__(self, delegate: Output):
		self.delegate = delegate

	def write(self, text: str):
		self.delegate.plain_text(text)


class RootDest(Destination):
	# TODO: in theory this should only allow one write
	def write(self, text):
		if text != '\u0000':
			raise ValueError(f"expected NUL but got {text}")


@dataclass(frozen=True)
class Font:
	name: str
	family: str
	charset: Optional[str] = None


class FontTable(Destination):

	def __init__(self, doc: Parser):
		self.doc = doc
		self.name = []

	def write(self, text):
		self.name.append(text)
		if text.endswith(';'):
			self.doc.fonts[self.doc.prop['f']] = Font(
				''.join(self.name)[:-1],
				self.doc.prop['family'],
				self.doc.prop.get('fcharset'))
			self.name = []


@dataclass(frozen=True)
class Color:
	red: int
	green: int
	blue: int


BLACK = Color(0, 0, 0)


class ColorTable(Destination):

	def __init__(self, doc: Parser):
		self.doc = doc

	def write(self, text):
		# if text != ';': WARN
		self.doc.colors.append(Color(
			self.doc.prop.get('red', 0),
			self.doc.prop.get('green', 0),
			self.doc.prop.get('blue', 0)))


# TODO: StyleTable?


class Numbering(Destination):

	def __init__(self, doc: Parser):
		self.doc = doc
		self.prop = {}
		self.style = None
		self.level = 0
		self.before = ''
		self.after = ''
		self.font_index = None
		self.indent = 0
		self.start = 1

	# TODO: move this somewhere else?
	def font(self, doc: Parser):
		index = self.font_index
		if index is None:
			index = doc.prop.get('f', doc.deff)
		return doc.fonts[index]

	def close(self):
		self.doc.output.numbering_on(self)


class Field(Destination):

	def __init__(self, delegate: Output):
		self.delegate = delegate
		self.instruction = ''
		self.result = ''
		self.set_instruction = TextSetter(self, 'instruction')
		self.set_result = TextSetter(self, 'result')

	def close(self):
		instr, *params = self.instruction.split()
		if instr == 'HYPERLINK':
			url = params[0].removeprefix('"').removesuffix('"')
			self.delegate.hyperlink(self.result, url)
		else:
			raise ValueError(f"unknown instruction: {self.instruction}")


class SetValue(Destination, ABC):

	def __init__(self, obj, prop):
		self.obj = obj
		self.prop = prop

	@abstractmethod
	def get_value(self):
		raise NotImplementedError

	def close(self):
		setattr(self.obj, self.prop, self.get_value())


class TextSetter(SetValue):

	def __init__(self, obj, prop):
		super().__init__(obj, prop)
		self.content = []

	def write(self, text):
		self.content.append(text)

	def get_value(self):
		return ''.join(self.content)


class TimeSetter(SetValue):

	def __init__(self, doc, obj, prop):
		super().__init__(obj, prop)
		self.doc = doc

	def get_value(self):
		return datetime(*(self.doc.prop[k] for k in ('yr', 'mo', 'dy')),
		                *(self.doc.prop.get(k, 0) for k in ('hr', 'min', 'sec')))


class NullDevice(Destination):
	def write(self, text):
		pass  # do nothing


NULL_DEVICE = NullDevice()


@dataclass
class Info:
	title: str = None
	subject: str = None
	author: str = None
	manager: str = None
	company: str = None
	operator: str = None
	category: str = None
	keywords: str = None
	comment: str = None
	doccomm: str = None
	hlinkbase: str = None
	creatim: datetime = None
	revtim: datetime = None
	printim: datetime = None
	buptim: datetime = None


@dataclass
class Group:
	parent: Optional[Group]
	own_dest: Optional[Destination]
	# TODO: do we want to include string values here?
	prop: dict[str, str | int | bool]

	@classmethod
	def root(cls):
		return cls(None, RootDest(), {})

	def open(self):
		# could use ChainMap here, but we'd have to replace pop() with setting to None/0.
		# seems simpler to just copy the map, since it also reduces lookup time
		return Group(self, None, self.prop.copy())

	@property
	def dest(self):
		return self.own_dest or self.parent.dest

	@dest.setter
	def dest(self, value):
		self.own_dest = value

	def close(self):
		if self.own_dest is not None:
			self.own_dest.close()
		return self.parent


def not_control(c: bytes) -> bool:
	return c not in META_CHARS


def is_letter(c: bytes):
	# control words are supposed to be lower case as per the spec, but in practice some have mixed case
	return b'a' <= c <= b'z' or b'A' <= c <= b'Z'


def is_digit(c: bytes):
	return b'0' <= c <= b'9'


def is_endline(c: bytes):
	return c == b'\r' or c == b'\n'


def read_while(f: BinaryIO, matcher: BytePredicate):
	buf = bytearray()
	read_into_while(f, buf, matcher)
	return buf


def read_into_while(f: BinaryIO, buf: bytearray, matcher: BytePredicate):
	while True:
		c = f.read(1)
		if not matcher(c):
			f.seek(-1, 1)
			return
		if not c:
			return
		buf.extend(c)


def read_word(f: BinaryIO):
	return read_while(f, is_letter).decode(ASCII)


def read_number(f, default: Optional[int] = None):
	c = f.read(1)
	if is_digit(c) or c == b'-':
		buf = bytearray(c)
		read_into_while(f, buf, is_digit)
		return int(buf)
	f.seek(-1, 1)
	return default


def end_control(f: BinaryIO):
	c = f.read(1)
	if not c.isspace():
		f.seek(-1, 1)
	# we don't need to handle the CRLF case because we read the CR here and
	# skip the LF later while reading normally


def consume(f: BinaryIO, expected: bytes):
	actual = f.read(len(expected))
	if actual != expected:
		raise ValueError(f"expected {expected}, got {actual} at {f.tell()}")


def skip_chars(f: BinaryIO, n: int):
	for _ in range(n):
		c = f.read(1)
		if c == b'\\':
			if read_word(f):
				read_number(f)  # unnecessary?
				end_control(f)
			elif f.read(1) == b"'":
				f.read(2)
		elif c == b'{' or c == b'}':
			f.seek(-1, 1)
			return


class Parser:

	def __init__(self, output):
		self.output = output(self)
		self.group = Group.root()
		self.rtf_version = 1
		self.charset: Optional[str] = None
		self.deff: Optional[int] = None
		self.fonts: dict[int, Font] = {}
		self.colors: list[Color] = []
		self.info = Info()
		self.numbering: Optional[Numbering] = None

	def parse(self, file: str | bytes | os.PathLike):
		with open(file, 'rb') as f:
			# TODO: try/except with f.tell()?
			while True:
				text = read_while(f, not_control).translate(None, b'\r\n').decode(ASCII)
				if text:
					self.dest.write(text)
				c = f.read(1)
				if c == b'\\':
					self.read_control(f)
				elif c == b'{':
					self.group = self.group.open()
				elif c == b'}':
					self.group = self.group.close()
				elif c == b'':
					self.output.end_doc()
					break
				else:
					raise ValueError(f"illegal char: {c} at {f.tell()}")

	def read_control(self, f: BinaryIO):
		word = read_word(f)
		if word:
			if esc := ESCAPE.get(word):
				end_control(f)
				self.dest.write(esc)
			else:
				param = read_number(f)
				end_control(f)
				# \u is sort of a control word but we handle it separately because it advances the reader
				if word == 'u':
					self.read_unicode(f, param)
				else:
					self.handle_control(word, param)
		else:
			c = f.read(1)
			# using bytes rather than strings here!
			if c == b"'":
				encoding = rtfcharset.get_encoding(self.current_font.charset, self.charset)
				self.dest.write(bytes([int(f.read(2), 16)]).decode(encoding))
			elif c in META_CHARS:
				self.dest.write(c.decode(ASCII))
			elif special := SPECIAL.get(c):
				self.dest.write(special)
			elif is_endline(c):
				self.dest.par()
			elif c == b'*':
				self.try_read_dest(f)
			else:
				raise ValueError(f"{c} at {f.tell()}")

	def read_unicode(self, f: BinaryIO, param: int):
		# rtf params are supposed to be signed 16-bit, so convert to their unsigned value.
		# but we'll accept larger positive numbers if that's what's on offer
		unsigned = param if param >= 0 else param + 0x10000
		if 0xD800 <= unsigned < 0xDC00:
			self.skip_replacement(f)
			# in the case of a high surrogate, assume another \u follows
			consume(f, b'\\u')
			# ensure low surrogate?
			low = read_number(f)
			end_control(f)
			self.dest.write(struct.pack('hh', param, low).decode('utf-16le'))
		else:
			self.dest.write(chr(unsigned))

		# always skip replacement chars
		self.skip_replacement(f)

	def skip_replacement(self, f: BinaryIO):
		skip_chars(f, self.prop.get('uc', 1))

	def handle_control(self, word: str, param: Optional[int]):
		if instr := getattr(self, '_' + word, None):
			if param is None:
				instr()
			else:
				instr(param)
			return

		if param is None:
			param = True

		if word in TOGGLE:
			self.toggle(word, param)
		elif word.startswith('q'):  # alignment
			self.prop['q'] = word[1:]
		elif word.startswith('ul'):
			self.prop['ul'] = word[2:] or True
		elif word in NUMBERING_STYLES:
			self.numbering.style = word
		elif word.startswith('pn'):
			# TODO: toggling?
			self.numbering.prop[word[2:]] = param
		elif word in UNSUPPORTED_DEST:
			self.dest = NULL_DEVICE
		elif charset := CHARSETS.get(word):
			self.charset = charset
		elif word in FONT_FAMILIES:
			# this property name is made up, maybe use sentinel object instead?
			self.prop['family'] = word[1:]
		elif word in TEXT_INFO:
			self.dest = TextSetter(self.info, word)
		elif word in DATE_INFO:
			self.dest = TimeSetter(self, self.info, word)
		elif word not in IGNORE_WORDS:
			self.prop[word] = param
		# else: pass  # ignore

	def toggle(self, word: str, param: int):
		# TODO: is toggling even a special case? If I just accept falsy values everything should still work.
		#  Could I use ChainMaps in that case?
		if param == 0:
			self.prop.pop(word, None)
		else:
			self.prop[word] = True

	def reset(self, properties: Iterable[str]):
		for name in properties:
			self.prop.pop(name, None)

	def try_read_dest(self, f: BinaryIO):
		read_while(f, is_endline)
		consume(f, b'\\')
		word = read_word(f)
		if instr := getattr(self, '_' + word, None):
			instr()
		else:
			self.dest = NULL_DEVICE

	@property
	def prop(self):
		return self.group.prop

	@property
	def dest(self):
		return self.group.dest

	@dest.setter
	def dest(self, value: Destination):
		self.group.dest = value

	@property
	def current_font(self):
		return self.fonts[self.prop.get('f', self.deff)]

	# INSTRUCTION TABLE: Methods prefixed with an underscore correspond to RTF control words, which we'll look up at
	# runtime. This has downsides (theoretical name collision), but it is simple!
	# Alternatively, create a "controlword" decorator which adds the function to a dict of valid control words.

	def _rtf(self, version=1):
		self.dest = self.output
		self.rtf_version = version

	def _ansicpg(self, page: int):
		self.charset = f"cp{page}"

	def _deff(self, n):
		self.deff = n

	def _fonttbl(self):
		self.dest = FontTable(self)

	def _colortbl(self):
		self.dest = ColorTable(self)

	def _par(self):
		self.dest.par()

	def _page(self):
		self.dest.page_break()

	def _ql(self):
		self.prop.pop('q', None)

	def _ulnone(self):
		self.prop.pop('ul', None)

	def _nosupersub(self):
		self.prop.pop('super', None)
		self.prop.pop('sub', None)

	def _nowidctlpar(self):
		self.prop.pop('widctlpar', None)

	def _pard(self):
		self.reset(PARFMT)
		if self.numbering is not None:
			self.output.numbering_off(self.numbering)
			self.numbering = None

	def _plain(self):
		self.reset(CHRFMT)
		# use actual font obj?
		self.prop['f'] = self.deff

	def _pntext(self):
		self.dest = PlainText(self.output)

	def _info(self):
		pass  # check for info group validity?

	# LISTS

	def _pn(self):
		self.numbering = self.dest = Numbering(self)

	def _pnf(self, n: int):
		self.numbering.font_index = n

	def _pnstart(self, n: int):
		self.numbering.start = n

	def _pnindent(self, n: int):
		self.numbering.indent = n

	def _pnlvl(self, n: int):
		self.numbering.level = n

	def _pnlvlbody(self):
		self._pnlvl(10)

	def _pnlvlblt(self):
		self._pnlvl(11)

	def _pntxtb(self):
		self.dest = TextSetter(self.numbering, 'before')

	def _pntxta(self):
		self.dest = TextSetter(self.numbering, 'after')

	def _bin(self, n: int):
		# TODO: this and other keywords that take 32-bit integers?
		pass

	def _result(self):
		# TODO: handle objects?
		self.dest = NULL_DEVICE

	def _field(self):
		self.dest = Field(self.output)

	def _fldinst(self):
		self.dest = self.dest.set_instruction

	def _fldrslt(self):
		self.dest = self.dest.set_result

	# TODO: \sect / \sectd


class Output(Destination, ABC):

	def plain_text(self, text: str):
		pass

	def hyperlink(self, text, url):
		pass

	def numbering_on(self, info: Numbering):
		pass

	def numbering_off(self, info: Numbering):
		pass

	def end_doc(self):
		pass


class Handler(Output):

	def __init__(self, doc):
		self._doc = doc

	@property
	def prop(self):
		return self._doc.prop

	@property
	def fonts(self):
		return self._doc.font_table.fonts

	@property
	def font(self):
		return self._doc.current_font

	def get_color(self, i):
		return BLACK if not self.colors else self.colors[i]

	@property
	def colors(self):
		return self._doc.color_table.colors

	@property
	def color_foreground(self):
		return self.get_color(self.prop.get('cf', 0))

	@property
	def color_background(self):
		return self.get_color(self.prop.get('cb', 0))

	@property
	def bold(self):
		return self.prop.get('b', False)

	@property
	def italic(self):
		return self.prop.get('i', False)

	@property
	def underline(self):
		return self.prop.get('u', False)

	@property
	def alignment(self):
		# TODO: make enum?
		return self.prop.get('q', 'l')

	@property
	def font_size(self):
		return self.prop.get('fs')

	@property
	def numbering(self):
		return self._doc.numbering

	# TODO: line spacing?


if __name__ == '__main__':

	# card, ord, and ordt not supported
	HTML_LIST_TYPES = {'pndec': '1',  'pnucltr': 'A', 'pnucrm': 'I', 'pnlcltr': 'a', 'pnlcrm': 'i', }

	def diff_prop(old, new):
		diffs = []
		for k in old.keys() & new.keys():
			oldval = old[k]
			newval = new[k]
			if oldval != newval:
				diffs.append(f"{k}: {oldval}->{newval}")
		for k in old.keys() - new.keys():
			if k in TOGGLE:
				diffs.append(f"+{k}")
			else:
				diffs.append(f"{k}: -{old[k]}")
		for k in new.keys() - old.keys():
			if k in TOGGLE:
				diffs.append(f"-{k}")
			else:
				diffs.append(f"{k}: +{new[k]}")
		return '; '.join(diffs)

	class Recorder(Handler):

		def __init__(self, doc):
			super().__init__(doc)
			self.paragraphs = []
			self.current = []
			self.last_prop = {}
			self.styles = set()
			self.par_num = 1
			self.stack = deque()

		def write(self, text):
			self.current.append(text)
			# TODO: the dilemma: handle prop changes as events, or continue doing prop diffs
			print(diff_prop(self.prop, self.last_prop), text)
			self.last_prop = self.prop.copy()

		def par(self):
			line = ''.join(self.current)
			if self._doc.numbering:
				print(f"<li>{line}</li>")
			else:
				print(f"<p>{line}</p>")
			self.paragraphs.append(line)
			self.current = []
			self.par_num += 1

		def numbering_on(self, info: Numbering):
			print(f'<ol type="{HTML_LIST_TYPES[info.style]}">')

		def numbering_off(self, info):
			print('</ol>')

		def hyperlink(self, text, url):
			print(f'<a href="{url}">{text}</a>')

		def end_doc(self):
			pass

	rtf = Parser(Recorder)
	rtf.parse('testdocs/generated.rtf')
