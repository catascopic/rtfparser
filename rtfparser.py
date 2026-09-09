from __future__ import annotations

import os
import struct

import rtfcharset
import rtffields

from abc import ABC, abstractmethod
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime
from types import SimpleNamespace
from typing import BinaryIO

# This parser makes a questionable, but I believe justified, decision to parse files in binary mode.
# RTF files are pure ascii, so it sort of doesn't matter. Strings are generally easier to work with in python,
# but using binary mode lets us call seek() on the reader, which is useful when we want to "peek" at the next token in
# the stream. (We also get to use bytearray a few times and avoid having to do ''.join on lists of chars.)
# The one awkward part is, in python, characters are 1-char strings, while individual bytes are just ints.
# We don't want to use ints, because they make the code hard to read, so we instead use byte sequences of length 1.
# This isn't ideal, but it's really the only downside to this paradigm (other than having to remember to start your
# strings with the b prefix in that section of the code).

# TODO: \upr, \ud (these are only used for not-output destinations)

ASCII = 'ascii'

CHARSETS = {
	'ansi': 'cp1252',
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
	'ltrch', 'cs', 'cchs', 'lang'
} | TOGGLE)

# TABS = { ... }

INFO_PROPS = frozenset({'version', 'edmins', 'nofpages', 'nofwords', 'word_count', 'nofchars', 'nofcharsws'})
TEXT_INFO = frozenset({'title', 'subject', 'author', 'manager', 'company', 'operator', 'category', 'keywords', 'comment', 'doccomm', 'hlinkbase'})
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
	b'~': '\N{NO-BREAK SPACE}',
	b'-': '\N{SOFT HYPHEN}',
	b'_': '\N{NON-BREAKING HYPHEN}',
}

# can't do frozenset(b'\\{}') because iterating bytes gives you ints
META_CHARS = frozenset({b'\\', b'{', b'}'})

IGNORE_WORDS = frozenset({'nouicompat', 'viewkind'})
# nonshppict is a legacy copy of the picture in the \*\shppict group right before it, so we skip it deliberately
UNSUPPORTED_DEST = frozenset({'filetbl', 'stylesheet', 'listtables', 'revtbl', 'nonshppict'})

# The blip (binary large image) types a \pict group can declare. Some of these also take a param
# (a mapping mode or metafile type), which we don't need, since the format alone identifies the data.
BLIPS = {
	'pngblip':    'png',
	'jpegblip':   'jpeg',
	'emfblip':    'emf',
	'macpict':    'pict',
	'wmetafile':  'wmf',
	'pmmetafile': 'pmf',
	'dibitmap':   'dib',
	'wbitmap':    'bmp',
}

BytePredicate = Callable[[bytes], bool]


class RtfWarning(ValueError):
	# something the file did that we don't model or didn't expect. Reported through Output.warning
	# normally, raised when the parser is strict. Subclasses ValueError to match the rest of the
	# parser's errors -- it has nothing to do with the stdlib warnings module.

	def __init__(self, message: str, position: int | None = None):
		super().__init__(message if position is None else f"{message} at {position}")
		self.message = message
		self.position = position


class Destination(ABC):

	def write(self, text: str):
		raise ValueError(f"{type(self)} can't handle text: {text}")

	def write_bin(self, data: bytes):
		# a destination's content isn't always text: \bin gives us raw bytes, which must not be decoded
		raise ValueError(f"{type(self)} can't handle binary data: {len(data)} bytes")

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
	
	def __init__(self):
		self.open = True
	
	def write(self, text):
		if not self.open:
			raise ValueError(f"Root already written but got {text}")
		if text != '\N{NULL}':
			raise ValueError(f"expected NULL but got {text}")
			self.open = False


@dataclass(frozen=True)
class Font:
	name: str
	family: str
	charset: str | None = None


# stands in for a font a document references but never defined. Its charset is None, which means
# \'hh bytes in it decode with the document's charset -- the same thing \fcharset1 asks for.
DEFAULT_FONT = Font('', 'nil')


class FontTable(Destination):

	def __init__(self, doc: Parser):
		self.doc = doc
		self.name = []

	def write(self, text):
		self.name.append(text)
		if text.endswith(';'):
			name = ''.join(self.name)[:-1]
			charset = self.doc.prop.get('fcharset')
			# checked here, where a font is defined once, rather than at every \'hh that uses it
			if charset is not None and charset not in rtfcharset.CHARSETS:
				self.doc.warn(f"unknown charset {charset} for font {name!r}")
			self.doc.fonts[self.doc.prop['f']] = Font(name, self.doc.prop['family'], charset)
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
		if text != ';':
			self.doc.warn(f"unexpected text in color table: {text!r}")
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
		return doc.get_font(index)

	def close(self):
		self.doc.output.numbering_on(self)


class Picture(Destination):
	# The data is hex by default, or raw bytes if it arrives through \bin. The size properties live in the
	# group's prop dict like any other property, so we snapshot them on close, while that group is still current.

	def __init__(self, doc: Parser):
		self.doc = doc
		self.format = None
		self.data = bytearray()
		self.width = None  # in the source's own units, which is pixels for bitmaps
		self.height = None
		self.goal_width = None  # the size it wants to be displayed at, in twips
		self.goal_height = None
		self.scale_x = 100  # percent
		self.scale_y = 100

	def write(self, text):
		self.data.extend(bytes.fromhex(text))

	def write_bin(self, data: bytes):
		self.data.extend(data)

	def close(self):
		prop = self.doc.prop
		self.width = prop.get('picw')
		self.height = prop.get('pich')
		self.goal_width = prop.get('picwgoal')
		self.goal_height = prop.get('pichgoal')
		self.scale_x = prop.get('picscalex', 100)
		self.scale_y = prop.get('picscaley', 100)
		self.doc.output.picture(self)


class Text(Destination):
	# a destination whose content is plain text, kept for whoever owns it to read

	def __init__(self):
		self.content = []

	def write(self, text):
		self.content.append(text)

	@property
	def text(self):
		return ''.join(self.content)
	
	def __repr__(self):
		return f"Text<{self.text}>"


class Field(Destination):

	def __init__(self, doc: Parser, delegate: Output):
		self.doc = doc
		self.delegate = delegate
		self.instruction = Text()
		self.result = Text()

	def close(self):
		name, _, args = self.instruction.text.strip().partition(' ')
		parser = rtffields.PARSERS.get(name)
		handler = getattr(self.delegate, name.lower(), None) if name else None
		if parser is None or handler is None:
			# a field type we don't model -- PAGE, DATE and TOC are all over real documents.
			# its result is whatever the writer last rendered for it, which is the best thing
			# we have to show, so pass that through as ordinary text rather than dropping it.
			self.doc.warn(f"unsupported field instruction: {name or '(empty)'}")
			if self.result.text:
				self.delegate.write(self.result.text)
			return
		handler(self.result.text, parser.parse(args))


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
		try:
			return datetime(
				*(self.doc.prop[k] for k in ('yr', 'mo', 'dy')), 
				*(self.doc.prop.get(k, 0) for k in ('hr', 'min', 'sec'))
			)
		except ValueError:
			return None


class NullDevice(Destination):

	def write(self, text):
		pass

	def write_bin(self, data):
		pass

	def par(self):
		pass

	def page_break(self):
		pass

	def close(self):
		pass

	def __getattr__(self, name):
		# any miscellaneous property is assumed to be a sub-destination
		return self

	def __setattr__(self, name, value):
		# ignores setting special properties, like `format` for pictures
		pass


NULL_DEVICE = NullDevice()


@dataclass
class Info:
	title: str | None = None
	subject: str | None = None
	author: str | None = None
	manager: str | None = None
	company: str | None = None
	operator: str | None = None
	category: str | None = None
	keywords: str | None = None
	comment: str | None = None
	doccomm: str | None = None
	hlinkbase: str | None = None
	creatim: datetime | None = None
	revtim: datetime | None = None
	printim: datetime | None = None
	buptim: datetime | None = None


@dataclass
class Group:
	parent: Group | None
	own_dest: Destination | None
	# TODO: do we want to include string values here?
	prop: dict[str, str | int | bool]

	@classmethod
	def root(cls):
		return cls(None, RootDest(), {})

	def open(self):
		# could use ChainMap here, but we'd have to replace pop() with setting to None/0.
		# seems simpler to just copy the map, since it also reduces lookup time
		return Group(self, None, self.prop.copy())
	
	def close(self):
		if self.own_dest is not None:
			self.own_dest.close()
		return self.parent

	@property
	def dest(self):
		return self.own_dest or self.parent.dest

	@dest.setter
	def dest(self, value):
		self.own_dest = value


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


def read_number(f, default: int | None = None):
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


def call(instr: Callable, param: int | None):
	# the instruction table is keyed by name alone, so we don't know a word's arity up front
	# TODO: error message if arity doesn't match
	if param is None:
		instr()
	else:
		instr(param)


class Parser:

	def __init__(self, output: type[Output], strict: bool = False):
		self.output = output(self)
		# strict turns every warning into an RtfWarning, which is how you find out what a corpus
		# contains that this parser doesn't model. Leave it off to read documents in the wild.
		self.strict = strict
		self.file: BinaryIO | None = None
		self.group = Group.root()
		self.rtf_version = 1
		self.charset: str | None = None
		self.deff: int | None = None
		self.fonts: dict[int, Font] = {}
		self.colors: list[Color] = []
		self.info = Info()
		self.numbering: Numbering | None = None

	def parse(self, file: str | bytes | os.PathLike):
		with open(file, 'rb') as f:
			# kept on the parser so warnings raised deep in a destination can still say where we are
			self.file = f
			try:
				self.read_all(f)
			finally:
				self.file = None

	def read_all(self, f: BinaryIO):
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
				# a few control words consume raw bytes from the stream, so they can't go through handle_control.
				if word == 'u':
					self.read_unicode(f, param)
				elif word == 'bin':
					self.read_bin(f, param)
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

	def read_bin(self, f: BinaryIO, n: int):
		# the next n bytes are raw data rather than rtf. \bin with no param means no data at all
		self.dest.write_bin(f.read(n))

	def skip_replacement(self, f: BinaryIO):
		skip_chars(f, self.prop.get('uc', 1))

	def handle_control(self, word: str, param: int | None):
		if instr := getattr(self, '_' + word, None):
			call(instr, param)
			return

		if param is None:
			param = True

		if word in TOGGLE:
			self.toggle(word, param)
		elif blip := BLIPS.get(word):
			# this has to come before the prefix matches below, which would otherwise claim \pngblip for numbering
			self.dest.format = blip
		elif word.startswith('q'):  # alignment
			self.prop['q'] = word[1:]
		elif word.startswith('ul'):
			self.prop['ul'] = word[2:] or True
		elif word in NUMBERING_STYLES:
			self.set_numbering('style', word)
		elif word.startswith('pn'):
			# TODO: toggling?
			if self.numbering is not None:
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
		# a \* destination is still a control word: it can take a param (\*\pnseclvl3) and it owns
		# the delimiter after it, which would otherwise turn up as text in the destination we open
		param = read_number(f)
		end_control(f)
		if instr := getattr(self, '_' + word, None):
			call(instr, param)
		else:
			self.dest = NULL_DEVICE

	@property
	def position(self) -> int | None:
		# how far we've read, which is just past the construct being complained about,
		# since its control word has already been consumed
		return None if self.file is None else self.file.tell()

	def warn(self, message: str):
		# the one funnel for "the file did something we don't model or didn't expect"
		if self.strict:
			raise RtfWarning(message, self.position)
		self.output.warning(message, self.position)

	@property
	def prop(self):
		return self.group.prop

	@property
	def dest(self):
		return self.group.dest

	@dest.setter
	def dest(self, value: Destination):
		# once we've decided to skip a group, nothing inside it can open a destination of its own.
		# this is what makes \nonshppict skip the \pict it wraps, rather than just its text
		if self.dest is not NULL_DEVICE:
			self.group.dest = value

	def get_font(self, index: int | None) -> Font:
		# documents reference fonts they never put in the table, and \deff can be missing entirely.
		# a font only steers decoding and styling, so a stand-in beats aborting the whole parse.
		return self.fonts.get(index, DEFAULT_FONT)

	@property
	def current_font(self) -> Font:
		return self.get_font(self.prop.get('f', self.deff))

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
		# reset() already dropped f, and current_font falls back to deff on its own,
		# so only write it back when there's an actual default font to name
		if self.deff is not None:
			self.prop['f'] = self.deff

	def _pntext(self):
		self.dest = PlainText(self.output)

	def _info(self):
		pass  # check for info group validity?

	# LISTS

	def _pn(self):
		self.numbering = self.dest = Numbering(self)

	# \pn* words aren't confined to a {\*\pn} group: \pnseclvl and friends describe section
	# numbering we don't model, and a malformed file can put any of them anywhere. With no
	# Numbering to set them on there's nothing meaningful to do, so we drop them.

	def set_numbering(self, attr: str, value):
		if self.numbering is not None:
			setattr(self.numbering, attr, value)
		else:
			self.warn_stray_numbering(attr)

	def numbering_text(self, attr: str) -> Destination:
		if self.numbering is None:
			self.warn_stray_numbering(attr)
			return NULL_DEVICE
		return TextSetter(self.numbering, attr)

	def warn_stray_numbering(self, attr: str):
		# inside a destination we're skipping ({\*\pnseclvl3 ...} and friends) this is expected and
		# not worth reporting. Out here it means the file put a \pn word somewhere it doesn't belong.
		# We only have the property name here, not the control word that set it -- the \pn* methods
		# are dispatched by name, so the word itself is gone by the time we get called.
		if self.dest is not NULL_DEVICE:
			self.warn(f"numbering property set outside a numbering group: {attr}")

	def _pnf(self, n: int):
		self.set_numbering('font_index', n)

	def _pnstart(self, n: int):
		self.set_numbering('start', n)

	def _pnindent(self, n: int):
		self.set_numbering('indent', n)

	def _pnlvl(self, n: int):
		self.set_numbering('level', n)

	def _pnlvlbody(self):
		self._pnlvl(10)

	def _pnlvlblt(self):
		self._pnlvl(11)

	def _pntxtb(self):
		self.dest = self.numbering_text('before')

	def _pntxta(self):
		self.dest = self.numbering_text('after')

	def _result(self):
		# TODO: handle objects?
		self.dest = NULL_DEVICE

	def _field(self):
		self.dest = Field(self, self.output)

	def _fldinst(self):
		self.dest = self.dest.instruction

	def _fldrslt(self):
		self.dest = self.dest.result

	def _pict(self):
		self.dest = Picture(self)

	def _shppict(self):
		pass  # a transparent wrapper around \pict, so leave the destination to the group inside it

	# TODO: \sect / \sectd

	def _loch(self):
		"""TODO"""
	
	def _hich(self):
		"""TODO"""
	
	def _dbch(self):
		"""TODO"""


class Output(Destination, ABC):

	def plain_text(self, text: str):
		pass

	def warning(self, message: str, position: int | None = None):
		# the document did something we don't model. Override to collect or report these;
		# pass strict=True to Parser to have them raised instead.
		pass

	def hyperlink(self, text, args):
		pass

	def includepicture(self, text, args):
		# an INCLUDEPICTURE field, i.e. a reference to an external image.
		pass

	def picture(self, pic: Picture):
		# an image embedded in the document, as pic.data bytes in pic.format
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
		return self._doc.fonts

	@property
	def font(self):
		return self._doc.current_font

	def get_color(self, i):
		return BLACK if not self.colors else self.colors[i]

	@property
	def colors(self):
		return self._doc.colors

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
