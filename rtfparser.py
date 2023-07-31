from __future__ import annotations

import re
import sys

from abc import ABC
from dataclasses import dataclass
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional


# TODO: \upr, \ud


if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path('test.rtf')

ASCII = 'ascii'

CHARSETS = frozenset({'ansi', 'mac', 'pc', 'pca'})

FONT_FAMILIES = frozenset({
	'fnil', 'froman', 'fswiss', 'fmodern', 
	'fscript', 'fdecor', 'ftech', 'fbidi'
})

# Paragraph formatting (reset with \pard)
PARFMT = frozenset({
	's', 'hyphpar', 'intbl', 'keep', 'nowidctlpar', 'widctlpar', 
	'keepn', 'level', 'noline', 'outlinelevel', 'pagebb', 'sbys',
	# use q to keep track of alignment
	'q',
	# Indentation
	'fi', 'li', 'ri',
	# Spacing
	'sb', 'sa', 'sl', 'slmult',
	# Subdocuments
	'subdocument',
	# Bidirectional
	'rtlpar', 'ltrpar',
})

# Character formatting (reset with \plain)
TOGGLE = frozenset({'b', 'caps', 'deleted', 'i', 'outl',
                    'scaps', 'shad', 'strike', 'ul', 'v'})  # hyphpar
CHRFMT = frozenset({
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'expnd', 'expndtw',
	'kerning', 'f', 'fs', 'strikedl', 'up', 'super', 'cf', 'cb', 'rtlch',
	'ltrch', 'cs', 'cchs', 'lang'} | TOGGLE)

# TABS = { ... }

TIME_UNITS = frozenset({'yr', 'mo', 'dy', 'hr', 'min', 'sec'})

INFO_PROPS = frozenset({'version', 'edmins', 'nofpages', 'nofwords',
                        'word_count', 'nofchars', 'nofcharsws'})
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
	b'~': u'\u00A0',  # nonbreaking space
	b'-': u'\u00AD',  # optional hyphen
	b'_': u'\u2011',  # nonbreaking hyphen
}

# can't do frozenset(b'\\{}') because iterating bytes gives you ints
META_CHARS = frozenset({b'\\', b'{', b'}'})

IGNORE_WORDS = frozenset({'nouicompat', 'viewkind'})


class Destination(ABC):

	def write(self, text):
		raise ValueError(f"{type(self)} can't handle text: {text}")

	def par(self):
		raise ValueError(f"{type(self)} can't handle paragraphs")
	
	def page_break(self):
		raise ValueError(f"{type(self)} can't handle page breaks")
	
	def close(self):
		pass


class NullDevice(Destination):
	def write(self, text):
		pass  # do nothing


NULL_DEVICE = NullDevice()


@dataclass(frozen=True)
class Font:
	name: str
	family: str
	charset: Optional[str]


class FontTable(Destination):

	def __init__(self, doc):
		self.doc = doc
		self.fonts = {}
		self.name = []

	def write(self, text):
		self.name.append(text)
		if text.endswith(';'):	
			self.fonts[self.doc.prop['f']] = Font(
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
	
	def __init__(self, doc):
		self.doc = doc
		self.colors = []

	def write(self, text):
		if text == ';':
			self.colors.append(Color(
				self.doc.prop.get('red', 0),
				self.doc.prop.get('green', 0),
				self.doc.prop.get('blue', 0)))
		else:
			raise ValueError(f"{text} in color table")


class TextDest(Destination):

	def __init__(self):
		self.content = []
	
	def write(self, text):
		self.content.append(text)

	def get_text(self):
		return ''.join(self.content)


@dataclass(frozen=True)
class TimeDest(Destination):
	yr: int = 0
	mo: int = 0
	dy: int = 0
	hr: int = 0
	min: int = 0
	sec: int = 0

	def to_date(self):
		return datetime(self.yr, self.mo, self.dy, self.hr, self.min, self.sec)


class Info(Destination):

	def __init__(self):
		# self.title = TextDest()
		# self.subject = TextDest()
		self.author = TextDest()
		# self.manager = TextDest()
		# self.company = TextDest()
		# self.operator = TextDest()
		# self.category = TextDest()
		# self.keywords = TextDest()
		# self.comment = TextDest()
		# self.doccomm = TextDest()
		self.create_time = TimeDest()
		self.revision_time = TimeDest()
		self.print_time = TimeDest()
		self.backup_time = TimeDest()


class ListType:

	def __init__(self):
		self.level = 0


def noop():
	pass


@dataclass
class Group:
	parent: Optional[Group]
	dest: Destination
	prop: dict[str, int | bool]
	# TODO: wait, what is the point of this?
	on_close: Callable[[], None] = noop

	@classmethod
	def root(cls):
		return cls(None, NULL_DEVICE, {})

	def make_child(self):
		# TODO: chain map? if so have to use set None instead of pop everywhere
		return Group(self, self.dest, self.prop.copy())


def not_control(c):
	return c not in META_CHARS


# TODO: is mmathPr and other stuff actually valid RTF?
def is_letter(c):
	return b'a' <= c <= b'z' or b'A' <= c <= b'Z'


def is_digit(c):
	return b'0' <= c <= b'9'


def read_while(f, matcher):
	buf = bytearray()
	read_into_while(f, buf, matcher)
	return buf


def read_into_while(f, buf, matcher):
	while True:
		c = f.read(1)
		if not matcher(c):
			f.seek(-1, 1)
			return
		if not c:
			return
		buf.extend(c)


def read_word(f):
	return read_while(f, is_letter).decode(ASCII)


def read_number(f, default=None):
	c = f.read(1)
	if is_digit(c) or c == b'-':
		buf = bytearray(c)
		read_into_while(f, buf, is_digit)
		return int(buf)
	f.seek(-1, 1)
	return default


def consume_end(f):
	c = f.read(1)
	if not c.isspace():
		f.seek(-1, 1)
	# we don't need to handle the CRLF case because we read the CR here and
	# skip the LF later while reading normally


def skip_chars(f, n):
	for _ in range(n):
		if f.read(1) == b'\\':
			if read_word(f):
				read_number(f)  # unnecessary?
				consume_end(f)
			elif f.read(1) == b"'":
				f.read(2)


class Parser:

	def __init__(self, output, plain_text=False):
		self.output = output(self)
		self.font_table = FontTable(self)
		self.color_table = ColorTable(self)
		self.info = Info()
		self.group = Group.root()
		self.plain_text = plain_text
		self.charset = 'ansi'
		self.deff = None
		self.rtf_version = None

	def parse(self, file):
		with open(file, 'rb') as f:
			while True:
				text = read_while(f, not_control).translate(None, b'\r\n').decode(ASCII)
				if text:
					self.dest.write(text)
				c = f.read(1)
				if c == b'\\':
					self.read_control(f)
				elif c == b'{':
					self.group = self.group.make_child()
				elif c == b'}':
					self.dest.close()
					self.group = self.group.parent
				else:
					# must be EOF
					break

	def read_control(self, f):
		word = read_word(f)
		if word:
			if s := ESCAPE.get(word):
				self.dest.write(s)
			else:
				param = read_number(f, True)
				# handle \u separately because it advances the reader
				if word == 'u':
					self.dest.write(chr(param))
					# we can always handle unicode
					skip_chars(f, self.prop.get('uc', 1))
				else:
					self.control(word, param)
			consume_end(f)
		else:
			c = f.read(1)
			# using bytes rather than strings here!
			if c == b"'":
				# can python handle charsets other than ansi?
				self.dest.write(bytes([int(f.read(2), 16)]).decode(self.charset))
			elif c in META_CHARS:
				self.dest.write(c.decode(ASCII))
			elif s := SPECIAL.get(c):
				self.dest.write(s)
			elif c == b'\r' or c == b'\n':
				self.dest.par()
			elif c == b'*':
				self.try_read_dest(f)
			else:
				raise ValueError(f"{c} at {f.tell()}")
	
	def control(self, word, param):	
		if instr := getattr(self, '_' + word, None):
			args = (param,) if type(param) is int else ()
			instr(*args)
		elif word in TOGGLE:
			self.toggle(word, param)
		elif word.startswith('q'):  # alignment
			self.prop['q'] = word[1:]
		elif word.startswith('ul'):
			self.prop['ul'] = word[2:]
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl'}:
			# these destinations are unsupported
			self.dest = NULL_DEVICE
		elif word in CHARSETS:
			self.charset = word
		elif word in FONT_FAMILIES:
			# this property name is made up, maybe use sentinel object instead?
			self.prop['family'] = word[1:]
		elif word in TIME_UNITS:
			if isinstance(self.dest, TimeDest):
				setattr(self.dest, word, param)
			else:
				raise ValueError(f"cannot set time unit {word} for {self.dest} ({type(self.dest)})")
		elif word not in IGNORE_WORDS:
			self.prop[word] = param

	def toggle(self, word, param):
		if param == 0:
			self.prop.pop(word, None)
		else:
			self.prop[word] = param

	def reset(self, properties):
		for name in properties:
			self.prop.pop(name, None)

	def try_read_dest(self, f):
		read_while(f, lambda c: c in b'\r\n')
		c = f.read(1)
		if c != b'\\':
			raise ValueError(f"{c} at {f.tell()}")
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
	def dest(self, value):
		self.group.dest = value
		self.group.on_close = self.dest.close

	# INSTRUCTION TABLE: Methods prefixed with an underscore correspond to RTF control words, which we'll look up at
	# runtime. This may not be ideal, but it's the simplest way. Alternatively, create a "controlword" decorator which
	# adds the function to a dict of valid control words.

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
		self.list_type = None

	def _plain(self):
		self.reset(CHRFMT)
		# use actual font obj?
		self.prop['f'] = self.deff

	def _rtf(self, a):
		self.dest = self.output
		self.rtf_version = a

	def _fonttbl(self):
		self.dest = self.font_table

	def _colortbl(self):
		self.dest = self.color_table

	def _pntext(self):
		self.dest = self.output if self.plain_text else NULL_DEVICE

	def _deff(self, a):
		self.deff = self.prop['f'] = a

	def _info(self):
		self.dest = self.info

	# title, subject, manager, company, operator, category, keywords, comment, doccomm, hlinkbase

	def _author(self):
		self.dest = self.info.author

	def _creatim(self):
		self.dest = self.info.create_time

	def _revtim(self):
		self.dest = self.info.revision_time

	def _printim(self):
		self.dest = self.info.print_time

	def _buptim(self):
		self.dest = self.info.backup_time

	def _pn(self):
		plist = ListType()
		self.list_type = plist
		self.dest = plist

	def _pnlvl(self, n):
		self.list_type.level = n

	def _pnlvlbody(self):
		self._pnlvl(10)

	def _pnlvlblt(self):
		self._pnlvl(11)

	def _pnindent(self, n):
		pass

	def _pntxta(self):
		# TODO: DOES THE GROUP CLOSE THE DEST?
		self.dest = TextDest()

	def _pntextb(self):
		doc = self

		class SetTextBefore(TextDest):
			def close(self):
				doc.list_type.before = self.get_text()

		self.dest = SetTextBefore()

	def _result(self):
		# TODO: handle objects?
		self.dest = NULL_DEVICE


class Output(Destination):

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
		return self.fonts[self.prop['f']]

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
		return self.prop['fs']
	
	# TODO: line spacing?


class Recorder(Output):

	def __init__(self, doc):
		super().__init__(doc)
		self.full_text = []
		self.last_prop = {}

	def write(self, text):
		if self.prop != self.last_prop:
			# print()
			# print('-', self.last_prop.items() - self.prop.items())
			# print('+', self.prop.items() - self.last_prop.items())
			self.last_prop = self.prop.copy()
		#  print(text, end='')
		self.full_text.append(text)

	def par(self):
		self.write('\n')

	def page_break(self):
		pass


rtf = Parser(Recorder, plain_text=True)
rtf.parse(file)


# print([f.name for f in rtf.font_table.fonts.values()])
full_text = ''.join(rtf.output.full_text)
print(full_text)
# (?:[^\W_]|['‘’\-])
WORDS = re.compile(r"[^\W_][\w'‘’\-]*")
words = WORDS.findall(full_text)
print(f'words: {len(words)}, chars: {len(full_text)}')

