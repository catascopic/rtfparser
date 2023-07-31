from __future__ import annotations

import re
import sys

from abc import ABC
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional


# TODO: \upr, \ud


if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path(r'test.rtf')

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
	'q',  # use q to keep track of alignment
	'fi', 'li', 'ri',  # indentation
	'sb', 'sa', 'sl', 'slmult',  # spacing
	'subdocument',
	'rtlpar', 'ltrpar',  # direction
})

# Character formatting (reset with \plain)
TOGGLE = frozenset({'b', 'caps', 'deleted', 'i', 'outl',
                    'scaps', 'shad', 'strike', 'ul', 'v'})  # hyphpar
CHRFMT = frozenset({
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'expnd', 'expndtw',
	'kerning', 'f', 'fs', 'strikedl', 'up', 'super', 'cf', 'cb', 'rtlch',
	'ltrch', 'cs', 'cchs', 'lang'} | TOGGLE)

# TABS = { ... }

# TIME_UNITS = frozenset({'yr', 'mo', 'dy', 'hr', 'min', 'sec'})

INFO_PROPS = frozenset({'version', 'edmins', 'nofpages', 'nofwords',
                        'word_count', 'nofchars', 'nofcharsws'})

TEXT_INFO = frozenset({'title', 'subject', 'author', 'manager', 'company', 'operator',
                       'category', 'keywords', 'comment', 'doccomm', 'hlinkbase'})
DATE_INFO = frozenset({'creatim', 'revtim', 'printim', 'buptim'})

LIST_STYLES = frozenset({'pncard', 'pndec', 'pnucltr', 'pnucrm',
                         'pnlcltr', 'pnlcrm', 'pnord', 'pnordt'})

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
UNSUPPORTED_DEST = frozenset({'filetbl', 'stylesheet', 'listtables', 'revtbl'})


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


class RootDest(Destination):
	# TODO: in theory this should only allow one write
	def write(self, text):
		if text != '\u0000':
			raise ValueError(f"expected NUL but got {text}")


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


class ListType(Destination):

	def __init__(self, doc: Parser):
		self.doc = doc
		self.style = None
		self.level = 0
		self.before = None
		self.after = None
		self.font_index = 0
		self.indent = 0
		self.start = 1

	def close(self):
		self.font_index = self.doc.prop.get('pnf', 0)
		self.start = self.doc.prop.get('pnstart', 1)
		self.indent = self.doc.prop.get('pnindent', 0)

	@property
	def font(self):
		return self.doc.font_table.fonts[self.font_index]


class SetValue(Destination, ABC):

	def __init__(self, obj, prop):
		self.obj = obj
		self.prop = prop

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
		return datetime(*(self.doc.prop[k] for k in ['yr', 'mo', 'dy', 'hr', 'min', 'sec']))


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
	# TODO: string values OK?
	prop: dict[str, str | int | bool]

	@classmethod
	def root(cls):
		return cls(None, RootDest(), {})

	def open(self):
		# TODO: chain map? if so have to use set None instead of pop everywhere
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
					self.group = self.group.open()
				elif c == b'}':
					self.group = self.group.close()
				else:
					# must be EOF
					break

	def read_control(self, f):
		word = read_word(f)
		if word:
			if s := ESCAPE.get(word):
				self.dest.write(s)
			else:
				param = read_number(f)
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
	
	def control(self, word: str, param: Optional[int]):
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
		elif word in LIST_STYLES:
			self.list_type.style = word
		elif word in UNSUPPORTED_DEST:
			self.dest = NULL_DEVICE
		elif word in CHARSETS:
			self.charset = word
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

	def toggle(self, word: str, param):
		if not param:
			self.prop.pop(word, None)
		else:
			self.prop[word] = True

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

	def _deff(self, n):
		self.deff = self.prop['f'] = n

	def _info(self):
		self.dest = self.info

	# LISTS

	def _pn(self):
		self.dest = self.list_type = ListType(self)

	def _pnlvl(self, n):
		self.list_type.level = n

	def _pnlvlbody(self):
		self._pnlvl(10)

	def _pnlvlblt(self):
		self._pnlvl(11)

	def _pntxtb(self):
		self.dest = TextSetter(self.list_type, 'before')

	def _pntxta(self):
		self.dest = TextSetter(self.list_type, 'after')

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

	@property
	def list_type(self):
		return self._doc.list_type
	
	# TODO: line spacing?


def diff_prop(old, new):
	diffs = []
	for k in old.keys() & new.keys():
		oldval = old[k]
		newval = new[k]
		if oldval != newval:
			diffs.append(f"{k}: {oldval}->{newval}")
	for k in old.keys() - new.keys():
		diffs.append(f"{k}: -{old[k]}")
	for k in new.keys() - old.keys():
		diffs.append(f"{k}: +{new[k]}")
	return '; '.join(diffs)


class Recorder(Output):

	def __init__(self, doc):
		super().__init__(doc)
		self.full_text = []
		self.last_prop = {}
		self.styles = set()

	def write(self, text):
		if self.prop != self.last_prop:
			print(diff_prop(self.last_prop, self.prop))
			self.last_prop = self.prop.copy()
			self.styles.add(frozenset((k, v) for k, v in self.prop.items() if k in PARFMT or k in CHRFMT))
		#  print(text, end='')
		print(text)
		self.full_text.append(text)

	def par(self):
		self.full_text.append('\n')

	def page_break(self):
		pass


rtf = Parser(Recorder)
rtf.parse(file)


# print([f.name for f in rtf.font_table.fonts.values()])
full_text = ''.join(rtf.output.full_text)
print(full_text)
# (?:[^\W_]|['‘’\-])
WORDS = re.compile(r"[^\W_][\w'‘’\-]*")
words = WORDS.findall(full_text)
print(f'words: {len(words)}, chars: {len(full_text)}')

it = iter(rtf.output.styles)
common = set(next(it))
for s in it:
	common = common.intersection(s)

print('COMMON:', common)

styles = [dict(s) for s in rtf.output.styles]

for s in rtf.output.styles:
	print(dict(s - common))
