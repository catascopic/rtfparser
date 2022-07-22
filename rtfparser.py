from __future__ import annotations

import re
import sys

from abc import ABC
from dataclasses import dataclass
from collections import Counter
from pathlib import Path
from typing import Union


# TODO: \upr, \ud


if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path('test.rtf')

ASCII = 'ascii'
CHARSETS = {'ansi', 'mac', 'pc', 'pca'}
FONT_FAMILIES = {
	'fnil', 'froman', 'fswiss', 'fmodern', 
	'fscript', 'fdecor', 'ftech', 'fbidi'
}
# Paragraph formatting (reset with \pard)
PARFMT = {
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
}

# Character formatting (reset with \plain)
TOGGLE = {'b', 'caps', 'deleted', 'i', 'outl', 'scaps', 'shad', 'strike', 'ul', 'v'}  # hyphpar
CHRFMT = {
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'expnd', 'expndtw',
	'kerning', 'f', 'fs', 'strikedl', 'up', 'super', 'cf', 'cb', 'rtlch',
	'ltrch', 'cs', 'cchs', 'lang',
} | TOGGLE

# TABS = { ... }

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

META_CHARS = {b'\\', b'{', b'}'}

IGNORE_WORDS = {'nouicompat', 'viewkind'}

NEWLINE = re.compile(r"[\r\n]")


class Destination(ABC):

	def write(self, text):
		return NotImplemented

	def par(self):
		self.write('\n')
	
	def page_break(self):
		pass


@dataclass
class Font:
	name: str
	family: str
	charset: str


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


@dataclass		
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
			raise ValueError(text)


class NullDevice(Destination):
	def write(self, text):
		pass  # do nothing


@dataclass
class Group:
	parent: Group
	dest: Destination
	prop: dict[str, Union[int, bool]]

	@classmethod
	def root(cls):
		return cls(None, NullDevice(), {})

	def make_child(self):
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
		return int(buf) if buf else default
	f.seek(-1, 1)


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


control_words = {}

def controlword(func):
	control_words[func.__name__] = func
	return func
	

class Parser:

	def __init__(self, output, plain_text=False):
		self.output = output(self)
		self.font_table = FontTable(self)
		self.color_table = ColorTable(self)
		self.group = Group.root()
		self.plain_text = plain_text
		self.charset = 'ansi'
		self.deff = None
		self.rtf_version = None

	def parse(self, file):
		with open(file, 'rb') as f:
			while True:
				text = NEWLINE.sub('', read_while(f, not_control).decode(ASCII))
				if text:
					self.dest.write(text)
				c = f.read(1)
				if c == b'\\':
					self.read_control(f)
				elif c == b'{':
					self.group = self.group.make_child()
				elif c == b'}':
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
				# TODO: fix this
				self.dest = NullDevice()
			else:
				raise ValueError(f"{c} at {f.tell()}")
	
	def control(self, word, param):	
		if instr := control_words.get(word):
			args = (param,) if type(param) is int else ()
			instr(self, param)
		elif word in TOGGLE:
			self.toggle(word, param)
		elif word.startswith('q'):  # alignment
			self.prop['q'] = word[1:]
		elif word.startswith('ul'):
			self.prop['ul'] = word[2:]
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl'}:
			# these destinations are unsupported
			self.dest = NullDevice()
		elif word in CHARSETS:
			self.charset = word
		elif word in FONT_FAMILIES:
			# this property name is made up
			self.prop['family'] = word[1:]
		elif word not in IGNORE_WORDS:
			self.prop[word] = param

	@controlword
	def par(self):
		self.dest.par()
	@controlword
	def page(self):
		self.dest.page_break()
	@controlword
	def ql(self):
		self.prop.pop('q', None)
	@controlword
	def ulnone(self):
		self.prop.pop('ul', None)
	@controlword
	def nosupersub(self):
		self.prop.pop('super', None)
		self.prop.pop('sub', None)
	@controlword
	def nowidctlpar(self):
		self.prop.pop('widctlpar', None)
	@controlword
	def pard(self):
		self.reset(PARFMT)
		self.list_type = None
	@controlword
	def plain(self):
		self.reset(CHRFMT)
		# use actual font obj?
		self.prop['f'] = self.deff
	@controlword
	def rtf(self, param):
		self.dest = self.output
		self.rtf_version = param
	@controlword
	def fonttbl(self):
		self.dest = self.font_table
	@controlword
	def colortbl(self):
		self.dest = self.color_table
	@controlword
	def pntext(self):
		self.dest = self.output if self.plain_text else NullDevice()
	@controlword
	def deff(self, param):
		self.deff = self.prop['f'] = param
	@controlword
	def pnlvlblt(self):
		pass

	def toggle(self, word, param):
		if param == 0:
			self.prop.pop(word, None)
		else:
			self.prop[word] = param

	def reset(self, properties):
		for name in properties:
			self.prop.pop(name, None)
	
	def try_read_destination(self, f):
		read_while(f, lambda c: c == b'\r' or c == b'\n')
		c = f.read(1)
		if c != b'\\':
			raise ValueError(f"{c} at {f.tell()}")
		word = read_word(f)
		if word == 'pn':
			self.list_type = ListType()
			self.dest = NullDevice()
		else:
			self.dest = NullDevice()

	@property
	def prop(self):
		return self.group.prop

	@property
	def dest(self):
		return self.group.dest

	@dest.setter
	def dest(self, value):
		self.group.dest = value


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

	def write(self, text):
		self.full_text.append(text)


rtf = Parser(Recorder)
rtf.parse(file)


# print([f.name for f in rtf.font_table.fonts.values()])
full_text = ''.join(rtf.output.full_text)
# (?:[^\W_]|['‘’\-])
WORDS = re.compile(r"[^\W_][\w'‘’\-]*")
words = WORDS.findall(full_text)
print(' '.join([w for w, c in Counter(w.lower().strip("'‘’") for w in words).items() if c == 1][-24:]))
print(f'words: {len(words)}, chars: {len(full_text)}')
mark = full_text.rfind('^')
if mark != -1:
	print('since mark:', len(WORDS.findall(full_text[mark + 1:])))

bad_quote_words = re.findall(r"\S+['\"]\S+", full_text)
if bad_quote_words:
	print('BAD QUOTES:', ' '.join(bad_quote_words))
