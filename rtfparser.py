from __future__ import annotations

import re
import sys

from abc import ABC
from dataclasses import dataclass
from collections import Counter
from pathlib import Path
from typing import Union

if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path('test.rtf')


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
TOGGLE = {'b', 'caps', 'deleted', 'i', 'outl', 'scaps', 'shad', 'strike', 'ul', 'v'}
CHRFMT = {
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'expnd', 'expndtw',
	'kerning', 'f', 'fs', 'strikedl', 'up', 'super', 'cf', 'cb', 'rtlch',
	'ltrch', 'cs', 'cchs', 'lang',
} | TOGGLE

# TABS = { ... }

ESCAPE = {
	'tab':       '\t',
	'emdash':    '—',
	'endash':    '-',
	'lquote':    '‘',
	'rquote':    '’',
	'ldblquote': '“',
	'rdblquote': '”',
	'bullet':    '•',
}

IGNORE_WORDS = {'nouicompat', 'viewkind'}

SPECIAL = {b'\\', b'{', b'}'}

NEWLINE = re.compile(r"[\r\n]")


class Destination(ABC):

	def write(self, text):
		return NotImplemented

	def par(self):
		self.write('\n')

	def line(self):
		self.par()

	def nbsp(self):
		self.write(u'\u00A0')

	def opt_hyphen(self):
		self.write(u'\u00AD')

	def nb_hyphen(self):
		self.write(u'\u2011')


class Output(Destination):

	def __init__(self, doc):
		self.doc = doc


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
	return c not in SPECIAL


def is_letter(c):
	return b'a' <= c <= b'z'


def is_digit(c):
	return b'0' <= c <= b'9' or c == b'-'


def read_while(f, matcher):
	start = f.tell()
	while True:
		c = f.read(1)
		if not matcher(c) or not c:
			end = f.tell()
			f.seek(start)
			return f.read(end - start - len(c))


def read_word(f):
	return read_while(f, is_letter)


def read_number(f):
	c = f.read(1)
	if c == b'-' or is_digit(c):
		return c + read_while(f, is_digit)
	f.seek(-1, 1)


def consume_end(f):
	c = f.read(1)
	if c == b'\r':
		f.read(1)
		# check that this next char is \n?
	elif not c.isspace():
		f.seek(-1, 1)


def skip_chars(f, n=1):
	for _ in range(n):
		if f.read(1) == b'\\':
			if read_word(f):
				read_number(f)
				consume_end(f)
			elif f.read(1) == b"'":
				f.read(2)


class Parser:

	def __init__(self, output):
		self.output = output(self)
		self.font_table = FontTable(self)
		self.color_table = ColorTable(self)
		self.group = Group.root()
		self.charset = 'ansi'
		self.deff = None

	def parse(self, file):
		with open(file, 'rb') as f:
			while True:
				text = NEWLINE.sub('', read_while(f, not_control).decode('ascii'))
				if text:
					self.dest.write(text)
				c = f.read(1)
				if c == b'{':
					self.group = self.group.make_child()
				elif c == b'}':
					self.group = self.group.parent
				elif c == b'\\':
					self.read_control(f)
				elif not c:
					break

	def read_control(self, f):
		word = read_word(f).decode('ascii')
		if word:
			if word in ESCAPE:
				self.dest.write(ESCAPE[word])
			else:
				param = read_number(f)
				param = int(param) if param else True
				if word == 'u':
					# we can always handle unicode
					self.dest.write(chr(param))
					skip_chars(f, self.prop.get('uc', 1))
					# don't do consume_end
					return
				self.control(word, param)
			consume_end(f)
		else:
			c = f.read(1)
			if c == b'*':
				self.dest = NullDevice()
			elif c == b"'":
				# can python handle charsets other than ansi?
				self.dest.write(bytes([int(f.read(2), 16)]).decode(self.charset))
			elif c in SPECIAL:
				self.dest.write(c.decode('ascii'))
			elif c == b'~':
				self.dest.nbsp()
			elif c == b'-':
				self.dest.opt_hyphen()
			elif c == b'_':
				self.dest.nb_hyphen()
			elif c == b':':
				raise ValueError('subentry not handled')
			elif c == '\r' or c == '\n':
				self.dest.par()
			else:
				raise ValueError(c)

	def control(self, word, param):
		if word == 'par':
			self.dest.par()
		elif word == 'line':
			self.dest.line()
		elif word in TOGGLE:
			self.toggle(word, param)
		elif word == 'ql':
			self.prop.pop('q', None)
		elif word.startswith('q'):  # alignment
			self.prop['q'] = word[1:]
		elif word == 'ulnone':
			self.prop.pop('ul', None)
		elif word.startswith('ul'):
			self.prop['ul'] = word[2:]
		elif word == 'nosupersub':
			self.prop.pop('super', None)
			self.prop.pop('sub', None)
		elif word == 'pard':
			self.reset(PARFMT)
		elif word == 'plain':
			self.reset(CHRFMT)
			self.prop['f'] = self.deff
		# use actual font obj?
		elif word == 'rtf':
			self.dest = self.output
		elif word == 'fonttbl':
			self.dest = self.font_table
		elif word == 'colortbl':
			self.dest = self.color_table
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl'}:
			# these destinations are unsupported
			self.dest = NullDevice()
		elif word in CHARSETS:
			self.charset = word
		elif word == 'deff':
			self.deff = self.prop['f'] = param
		elif word in FONT_FAMILIES:
			# this property name is made up
			self.prop['family'] = word[1:]
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

	@property
	def prop(self):
		return self.group.prop

	@property
	def dest(self):
		return self.group.dest

	@dest.setter
	def dest(self, value):
		self.group.dest = value
		
	@property
	def fonts(self):
		return self.font_table.fonts

	@property
	def font(self):
		return self.fonts[self.prop['f']]

	def get_color(self, i):
		return BLACK if not self.colors else self.colors[i]

	@property
	def colors(self):
		return self.color_table.colors

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
# print(sorted(Counter(w.lower().strip("'‘’") for w in words).items(), key=lambda x: -x[1])[128:256])
print(f'words: {len(words)}, chars: {len(full_text)}')
mark = full_text.rfind('^')
if mark != -1:
	print('since mark:', len(WORDS.findall(full_text[mark + 1:])))

bad_quote_words = re.findall(r"\S+['\"]\S+", full_text)
if bad_quote_words:
	print('BAD QUOTES:', ' '.join(bad_quote_words))

# from collections import Counter
# shared = set.intersection(*(set(s) for s in rtf.output.count.keys()))
# print('shared:', shared)
# for k, v in sorted(rtf.output.count.items(), key=lambda a: a[1]):
	# print(f'{v:6}: {sorted(k - shared)}')
