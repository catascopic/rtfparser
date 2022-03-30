import re
import sys

from abc import ABC
from dataclasses import dataclass
from pathlib import Path


if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path('test.rtf')


class Group:

	def __init__(self, parent):
		self.parent = parent
		if parent is None:
			self.dest = Root()
	
	def write(self, text):
		self.dest.write(text, self)

	def reset(self, properties):
		for name in properties:
			self.__dict__.pop(name, None)

	# jesus christ this might not be worth it
	def __getattr__(self, name):
		if self.parent is None:
			return None
		return getattr(self.parent, name)


class Destination(ABC):

	def write(self, text, group):
		return NotImplemented


class Output(Destination):

	def __init__(self):
		self.full_text = []

	def write(self, text, group):
		self.full_text.append(text)


@dataclass
class Font:
	name: str
	family: str
	charset: str
	

FONT_FAMILIES = {'fnil', 'froman', 'fswiss', 'fmodern', 
                 'fscript', 'fdecor', 'ftech', 'fbidi'}

class FontTable(Destination):

	def __init__(self):
		self.fonts = {}

	def write(self, text, group):
		# should this be more rigorous?
		name = text.removesuffix(';')
		self.fonts[group.f] = Font(name, group.family, group.fcharset)


@dataclass		
class Color:
	red: int
	green: int
	blue: int


class ColorTable(Destination):
	
	def __init__(self):
		self.colors = []

	def write(self, text, group):
		if text == ';':
			self.colors.append(Color(group.red or 0, group.green or 0, group.blue or 0))
		else:
			raise ValueError(text)


class Root(Destination):
	def write(self, text, group):
		if text != '\x00':
			raise ValueError(text.encode())


class NullDevice(Destination):
	def write(self, text, group):
		pass  # do nothing


CHARSETS = {'ansi', 'mac', 'pc', 'pca'}

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
TOGGLE = {'b', 'caps', 'deleted', 'i', 'outl',  'scaps', 'shad', 'strike', 'ul', 'v'}
CHRFMT = {
	'animtext', 'charscalex', 'dn', 'embo', 'impr', 'sub', 'nosupersub', 
	'expnd', 'expndtw', 'kerning', 'f', 'fs', 'strikedl', 'up',
	'super', 'cf', 'cb', 'rtlch', 'ltrch', 'cs', 'cchs', 'lang',
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


def is_letter(c):
	return b'a' <= c <= b'z' or b'A' <= c <= b'Z'

def is_digit(c):
	return b'0' <= c <= b'9' or c == '-'

def not_control(c):
	return c not in SPECIAL


def read_until(f, matcher):
	start = f.tell()
	while True:
		c = f.read(1)
		if not matcher(c) or not c:
			end = f.tell()
			f.seek(start)
			return f.read(end - start - len(c))


def consume_end(f):
	c = f.read(1)
	if c == b'\r':
		f.read(1)
		# check that this next char is \n?
	elif not c.isspace():
		f.seek(-1, 1)


class RTF:

	def __init__(self, file):
		self.file = file
		self.output = Output()
		self.font_table = FontTable()
		self.color_table = ColorTable()
		self.group = Group(None)

	def read_control(self, f):
		word = read_until(f, is_letter).decode()
		if word:
			if word in ESCAPE:
				self.write(ESCAPE[word])
			else:
				param = read_until(f, is_digit)
				self.set_control(word, int(param) if param else None)
			consume_end(f)
		else:
			c = f.read(1)
			if c == b'*':
				self.group.dest = NullDevice()
			elif c == b"'":
				self.write(chr(int(f.read(2), 16)))
			elif c in SPECIAL:
				self.write(c.decode())
			else:
				raise ValueError(c)
	
	def set_control(self, word, param):
		if word == 'par' or word == 'line':
			self.write('\n')
		elif word in TOGGLE:
			setattr(self.group, word, 1 if param is None else param)
		elif word == 'ulnone':
			self.group.ul = 0
		elif word.startswith('ul'):
			self.group.ul = word[2:]
		elif word.startswith('q'):  # alignment
			self.group.q = word[1:]
		elif word == 'pard':
			self.group.reset(PARFMT)
		elif word == 'plain':
			self.group.reset(CHRFMT)
		elif word == 'rtf':
			self.group.dest = self.output
		elif word == 'fonttbl':
			self.group.dest = self.font_table
		elif word == 'colortbl':
			self.group.dest = self.color_table
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl'}:
			self.group.dest = NullDevice()
		elif word in FONT_FAMILIES:
			self.group.family = word[1:]
		elif word not in IGNORE_WORDS:
			setattr(self.group, word, param)
	
	def write(self, text):
		self.group.write(text)

	def parse(self):
		with open(self.file, 'rb') as f:
			while True:
				text = read_until(f, not_control).replace(b'\r\n', b'')
				if text:
					self.write(text.decode())
				c = f.read(1)
				if c == b'{':
					self.group = Group(self.group)
				elif c == b'}':
					self.group = self.group.parent
				elif c == b'\\':
					self.read_control(f)
				elif not c:
					break


rtf = RTF(file)
rtf.parse()

print({f.name for f in rtf.font_table.fonts.values()})
full_text = ''.join(rtf.output.full_text)
words = re.findall(r"\w[\w'‘’\-]*", full_text)
print(f'words: {len(words)}, chars: {len(full_text)}')

# from collections import Counter
# print(Counter(w.lower() for w in words))
