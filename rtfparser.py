import re
import sys

from abc import ABC
from dataclasses import dataclass
from collections import defaultdict, Counter
from pathlib import Path


if len(sys.argv) > 1:
	file = Path(sys.argv[1])
else:
	file = Path('test.rtf')


class Group:

	def __init__(self, parent=None):
		self.parent = parent
		if parent:
			self.dest = parent.dest
			self.prop = parent.prop.copy()
		else:
			self.dest = NullDevice()
			self.prop = {}
	
	def write(self, text, doc):
		self.dest.write(text, self.prop, doc)
		
	def toggle(self, word, param):
		if param == 0:
			self.prop.pop(word, None)
		else:
			self.prop[word] = True if param is None else param

	def reset(self, properties):
		for name in properties:
			self.prop.pop(name, None)


class Destination(ABC):

	def write(self, text, group, doc):
		return NotImplemented


class Output(Destination):

	def __init__(self):
		self.full_text = []
		self.lastprop = {}
		self.count = Counter()

	def write(self, text, prop, doc):
		self.full_text.append(text)
		if not text.isspace():
			if prop['f'] == 1 and prop['fs'] == 36 and prop.get('q', 'l') == 'c':
				print(self.full_text[-5:])
			self.count[frozenset(prop.items())] += len(text)
			if prop != self.lastprop:
				diff = {k: (self.lastprop.get(k), prop.get(k)) for k in (prop.keys() | self.lastprop.keys())
						if self.lastprop.get(k) != prop.get(k)}
				# print(diff, text)
				self.lastprop = prop.copy()


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
		self.name = None

	def write(self, text, prop, doc):
		if text.endswith(';'):
			name = self.name if text == ';' else text[:-1]
			self.fonts[prop['f']] = Font(name, prop['family'], prop.get('fcharset'))
		else:
			self.name = text


@dataclass		
class Color:
	red: int
	green: int
	blue: int


class ColorTable(Destination):
	
	def __init__(self):
		self.colors = []

	def write(self, text, prop, doc):
		if text == ';':
			self.colors.append(Color(
				prop.get('red', 0),
				prop.get('green', 0),
				prop.get('blue', 0)))
		else:
			raise ValueError(text)


class NullDevice(Destination):
	def write(self, text, prop, doc):
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
TOGGLE = {'b', 'caps', 'deleted', 'i', 'outl', 'scaps', 'shad', 'strike', 'ul', 'v'}
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


def skip_char(f):
	c = f.read(1)
	if c == b'\\':
		if f.read(1) != b"'":
			raise ValueError()
		f.read(2)


class Parser:

	def __init__(self, file):
		self.file = file
		self.output = Output()
		self.font_table = FontTable()
		self.color_table = ColorTable()
		self.group = Group()
		self.charset = 'ansi'
		self.deff = None

	def read_control(self, f):
		word = read_until(f, is_letter).decode()
		if word:
			if word in ESCAPE:
				self.write(ESCAPE[word])
			else:
				param = read_until(f, is_digit)
				param = int(param) if param else None
				if word == 'u':
					self.write(chr(param))
					skip_char(f)
				else:
					self.control(word, param)
			consume_end(f)
		else:
			c = f.read(1)
			if c == b'*':
				self.group.dest = NullDevice()
			elif c == b"'":
				self.write(bytes([int(f.read(2), 16)]).decode('ansi'))
			elif c in SPECIAL:
				self.write(c.decode())
			else:
				raise ValueError(c)
	
	def control(self, word, param):
		group = self.group
		if word == 'par' or word == 'line':
			self.write('\n')
		elif word in TOGGLE:
			group.toggle(word, param)
		elif word == 'ql':
			group.prop.pop('q', None)
		elif word.startswith('q'):  # alignment
			group.prop['q'] = word[1:]
		elif word == 'ulnone':
			group.prop.pop('ul', None)
		elif word.startswith('ul'):
			group.prop['ul'] = word[2:]
		elif word == 'pard':
			group.reset(PARFMT)
		elif word == 'plain':
			group.reset(CHRFMT)
			group.prop['f'] = self.deff
		elif word == 'rtf':
			group.dest = self.output
		elif word == 'fonttbl':
			group.dest = self.font_table
		elif word == 'colortbl':
			group.dest = self.color_table
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl'}:
			group.dest = NullDevice()
		elif word in CHARSETS:
			self.charset = word
		elif word == 'deff':
			self.deff = group.prop['f'] = param
		elif word in FONT_FAMILIES:
			group.prop['family'] = word[1:]
		elif word not in IGNORE_WORDS:
			group.prop[word] = param			
	
	def write(self, text):
		self.group.write(text, self)

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


rtf = Parser(file)
rtf.parse()

print([f.name for f in rtf.font_table.fonts.values()])
full_text = ''.join(rtf.output.full_text)
words = re.findall(r"\w[\w'‘’\-]*", full_text)
print(f'words: {len(words)}, chars: {len(full_text)}')

# from collections import Counter
# print(Counter(w.lower() for w in words))
shared = set.intersection(*(set(s) for s in rtf.output.count.keys()))
# print('shared:', shared)
for k, v in sorted(rtf.output.count.items(), key=lambda a: a[1]):
	print(f'{v:6}: {sorted(k - shared)}')
