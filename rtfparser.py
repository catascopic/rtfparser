import re
import sys

from abc import ABC
from dataclasses import dataclass
from collections import Counter
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
		
	def par(self, group, doc):
		raise ValueError
		
	def line(self, group, doc):
		raise ValueError

	def nbsp(self, group, doc):
		self.write(u'\u00A0', group, doc)

	def opt_hyphen(self, group, doc):
		self.write(u'\u00AD', group, doc)

	def nb_hyphen(self, group, doc):
		self.write(u'\u2011', group, doc)


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
	
	def par(self, group, doc):
		self.write('\n', group, doc)
	
	def line(self, group, doc):
		self.par(group, doc)


@dataclass
class Font:
	name: str
	family: str
	charset: str


class FontTable(Destination):

	def __init__(self):
		self.fonts = {}
		self.name = []

	def write(self, text, prop, doc):
		self.name.append(text)
		if text.endswith(';'):	
			self.fonts[prop['f']] = Font(
				''.join(self.name)[:-1], 
				prop['family'], 
				prop.get('fcharset'))
			self.name = []


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


def read_while(f, matcher):
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


def skip_chars(f, n=1):
	for _ in range(n):
		c = f.read(1)
		if c == b'\\':
			read_while(f, is_letter)
			read_while(f, is_digit)
			consume_end(f)


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
		word = read_while(f, is_letter).decode()
		if word:
			if word in ESCAPE:
				self.write(ESCAPE[word])
			else:
				param = read_while(f, is_digit)
				param = int(param) if param else None  # can we make the default True?
				if word == 'u':
					# we can always handle unicode
					self.write(chr(param))
					skip_chars(f, self.group.prop.get('uc', 1))
					# don't do consume_end
					return
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
			elif c == b'~':
				self.group.dest.nbsp(self.group, self)
			elif c == b'-':
				self.group.dest.opt_hyphen(self.group, self)
			elif c == b'_':
				self.group.dest.nb_hyphen(self.group, self)
			elif c == b':':
				raise ValueError('subentry')
			else:
				raise ValueError(c)

	def control(self, word, param):
		group = self.group
		if word == 'par':
			group.dest.par(group, self)
		elif word == 'line':
			group.dest.line(group, self)
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
			# use actual font obj?
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
				text = read_while(f, not_control).replace(b'\r\n', b'')
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
	
	@property
	def prop(self):
		return self.group.prop

	def font(self):
		return self.font_table.fonts[self.prop['f']]


rtf = Parser(file)
rtf.parse()


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
