from pathlib import Path
from dataclasses import dataclass
from abc import ABC


file = Path('The Inner Searchlight.rtf')
# file = Path('test.rtf')


class Group:

	def __init__(self, parent):
		self.parent = parent
		if parent is None:
			self.dest = Root()
	
	def control(self, word, param):
		if word in ESCAPE:
			self.dest.write(ESCAPE[word])
		elif word == "'":
			self.dest.write(param)
		elif word == 'par':
			self.dest.write('\n')
		elif word == 'pard':
			pass
		elif word in FONT_FAMILIES:
			self.family = word[1:]
		elif word == 'rtf':
			self.dest = output
		elif word == 'fonttbl':
			self.dest = font_table
		elif word == 'colortbl':
			self.dest = color_table
		elif word in {'filetbl', 'stylesheet', 'listtables', 'revtbl', '*'}:
			self.dest = NullDevice()
		else:
			setattr(self, word, param)

	# jesus christ this might not be worth it
	def __getattr__(self, name):
		if self.parent is None:
			return None
		return getattr(self.parent, name)


class Destination(ABC):

	def write(self, text):
		return NotImplemented
	
	def control(self, word, param):
		return NotImplemented
			

class Output(Destination):

	def __init__(self):
		self.full_text = []

	def write(self, text):
		self.full_text.append(text)
		# print(text, end='')
		

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

	def write(self, text):
		# should this be more rigorous?
		name = text.removesuffix(';')
		self.fonts[group.f] = Font(name, group.family, group.charset)


@dataclass		
class Color:
	red: int
	green: int
	blue: int


class ColorTable(Destination):
	
	def __init__(self):
		self.colors = []

	def write(self, text):
		if text == ';':
			self.colors.append(Color(group.red or 0, group.green or 0, group.blue or 0))


class Root(Destination):
	def write(self, text):
		if text != '\x00':
			raise ValueError(text.encode())


class NullDevice(Destination):
	def write(self, text):
		pass  # do nothing


ESCAPE = {
	'tab': '\t',
	'emdash': '—',
	'endash': '-',
	'lquote': '‘',
	'rquote': '’',
	'ldblquote': '“',
	'rdblquote': '”',
}


def is_letter(c):
	return b'a' <= c <= b'z' or b'A' <= c <= b'Z'

def is_digit(c):
	return b'0' <= c <= b'9' or c == '-'

def not_control(c):
	return c not in {b'\\', b'{', b'}'}


def read_until(f, matcher):
	start = f.tell()
	while True:
		c = f.read(1)
		if not matcher(c) or not c:
			end = f.tell()
			f.seek(start)
			return f.read(end - start - len(c))
	

def read_control(f):
	word = read_until(f, is_letter).decode()
	if not word:
		c = f.read(1)
		if c == b'*':  # macro
			return '*', None
		if c == b"'":
			return "'", chr(int(f.read(2), 16))
		raise ValueError(c)

	param = read_until(f, is_digit)
	c = f.read(1)
	if c == b'\r':
		f.read(1)
		# check that this next char is \n?
	elif not c.isspace():
		f.seek(-1, 1)
	return word, int(param) if param else None


output = Output()
font_table = FontTable()
color_table = ColorTable()
group = Group(None)

with open(file, 'rb') as f:
	while True:
		text = read_until(f, not_control).replace(b'\r\n', b'')
		if text:
			group.dest.write(text.decode())
		c = f.read(1)
		if c == b'{':
			group = Group(group)
		elif c == b'}':
			group = group.parent
		elif c == b'\\':
			group.control(*read_control(f))
		elif not c:
			break


print(font_table.fonts)
print(color_table.colors)
print(len(''.join(output.full_text)))
