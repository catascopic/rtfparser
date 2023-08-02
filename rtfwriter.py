import struct

from rtfparser import Font, Color
from typing import Optional, Iterable, TextIO


ESCAPE = {
	'\n': 'line',
	'\t': 'tab',
	'—': 'emdash',
	'-': 'endash',
	'‘': 'lquote',
	'’': 'rquote',
	'“': 'ldblquote',
	'”': 'rdblquote',
	'•': 'bullet'
}

META_CHARS = frozenset('\\{}')


def write_unicode(c):
	return struct.unpack('hh', c.encode('utf-16-le'))


class RtfWriter:

	# TODO: BinaryIO?
	def __init__(self, handle: TextIO):
		self.handle = handle
		self.separator = False
		self.depth = 0

	def header(self, ansicpg=1252, deff=0, deflang=1033):
		self.control('rtf', 1)
		self.control('ansi')
		self.control('ansicpg', ansicpg)
		self.control('deff', deff)
		self.control('deflang', deflang)

	def viewkind(self, value=4):
		self.control('viewkind', value)

	def uc(self, value=1):
		self.control('uc', value)

	def generator(self, description):
		self.open_group()
		self.special_dest('generator')
		self.text(description)
		self.close_group()

	def special_dest(self, dest):
		self.handle.write('\\*')
		self.control(dest)

	def control(self, word, param: Optional[int] = None):
		self.handle.write('\\')
		self.handle.write(word)
		if param is not None:
			self.handle.write(str(param))
		self.separator = True

	def text(self, text):
		for c in text:
			if c in META_CHARS:
				self.handle.write('\\')
				self.handle.write(c)
			elif ord(c) < 128:
				if self.separator:
					self.handle.write(' ')
					self.separator = False
				self.handle.write(c)
			elif esc := ESCAPE.get(c):
				self.control(esc)
			else:
				try:
					b = c.encode('cp1252')
				except UnicodeEncodeError:
					# we could try other codepages with ANSI, but why bother when we can just use unicode?
					self.handle.write('\\u')
					self.handle.write(str(ord(c)))
					# TODO: if codepoint is above 0xffff...
					# TODO: try to find better replacement char?
					self.handle.write('?')
				else:
					self.handle.write("\\'")
					self.handle.write(f"{b[0]:02x}")

	def par(self):
		self.control('par')
		self.newline()

	def newline(self):
		self.handle.write('\n')
		self.separator = False

	def open_group(self):
		self.depth += 1
		self.handle.write('{')

	def close_group(self):
		self.depth -= 1
		if self.depth < 0:
			raise ValueError()
		self.handle.write('}')

	def finish(self):
		if self.depth != 0:
			raise ValueError(f"depth = {self.depth}")
		f.write('\n\u0000')

	def font_table(self, fonts: dict[int, Font] | Iterable[Font]):
		if isinstance(fonts, dict):
			fonts = fonts.items()
		else:
			fonts = enumerate(fonts)

		self.open_group()
		self.control('fonttbl')
		for i, font in fonts:
			self.open_group()
			self.control('f', i)
			self.control('f' + font.family)
			if font.charset is not None:
				self.control('fcharset', font.charset)
			self.text(font.name)
			self.text(';')
			self.close_group()
		self.close_group()

	def color_table(self, colors: Iterable[Color]):
		self.open_group()
		self.control('colortbl')
		for color in colors:
			self.control('red', color.red)
			self.control('green', color.green)
			self.control('blue', color.blue)
			self.handle.write(';')
		self.close_group()


with open('generated.rtf', 'w', encoding='ascii') as f:
	w = RtfWriter(f)
	w.open_group()
	w.header()
	w.font_table([Font('Arial', 'nil')])
	w.newline()
	w.generator('micro-rtf 0.1')
	w.viewkind()
	w.uc(1)
	w.newline()
	w.control('fs', 20)
	w.text(chr(0x10001) + 'This is a test—β——Vigenère cipher😀😁')
	w.par()
	w.close_group()
	w.finish()

for encoding in ['ansi', 'cp1252', 'cp1253']:
	try:
		print(encoding, bytes([0xe2]).decode(encoding))
	except UnicodeEncodeError:
		print(encoding, 'FAILED')


print(struct.unpack('I', struct.pack('h', -10179) + struct.pack('h', -8704)))
print(chr(0x10001))
print(hex((1<<16) + -8191))
print(b'\xe2'.decode('johab'))

