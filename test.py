from collections import deque

from rtfparser import *

# card, ord, and ordt not supported
HTML_LIST_TYPES = {'pndec': '1',  'pnucltr': 'A', 'pnucrm': 'I', 'pnlcltr': 'a', 'pnlcrm': 'i', }

def diff_prop(old, new):
	diffs = []
	for k in old.keys() & new.keys():
		oldval = old[k]
		newval = new[k]
		if oldval != newval:
			diffs.append(f"{k}: {oldval}->{newval}")
	for k in old.keys() - new.keys():
		if k in TOGGLE:
			diffs.append(f"+{k}")
		else:
			diffs.append(f"{k}: -{old[k]}")
	for k in new.keys() - old.keys():
		if k in TOGGLE:
			diffs.append(f"-{k}")
		else:
			diffs.append(f"{k}: +{new[k]}")
	if not diffs:
		return ''
	return f"<{'; '.join(diffs)}>"


HTML_TRANS = str.maketrans({'&': '&amp;', '<': '&lt;', '>': '&gt;', '\n': '\n<br>'})

with open('result.html', 'w', encoding='utf-8') as writer:

	class Recorder(Handler):

		def __init__(self, doc):
			super().__init__(doc)
			self.paragraphs = []
			self.current = []
			self.last_prop = {}
			self.styles = set()
			self.par_num = 1
			self.stack = deque()
			self.blockquote = False

		def write(self, text):
			# TODO: the dilemma: handle prop changes as events, or continue doing prop diffs
			# print(diff_prop(self.prop, self.last_prop), end='')
			prop_on = (self.prop.keys() - self.last_prop.keys()) & TOGGLE
			prop_off = (self.last_prop.keys() - self.prop.keys()) & TOGGLE
			if 'i' in prop_on:
				self.current.append('<i>')
				self.stack.append('i')
			if 'i' in prop_off:
				if self.stack.pop() != 'i':
					raise ValueError
				if self.current[-1] == '<i>':
					self.current.pop()
				else:
					self.current.append('</i>')

			self.last_prop = self.prop.copy()
			self.current.append(text.translate(HTML_TRANS))

		def par(self):
			for i in self.stack:
				if self.current[-1] == f"<{i}>":
					self.current.pop()
				else:
					self.current.append(f"</{i}>")

			line = ''.join(self.current)
			if self._doc.numbering:
				tag = 'li'
			else:
				tag = 'p'

			css = []
			if self.prop.get('q') == 'c':
				css.append('text-align: center;')
			style = f" style=\"{' '.join(css)}\"" if css else ''

			endtag = f"</{tag}>" if tag != 'p' else ''

			if self.prop.get('li', 0) > 0 and not self.blockquote:
				self.blockquote = True
				writer.write('\n<blockquote>')
			elif self.prop.get('li', 0) == 0 and self.blockquote:
				self.blockquote = False
				writer.write('\n</blockquote>')
			writer.write(f"\n<{tag}{style}>{line}{endtag}")

			self.paragraphs.append(line)
			self.current = [f"<{i}>" for i in self.stack]
			self.par_num += 1

		def numbering_on(self, info: Numbering):
			writer.write(f'<ol type="{HTML_LIST_TYPES[info.style]}">')

		def numbering_off(self, info):
			writer.write('</ol>')

		def hyperlink(self, text, args):
			# TODO: condense hyperlinks? They can be spread out across multiple words sometimes.
			# This is a more general problem that can apply to any kind of styling, though.
			self.current.append(f'<a href="{args.url}">{text}</a>')

		def end_doc(self):
			pass

	rtf = Parser(Recorder)
	rtf.parse('imagetest.rtf')
