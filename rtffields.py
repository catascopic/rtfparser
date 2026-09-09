from __future__ import annotations

# Field instructions look like `INCLUDEPICTURE "pics\\logo.png" \d \* MERGEFORMAT`, i.e. an instruction name,
# some positional arguments, and some switches, which either stand alone or take the next token as their value.
# That's close enough to a command line that we model this on argparse, minus everything we don't need.
# Note that these tokens have already been through the rtf lexer, so a `\\` in the file arrives as a single
# backslash. Word still doubles the backslashes of paths inside quoted arguments, so we undo that in the tokenizer.

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from types import SimpleNamespace
from typing import NamedTuple

SWITCH = '\\'
QUOTE = '"'


class InstructionError(ValueError):
	pass


class Token(NamedTuple):
	text: str
	quoted: bool

	@property
	def is_switch(self):
		# a quoted token is always a value, even if it starts with a backslash
		return not self.quoted and self.text.startswith(SWITCH)


@dataclass
class Argument:
	name: str
	dest: str
	flag: bool = False
	type: Callable[[str], object] = str
	default: object = None


def tokenize(text: str) -> Iterator[Token]:
	i = 0
	while i < len(text):
		if text[i].isspace():
			i += 1
		elif text[i] == QUOTE:
			i += 1
			buf = []
			while i < len(text) and text[i] != QUOTE:
				# word escapes backslashes (and quotes) inside quoted arguments
				if text[i] == SWITCH and i + 1 < len(text) and text[i + 1] in (SWITCH, QUOTE):
					i += 1
				buf.append(text[i])
				i += 1
			if i == len(text):
				raise InstructionError(f"unterminated quote: {text}")
			i += 1  # closing quote
			yield Token(''.join(buf), True)
		else:
			start = i
			while i < len(text) and not text[i].isspace() and text[i] != QUOTE:
				i += 1
			yield Token(text[start:i], False)


class InstructionParser:

	def __init__(self, name: str):
		self.name = name
		self.switches: dict[str, Argument] = {}
		self.positionals: list[Argument] = []
		# \* is the general formatting switch, which every field type accepts
		self.add_argument(r'\*', dest='format')

	def add_argument(self, 
		name: str, 
		dest: str | None = None, 
		flag: bool = False,
		type: Callable[[str], object] = str, 
		default: object = None
	):
		# a name starting with a backslash is a switch, anything else is a positional argument
		if dest is None:
			dest = name.lstrip(SWITCH)
			if not dest.isidentifier():
				raise ValueError(f"{name} needs an explicit dest")
		arg = Argument(name, dest, flag, type, default)
		if name.startswith(SWITCH):
			self.switches[name] = arg
		elif flag:
			raise ValueError(f"positional argument {name} can't be a flag")
		else:
			self.positionals.append(arg)
		return arg

	def parse(self, text: str) -> SimpleNamespace:
		args = SimpleNamespace(**{
			a.dest: False if a.flag and a.default is None else a.default
			for a in self.arguments()
		})
		positionals = iter(self.positionals)
		tokens = tokenize(text)
		for token in tokens:
			if token.is_switch:
				arg = self.switches.get(token.text)
				if arg is None:
					raise InstructionError(f"unknown switch for {self.name}: {token.text}")
				setattr(args, arg.dest, True if arg.flag else arg.type(self.read_value(arg, tokens)))
			else:
				arg = next(positionals, None)
				if arg is None:
					raise InstructionError(f"too many arguments for {self.name}: {token.text}")
				setattr(args, arg.dest, arg.type(token.text))
		# like argparse, positionals are required, unless they were given a default to fall back on
		for arg in positionals:
			if arg.default is None:
				raise InstructionError(f"missing argument for {self.name}: {arg.name}")
		return args

	def arguments(self) -> Iterator[Argument]:
		yield from self.positionals
		yield from self.switches.values()

	def read_value(self, arg: Argument, tokens: Iterator[Token]) -> str:
		token = next(tokens, None)
		if token is None or token.is_switch:
			raise InstructionError(f"missing value for {self.name} {arg.name}")
		return token.text


INCLUDEPICTURE = InstructionParser('INCLUDEPICTURE')
INCLUDEPICTURE.add_argument('name')
INCLUDEPICTURE.add_argument(r'\c', dest='converter')
INCLUDEPICTURE.add_argument(r'\d', flag=True)
INCLUDEPICTURE.add_argument(r'\x', type=int)
INCLUDEPICTURE.add_argument(r'\y', type=int)

HYPERLINK = InstructionParser('HYPERLINK')
HYPERLINK.add_argument('url')
HYPERLINK.add_argument(r'\l', dest='location')  # a bookmark within the target
HYPERLINK.add_argument(r'\m', dest='image_map', flag=True)
HYPERLINK.add_argument(r'\n', dest='new_window', flag=True)
HYPERLINK.add_argument(r'\o', dest='tip')
HYPERLINK.add_argument(r'\t', dest='target')

PARSERS = {p.name: p for p in (INCLUDEPICTURE, HYPERLINK)}
