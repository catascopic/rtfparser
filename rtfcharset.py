from typing import Optional


CHARSETS = {
	0: 'ansi',
	1: None,  # use default
	2: 'ansi',  # "Symbol"; as far as I can tell this is just ansi
	# 3 is "invalid"
	77: 'macintosh',
	# 78 - 89 are more macintosh charsets that no one should be using
	128: 'cp932',   # Shift JIS
	129: 'cp949',   # Hangul (1362?)
	130: 'johab',
	134: 'gb2312',
	136: 'big5',
	161: 'cp1253',  # Greek ('greek' also works)
	162: 'cp1254',  # Turkish
	163: 'cp1258',  # Vietnamese
	177: 'cp1255',  # Hebrew
	178: 'cp1256',  # Arabic
	# 179 - 181 are obsolete
	186: 'cp1257',  # Baltic
	204: 'cp1251',  # Russian
	222: 'cp874',   # Thai
	238: 'cp1250',  # Eastern European
	254: 'cp437',   # original IBM codepage, god forbid
	255: 'oem',
}


def get_encoding(charset: Optional[int], default: str) -> str:
	if charset is None or charset == 1:
		return default
	return CHARSETS[charset]
