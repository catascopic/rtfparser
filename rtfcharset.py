CHARSETS = {
	0:   'cp1252',     # ANSI
	1:   None,         # Default; defer to document charset
	2:   'cp1252',     # Symbol (no real text codec; ANSI is the least-bad fallback)
	77:  'mac_roman',  # Mac
	78:  'shift_jis',  # Mac Shift-JIS
	128: 'shift_jis',  # Shift-JIS
	129: 'cp949',      # Hangul
	130: 'johab',      # Johab
	134: 'gb2312',     # GB2312 (Simplified Chinese)
	136: 'big5',       # Big5 (Traditional Chinese)
	161: 'cp1253',     # Greek
	162: 'cp1254',     # Turkish
	163: 'cp1258',     # Vietnamese
	177: 'cp1255',     # Hebrew
	178: 'cp1256',     # Arabic
	186: 'cp1257',     # Baltic
	204: 'cp1251',     # Russian / Cyrillic
	222: 'cp874',      # Thai
	238: 'cp1250',     # Eastern European
	254: 'cp437',      # PC 437
	255: 'cp1252',     # OEM
}


def get_encoding(charset: int | None, default: str) -> str:
	if charset is None or charset == 1:
		return default
	return CHARSETS[charset]
