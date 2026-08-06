"""
Minimal helper for resolving RTF \\fcharsetN codes (and \\anscipg / \\ansi / \\pc / \\pca / \\mac
document-level charsets) to Python codec names.

rtfparser.py imports this module but does not ship it, so this is a small
best-effort implementation covering the common RTF character sets. It's used
only to decode \\'xx hex-escaped bytes.
"""

# Maps the numeric \fcharsetN value (set per-font in the font table) to a Python codec name.
# None means "no specific mapping -- fall back to the document/section charset".
_FCHARSET_TO_CODEC = {
    0:   'cp1252',    # ANSI
    1:   None,        # Default -- defer to document charset
    2:   'cp1252',    # Symbol (no real text codec; ANSI is the least-bad fallback)
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

# Maps the document-level charset name (as produced by rtfparser's CHARSETS dict, or
# an "cp<NNN>" string from \ansicpg) to a Python codec name.
_DOC_CHARSET_TO_CODEC = {
    'ansi': 'cp1252',
    'mac': 'mac_roman',
}

DEFAULT_CODEC = 'cp1252'


def get_encoding(font_charset, doc_charset):
    """
    font_charset: the value of \\fcharsetN for the current font (an int), or None.
    doc_charset: the current document/section charset, e.g. 'ansi', 'mac', 'cp437', 'cp1252'.
    """
    if font_charset is not None:
        codec = _FCHARSET_TO_CODEC.get(font_charset)
        if codec:
            return codec

    if doc_charset:
        return _DOC_CHARSET_TO_CODEC.get(doc_charset, doc_charset)

    return DEFAULT_CODEC