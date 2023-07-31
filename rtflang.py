from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Language:
	description: str
	tag: Optional[str]
	locale: str = None


NO_LANGUAGE = Language('No language', None)


LANGUAGES = {
	1: Language('Arabic', 'ar'),
	4: Language('Chinese', 'zh'),
	9: Language('English', 'en'),
	1024: NO_LANGUAGE,
	1025: Language('Arabic', 'ar', 'SA'),
	1026: Language('Bulgarian', 'bg', 'BG'),
	1027: Language('Catalan', 'ca', 'ES'),
	1028: Language('Traditional Chinese', 'zh', 'TW'),
	1029: Language('Czech', 'cs', 'CZ'),
	1030: Language('Danish', 'da', 'DK'),
	1031: Language('German', 'de', 'DE'),
	1032: Language('Greek', 'el', 'GR'),
	1033: Language('English (U.S.)', 'en', 'US'),
	1034: Language('Spanish (Castilian)', 'es', 'ES'),
	1035: Language('Finnish', 'fi', 'FI'),
	1036: Language('French', 'fr', 'FR'),
	1037: Language('Hebrew', 'he', 'IL'),
	1038: Language('Hungarian', 'hu', 'HU'),
	1039: Language('Icelandic', 'is', 'IS'),
	1040: Language('Italian', 'it', 'IT'),
	1041: Language('Japanese', 'ja', 'JP'),
	1042: Language('Korean', 'ko', 'KR'),
	1043: Language('Dutch', 'nl', 'NL'),
	1044: Language('Norwegian (Bokmal)', 'nb', 'NO'),
	1045: Language('Polish', 'pl', 'PL'),
	1046: Language('Brazilian Portuguese', 'pt', 'BR'),
	1047: Language('Rhaeto-Romanic', 'rm', 'CH'),
	1048: Language('Romanian', 'ro', 'RO'),
	1049: Language('Russian', 'ru', 'RU'),
	1050: Language('Croato-Serbian (Latin)', 'hr', 'HR'),
	1051: Language('Slovak', 'sk', 'SK'),
	1052: Language('Albanian', 'sq', 'AL'),
	1053: Language('Swedish', 'sv', 'SE'),
	1054: Language('Thai', 'th', 'TH'),
	1055: Language('Turkish', 'tr', 'TR'),
	1056: Language('Urdu', 'ur', 'PK'),
	1057: Language('Bahasa', 'id', 'ID'),
	1065: Language('Farsi (Persian)', 'fa', 'IR'),
	1072: Language('Sesotho (Sotho)', 'st', 'ZA'),
	1073: Language('Tsonga', 'ts', 'ZA'),
	1074: Language('Tswana', 'tn', 'ZA'),
	1075: Language('Venda', 'ven', 'ZA'),
	1076: Language('Xhosa', 'xh', 'ZA'),
	1077: Language('Zulu', 'zu', 'ZA'),
	1078: Language('Afrikaans', 'af', 'ZA'),
	2052: Language('Simplified Chinese', 'zh', 'CN'),
	2055: Language('Swiss German', 'de', 'CH'),
	2057: Language('English (U.K.)', 'en', 'GB'),
	2058: Language('Spanish (Mexican)', 'es', 'MX'),
	2060: Language('Belgian French', 'fr', 'BE'),
	2064: Language('Swiss Italian', 'it', 'CH'),
	2067: Language('Belgian Dutch', 'nl', 'BE'),
	2068: Language('Norwegian (Nynorsk)', 'nn', 'NO'),
	2070: Language('Portuguese', 'pt', 'PT'),
	2074: Language('Serbo-Croatian (Cyrillic)', 'sr', 'CS'),
	3081: Language('English (Australian)', 'en', 'AU'),
	3084: Language('French (Canadian)', 'fr', 'CA'),
	4108: Language('Swiss French', 'fr', 'CH'),
}
