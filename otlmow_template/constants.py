"""The constants that the parts of the template generation share."""

from pathlib import Path

ROOT_DIR = (Path(__file__).parent.parent).resolve()

SHEET_TITLE_HAS_NO_HASH = 'Sheet title does not contain a #'
HEADER_EXISTS_AT_COLUMN = 'Header already exists at column {column_nr}: {value}'

short_to_long_ns = {
    'ond': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'onderdeel': 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#',
    'ins': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'installatie': 'https://wegenenverkeer.data.vlaanderen.be/ns/installatie#',
    'imp': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
    'implementatieelement': 'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#',
    'proefenmeting': 'https://wegenenverkeer.data.vlaanderen.be/ns/proefenmeting#',
    'pro': 'https://wegenenverkeer.data.vlaanderen.be/ns/proefenmeting#',
    'lev': 'https://wegenenverkeer.data.vlaanderen.be/ns/levenscyclus#',
    'levenscyclus': 'https://wegenenverkeer.data.vlaanderen.be/ns/levenscyclus#',
}

AGENT_TYPE_URI = 'http://purl.org/dc/terms/Agent'
HEEFT_BETROKKENE_URI = 'https://wegenenverkeer.data.vlaanderen.be/ns/onderdeel#HeeftBetrokkene'
TYPE_URI_DESCRIPTION = 'De URI van het object volgens https://www.w3.org/2001/XMLSchema#anyURI .'
QUOTE_CHAR = '"'
# the validations are only added up to this row, so a template stays usable
MAX_VALIDATION_ROW = 1000
