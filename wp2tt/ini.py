"""Ini file helper"""
import configparser

from typing import Dict
from typing import Generator
from typing import Tuple
from typing import Union

import attr

ATTR_KEY = "special"
ATTR_VALUE_READONLY = "readonly"
ATTR_VALUE_HIDDEN = "internal"


ConfigSection = Union[configparser.SectionProxy, Dict[str, str]]


def ini_fields(klass, writeable=False) -> Generator[Tuple[str, str], None, None]:
    """Yields a pair (name, ini_name) for all attributes."""
    for field in attr.fields(klass):
        special = field.metadata.get(ATTR_KEY)
        if special == ATTR_VALUE_HIDDEN:
            continue
        ini_name = name = field.name
        if special == ATTR_VALUE_READONLY:
            if writeable:
                continue
            ini_name += " (readonly)"
        yield (name, ini_name)
