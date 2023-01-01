"""Ini file helper"""
import configparser

from typing import Iterable

import attr

ATTR_KEY = "special"
ATTR_VALUE_READONLY = "readonly"
ATTR_VALUE_HIDDEN = "internal"


ConfigSection = (configparser.SectionProxy | dict[str, str])


def ini_fields(klass, writeable=False) -> Iterable[tuple[str, str]]:
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
