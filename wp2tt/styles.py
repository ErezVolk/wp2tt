#!/usr/bin/env python3
"""wp2tt Style objects"""
from typing import Union

import attr
from wp2tt.ini import ATTR_KEY
from wp2tt.ini import ATTR_VALUE_READONLY
from wp2tt.ini import ATTR_VALUE_HIDDEN
from wp2tt.format import ManualFormat

ATTR_READONLY = {ATTR_KEY: ATTR_VALUE_READONLY}
ATTR_NO_INI = {ATTR_KEY: ATTR_VALUE_HIDDEN}

OptionalStyle = Union["Style", None]


@attr.s(slots=True)
class Style:
    """A character/paragraph style, normally found in the input file."""

    realm: str = attr.ib(metadata=ATTR_NO_INI)
    wpid: str = attr.ib(metadata=ATTR_READONLY)  # Used for xrefs in docx; localized
    internal_name: str = attr.ib(metadata=ATTR_NO_INI)  # Used in the section names
    name: str = attr.ib()  # What the user (and InDesign) see
    parent_wpid: str = attr.ib(default=None, metadata=ATTR_READONLY)
    next_wpid: str = attr.ib(default=None, metadata=ATTR_READONLY)
    automatic: bool = attr.ib(default=None, metadata=ATTR_READONLY)
    custom: bool = attr.ib(default=False, metadata=ATTR_READONLY)
    fmt: ManualFormat = attr.ib(default=ManualFormat.NORMAL, metadata=ATTR_READONLY)
    idtt: str = attr.ib(default="")
    variable: str = attr.ib(default=None)

    used: bool = attr.ib(default=False, metadata=ATTR_NO_INI, eq=False)
    count: int = attr.ib(default=0, metadata=ATTR_NO_INI, eq=False)

    parent_style: OptionalStyle = attr.ib(default=None, metadata=ATTR_NO_INI, eq=False)
    next_style: OptionalStyle = attr.ib(default=None, metadata=ATTR_NO_INI, eq=False)

    def __str__(self):
        if self.custom:
            return f"<{self.realm} {repr(self.name)}"
        return f"<{self.realm} {repr(self.name)} (built-in)>"


@attr.s(slots=True)
class Rule:
    """A derivation rule for Styles."""

    mnemonic: str = attr.ib(metadata=ATTR_NO_INI)
    description: str = attr.ib(metadata=ATTR_NO_INI)
    turn_this: str = attr.ib(default=None)
    into_this: str = attr.ib(default=None)
    when_following: str = attr.ib(default=None)
    when_first_in_doc: str = attr.ib(default=None)

    turn_this_style: OptionalStyle = attr.ib(default=None, metadata=ATTR_NO_INI)
    into_this_style: OptionalStyle = attr.ib(default=None, metadata=ATTR_NO_INI)
    when_following_styles: list[Style] = attr.ib(default=None, metadata=ATTR_NO_INI)

    valid: bool = attr.ib(default=True, metadata=ATTR_NO_INI)
    applied: int = attr.ib(default=0, metadata=ATTR_NO_INI)

    def __str__(self):
        return f"<{self.mnemonic} {repr(self.description)}>"


@attr.s(slots=True)
class DocumentProperties:
    """Things we can tell about a document."""

    has_rtl: bool = attr.ib(default=True)
    pure_ascii: bool = attr.ib(default=False)
