"""wp2tt Style objects."""
import dataclasses as dcl
from typing import Union

from wp2tt.format import ManualFormat

ATTR_KEY = "special"
ATTR_VALUE_READONLY = "readonly"
ATTR_VALUE_HIDDEN = "internal"

ATTR_READONLY = {ATTR_KEY: ATTR_VALUE_READONLY}
ATTR_NO_INI = {ATTR_KEY: ATTR_VALUE_HIDDEN}

OptionalStyle = Union["Style", None]


@dcl.dataclass
class Style:
    """A character/paragraph style, normally found in the input file."""

    realm: str = dcl.field(metadata=ATTR_NO_INI)
    wpid: str = dcl.field(metadata=ATTR_READONLY)  # Used for xrefs in docx; localized
    internal_name: str = dcl.field(metadata=ATTR_NO_INI)  # Used in the section names
    name: str = dcl.field()  # What the user (and InDesign) see
    parent_wpid: str | None = dcl.field(default=None, metadata=ATTR_READONLY)
    next_wpid: str | None = dcl.field(default=None, metadata=ATTR_READONLY)
    automatic: bool = dcl.field(default=False, metadata=ATTR_READONLY)
    custom: bool = dcl.field(default=False, metadata=ATTR_READONLY)
    fmt: ManualFormat = dcl.field(default=ManualFormat.NORMAL, metadata=ATTR_READONLY)
    idtt: str = ""
    variable: str | None = None

    used: bool = dcl.field(default=False, metadata=ATTR_NO_INI, compare=False)
    count: int = dcl.field(default=0, metadata=ATTR_NO_INI, compare=False)

    parent_style: OptionalStyle = dcl.field(
        default=None, metadata=ATTR_NO_INI, compare=False,
    )
    next_style: OptionalStyle = dcl.field(
        default=None, metadata=ATTR_NO_INI, compare=False,
    )

    def __str__(self) -> str:
        if self.custom:
            return f"<{self.realm} {self.name!r}"
        return f"<{self.realm} {self.name!r} (built-in)>"

    def __hash__(self) -> int:
        return hash(f"{self.realm}:{self.name}")


@dcl.dataclass
class Rule:
    """A derivation rule for Styles."""

    mnemonic: str = dcl.field(metadata=ATTR_NO_INI)
    description: str = dcl.field(metadata=ATTR_NO_INI)
    turn_this: str | None = None
    into_this: str | None = None
    when_following: str | None = None
    when_first_in_doc: str | None = None

    turn_this_style: OptionalStyle = dcl.field(default=None, metadata=ATTR_NO_INI)
    into_this_style: OptionalStyle = dcl.field(default=None, metadata=ATTR_NO_INI)
    when_following_styles: list[Style] | None = dcl.field(
        default=None, metadata=ATTR_NO_INI,
    )

    valid: bool = dcl.field(default=True, metadata=ATTR_NO_INI)
    applied: int = dcl.field(default=0, metadata=ATTR_NO_INI)

    def __str__(self) -> str:
        return f"<{self.mnemonic} {self.description!r}>"


@dcl.dataclass
class DocumentProperties:
    """Things we can tell about a document."""

    has_rtl: bool = True
    pure_ascii: bool = False
