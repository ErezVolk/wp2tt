import dataclasses as dcl
import re

from wp2tt.styles import ATTR_READONLY
from wp2tt.styles import ATTR_NO_INI
from wp2tt.styles import OptionalStyle
from wp2tt.styles import Style


@dcl.dataclass
class Rule:
    """A derivation rule for Styles."""

    mnemonic: str = dcl.field(metadata=ATTR_NO_INI)
    description: str = dcl.field(metadata=ATTR_NO_INI)
    turn_this: str | None = None
    into_this: str | None = None
    when_following: str | None = None
    when_first_in_doc: str | None = None
    when_matches: str | None = None
    unless_empty: str | None = None

    turn_this_style: OptionalStyle = dcl.field(
        default=None, metadata=ATTR_NO_INI, init=False,
    )
    into_this_style: OptionalStyle = dcl.field(
        default=None, metadata=ATTR_NO_INI, init=False,
    )
    when_following_styles: list[Style] | None = dcl.field(
        default=None, metadata=ATTR_NO_INI, init=False,
    )
    when_matches_patterns: list[re.Pattern] | None = dcl.field(
        default=None, metadata=ATTR_NO_INI, init=False,
    )

    valid: bool = dcl.field(default=True, metadata=ATTR_NO_INI, init=False)
    applied: int = dcl.field(default=0, metadata=ATTR_NO_INI, init=False)

    _TRUE_VALUES = ("true", "yes", "on", "1")
    _FALSE_VALUES = ("false", "no", "off", "0")
    _TO_BOOL = {
        as_str: as_bool
        for as_bool, as_strs in ((True, _TRUE_VALUES), (False, _FALSE_VALUES))
        for as_str in as_strs
    }

    @classmethod
    def to_bool(cls, as_str: str | None) -> None | bool:
        if as_str is None:
            return None
        return cls._TO_BOOL[as_str.lower()]

    def __post_init__(self):
        for key in ("when_first_in_doc", "unless_empty"):
            if (value := getattr(self, key)) is not None:
                if value.lower() not in self._TO_BOOL:
                    raise ValueError(f"{key} may only be True or False, not {value!r}")

    def __str__(self) -> str:
        return f"<{self.mnemonic} {self.description!r}>"


