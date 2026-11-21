import dataclasses as dcl
import re
import typing as t

from wp2tt.styles import ATTR_READONLY
from wp2tt.styles import ATTR_NO_INI
from wp2tt.styles import OptionalStyle
from wp2tt.styles import Style


def _calculated[T](default: T) -> T:
    return dcl.field(default=default, metadata=ATTR_NO_INI, init=False)


@dcl.dataclass
class Rule:
    """A derivation rule for Styles."""

    mnemonic: str = dcl.field(metadata=ATTR_NO_INI)
    description: str = dcl.field(metadata=ATTR_NO_INI)

    turn_this: str | None = None
    turn_this_style: OptionalStyle = _calculated(None)

    turn_these: str | None = None
    turn_these_patterns: list[re.Pattern[str]] | None = _calculated(None)

    into_this: str | None = None
    into_this_style: OptionalStyle = _calculated(None)

    when_following: str | None = None
    when_following_styles: list[Style] | None = _calculated(None)

    when_first_in_doc: str | None = None
    when_first_in_doc_flag: bool | None = _calculated(None)

    when_matches: str | None = None
    when_matches_patterns: list[re.Pattern[str]] | None = _calculated(None)

    valid: bool = _calculated(True)
    applied: int = _calculated(0)

    _WILDCARD_MARKER = "§"
    _METACHARS = r".^$*+?{}[]\|()"
    _TRUE_VALUES = ("true", "yes", "on", "1")
    _FALSE_VALUES = ("false", "no", "off", "0")
    _TO_BOOL = {
        as_str: as_bool
        for as_bool, as_strs in ((True, _TRUE_VALUES), (False, _FALSE_VALUES))
        for as_str in as_strs
    }

    def __post_init__(self):
        self._parse_flag("when_first_in_doc")
        self._parse_patterns("when_matches")
        self._parse_patterns("turn_these")
        if bool(self.turn_this) == bool(self.turn_these):
            raise ValueError("You must supply EITHER turn_this OR turn_these")

    def _parse_flag(self, field: str) -> None:
        if (text := getattr(self, field)):
            if (value := self._TO_BOOL.get(text)) is None:
                raise ValueError(f"{field} must be one of {' '.join(self._TO_BOOL)}")
            setattr(self, f"{field}_flag", value)

    def _parse_patterns(self, field: str) -> None:
        if not (text := getattr(self, field)):
            return

        patterns = []
        tail = text
        while tail:
            mobj = re.match(r"(?P<marker>\S)(?P<expr>(?:(?!\1).)+)\1", tail)
            if mobj is None:
                raise ValueError(f"Bad {field}: {text!r}")
            patterns.append(self._build_re(mobj.group("marker"), mobj.group("expr")))
            tail = tail[mobj.end()+1:].lstrip()

        setattr(self, f"{field}_patterns", patterns)

    def _build_re(self, marker: str, expr: str) -> re.Pattern[str]:
        if marker == self._WILDCARD_MARKER:
            expr = "".join(self._wildcard_char(c) for c in expr)
            flags = re.IGNORECASE
        else:
            flags = 0
        return re.compile(expr, flags)

    def _wildcard_char(self, char: str) -> str:
        match char:
            case "*":
                return ".*"
            case "?":
                return "."
            case c if c in self._METACHARS:
                return "\\" + c
            case _:
                return char

    def __str__(self) -> str:
        return f"<{self.mnemonic} {self.description!r}>"

    def __repr__(self) -> str:
        parts = [
            self.mnemonic,
            repr(self.description),
            self.turn_this_style or self.turn_these_patterns or repr(self.turn_this) or repr(self.turn_these),
            "->",
            str(self.into_this_style or self.into_this),
        ]
        return f"<{' '.join(map(str, parts))}>"


