"""Manual Formatting."""
import dataclasses as dcl
import typing as t


type DIRECTION = t.Literal["RTL", "LTR"] | None
type ALIGNMENT = t.Literal["CENTERED", "JUSTIFIED", "LEFT", "RIGHT"] | None


@dcl.dataclass(frozen=True)
class ManualFormat:
    NORMAL: t.ClassVar[ManualFormat]

    align: ALIGNMENT | None = None
    direction: DIRECTION = None
    new_page: t.Literal[True] | None = None
    spaced: t.Literal[True] | None = None
    stars: t.Literal[True] | None = None
    post_stars: t.Literal[True] | None = None
    indented: t.Literal[True] | None = None
    bold: t.Literal[True] | None = None
    italic: t.Literal[True] | None = None
    highlight: t.Literal[True] | None = None
    subsup: t.Literal["SUBSCRIPT", "SUPERSCRIPT"] | None = None
    baseline: t.Literal["RAISED", "LOWERED"] | None = None
    size: int | None = None
    indent: t.Literal["FIRSTLINE", "HANGING"] | None = None

    def replace(self, **kwargs) -> ManualFormat:
        return dcl.replace(self, **kwargs)

    def _active(self) -> t.Generator[tuple[str, t.Literal[True] | str | float]]:
        for name, value in dcl.asdict(self).items():
            if value is not None:
                yield (name, value)

    def add(self, other: ManualFormat) -> ManualFormat:
        return self.replace(**dict(other._active()))

    def sub(self, other: ManualFormat) -> ManualFormat:
        return self.replace(**{
            name: None
            for name, _ in other._active()
        })

    def __str__(self) -> str:
        parts = []
        for name, value in self._active():
            if value is True:
                parts.append(name.upper())
            elif isinstance(value, str):
                parts.append(value)
            elif isinstance(value, (int)):
                parts.append(f"{name.upper()}_{value}")
        if not parts:
            return "DEFAULT"
        return ".".join(sorted(parts))


ManualFormat.NORMAL = ManualFormat()

