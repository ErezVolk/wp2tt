"""Math conversion."""
# InDesign now supports MathML; we should move to that.

from pathlib import Path

from lxml import etree
import ziamath.config
import ziamath.zmath

# This is required for InDesign compatibility
ziamath.config.svg2 = False  # type: ignore[ty:unresolved-attribute]


class MathConverter:
    """Convert Office Math Markup Language -> MathML -> SVG."""

    MS_XSLT = Path(
        "/Applications/Microsoft Word.app/Contents"
        "/Resources/omml2mathml.xsl",
    )
    transform: etree.XSLT | None = None

    @classmethod
    def omml_to_mathml(cls, omml: etree._Element) -> etree._ElementTree:
        """Convert Office Math Markup Language to MathML."""
        if cls.transform is None:
            cls.transform = cls._load_xslt()
        assert cls.transform is not None
        return cls.transform(omml)

    @classmethod
    def mathml_to_svg(cls, mathml: str, size: int | None) -> bytes:
        """Convert MathML to SVG."""
        converted = ziamath.zmath.Math(mathml, size=size or 12)
        return converted.svg().encode("utf-8")

    @classmethod
    def _load_xslt(cls, *, try_ms: bool = True) -> etree.XSLT:
        """Read transformer."""
        if try_ms and cls.MS_XSLT.is_file():
            # Use latest version from install MS Word
            path = cls.MS_XSLT
        else:
            # Use older version from https://raw.githubusercontent.com/TEIC/Stylesheets
            path = Path(__file__).parent / "omml2mml.xsl"

        return etree.XSLT(etree.parse(path))
