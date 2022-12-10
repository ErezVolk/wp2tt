"""Math conversion"""
from pathlib import Path

from lxml import etree
import ziamath.config
import ziamath.zmath
ziamath.config.svg2 = False  # Otherwise InDesign chokes


class MathConverter:
    """Convert Office Math Markup Language -> MathML -> SVG"""
    MS_XSLT = Path("/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl")
    transform: etree.XSLT | None = None

    @classmethod
    def omml_to_mathml(cls, omml: etree._Entity) -> etree._ElementTree:
        """Convert Office Math Markup Language to MathML"""
        if cls.transform is None:
            cls.transform = cls._load_xslt()

        return cls.transform.apply(omml)

    @classmethod
    def mathml_to_svg(cls, mathml: etree._ElementTree, size: int | None) -> bytes:
        """Convert MathML to SVG"""
        converted = ziamath.zmath.Math(mathml, size=size or 12)
        return converted.svg().encode("utf-8")

    @classmethod
    def _load_xslt(cls, try_ms=True) -> etree.XSLT:
        """Read transformer"""
        if try_ms and cls.MS_XSLT.is_file():
            # Use latest version from install MS Word
            path = cls.MS_XSLT
        else:
            # Use older version from https://raw.githubusercontent.com/TEIC/Stylesheets
            path = Path(__file__).parent / "omml2mml.xsl"

        return etree.XSLT(etree.parse(path))
