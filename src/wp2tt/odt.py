""".odt file parsing."""
import contextlib
import logging
from pathlib import Path
import typing as t

from lxml import etree

from wp2tt.input import IDocComment
from wp2tt.input import IDocFootnote
from wp2tt.input import IDocInput
from wp2tt.input import IDocParagraph
from wp2tt.input import IDocSpan
from wp2tt.styles import DocumentProperties
from wp2tt.zip import ZipDocument

log = logging.getLogger(__name__)


class OoXml:
    """Basic helper class for the OpenOffice XML format."""

    _NS: t.Mapping[str, str] = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }

    _FAMILY_TO_REALM: t.ClassVar[dict[str, str]] = {
        "paragraph": "paragraph",
        "text": "character",
    }

    def _xpath(self, node: etree._Element, expr: str) -> t.Iterable[etree._Element]:
        return node.xpath(expr, namespaces=self._NS)

    def _ootag(self, tag: str) -> str:
        namespace, tag = tag.split(":", 1)
        prefix = self._NS[namespace]
        return f"{{{prefix}}}{tag}"

    def _ooget(self, node: etree._Element, tag: str) -> str | None:
        return node.get(self._ootag(tag))


class XodtInput(contextlib.ExitStack, OoXml, IDocInput):
    """A reader for .odt and .fodt."""

    def __init__(self, path: Path, *, zipped: bool) -> None:
        super().__init__()
        self._zipped = zipped
        if zipped:
            self._zip = self._open_zip(path)
        else:
            self._flat = self._open_flat(path)
        self._content = self._load_xml("content.xml")
        self._initialize_properties()

    def _initialize_properties(self) -> None:
        self._properties = DocumentProperties(
            has_rtl=self._has_node(
                '//style:paragraph-properties[@style:writing-mode="rl-tb"]',
            ),
        )

    def _has_node(self, ootag: str) -> bool:
        for _ in self._xpath(self._content, ootag):
            return True
        return False

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return self._properties

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        styles = self._load_xml("styles.xml")
        for node in self._xpath(styles, "//office:styles/style:style"):
            yield self._style_kwargs(node)
        for node in self._xpath(self._content, "//office:automatic-styles/style:style"):
            yield self._style_kwargs(node, automatic=True)

    def _style_kwargs(self, node: etree._Element, **extras) -> dict:
        name = self._ooget(node, "style:name")
        family = self._ooget(node, "style:family")
        if not (name and family):
            raise RuntimeError("Bad family/name for style")
        style_kwargs = dict(extras) if extras else {}
        style_kwargs.update(
            {
                "realm": self._FAMILY_TO_REALM[family],
                "internal_name": self._ooget(node, "style:display-name") or name,
                "wpid": name,
                "parent_wpid": self._ooget(node, "style:parent-style-name"),
                "next_wpid": self._ooget(node, "style:next-style-name"),
            },
        )
        return style_kwargs

    def styles_in_use(self) -> t.Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        for realm, tag in (("paragraph", "text:p"), ("character", "text:span")):
            for sname in self._xpath(self._content, f"//{tag}"):
                wpid = self._ooget(sname, "text:style-name")
                yield (realm, wpid)

    def paragraphs(self) -> t.Iterable[IDocParagraph]:
        """Yield a OdtParagraph object for each body paragraph."""
        for para in self._xpath(self._content, "//office:body/office:text/text:p"):
            yield OdtParagraph(self, para)

    def _open_zip(self, path: Path) -> ZipDocument:
        return self.enter_context(ZipDocument(path))

    @classmethod
    def _open_flat(cls, path: Path) -> etree._Element:
        with path.open(encoding="utf8") as fobj:
            return etree.parse(fobj).getroot()

    def _load_xml(self, path_in_zip: str) -> etree._Element:
        """Parse an XML file inside the zipped doc, return root node."""
        if not self._zipped:
            return self._flat

        return self._zip.load_xml(path_in_zip)


class OdtNode(OoXml):
    """Base helper class for object which represent a node in a docx."""

    def __init__(self, node: etree._Element) -> None:
        self.node = node

    def _node_ooget(self, tag: str) -> str | None:
        return self._ooget(self.node, tag)

    def _node_xpath(self, expr: str) -> t.Iterable[etree._Element]:
        return self.node.xpath(expr, namespaces=self._NS)


class OdtParagraph(OdtNode, IDocParagraph):
    """A Paragraph inside a .docx."""

    def style_wpid(self) -> str | None:
        """Return the wpid for this paragraph's style."""
        return self._node_ooget("text:style-name")

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        yield from self.node.itertext()

    def chunks(self) -> t.Iterable[IDocParagraph.Chunk]:
        """Yield OdtSpan per text span."""
        for event, node in etree.iterwalk(self.node, events=("start", "end")):
            if event == "start":
                if node.tag == self._ootag("text:tab"):
                    yield OdtTabSpan(node)
                elif node.tag == self._ootag("text:span"):
                    yield OdtSpanSpan(node)
                elif node.tag == self._ootag("text:p"):
                    yield OdtHeadSpan(node)
                else:
                    log.debug(
                        "Not sure what to do with a <%s> %r", node.tag, node.text[:8],
                    )
                    yield OdtHeadSpan(node)
            else:
                yield OdtTailSpan(node)


class OdtSpanBase(OdtNode, IDocSpan):
    """Base for .odt span classes."""


class OdtHeadSpan(OdtSpanBase):
    """Beginning of a span."""

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        if self.node.text:
            yield self.node.text


class OdtTabSpan(OdtSpanBase):
    """A tab character."""

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        yield "\t"


class OdtSpanSpan(OdtSpanBase):
    """A proper text span."""

    def style_wpid(self) -> str | None:
        """Return the wpid for this span's style."""
        return self._node_ooget("text:style-name")

    def footnotes(self) -> t.Iterable[IDocFootnote]:
        """Yield an IDocFootnote object for each footnote in this span."""
        for fnr in self._node_xpath('text:note[@text:node-class="footnote"]'):
            yield OdtFootnote(fnr)

    def comments(self) -> t.Iterable[IDocComment]:
        """Yield an IDocComment object for each comment in this span."""
        yield from ()

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        if self.node.text:
            yield self.node.text


class OdtTailSpan(OdtSpanBase):
    """End of a span."""

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        if self.node.tail:
            yield self.node.tail


class OdtFootnote(OdtNode, IDocFootnote):
    """Footnote in .odt."""

    def paragraphs(self) -> t.Iterable[IDocParagraph]:
        """Yield an IDocParagraph object for each footnote paragraph."""
        for para in self._node_xpath("text:note-body/text-p"):
            yield OdtParagraph(para)
