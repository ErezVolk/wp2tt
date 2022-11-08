#!/usr/bin/env python3
"""MS Word .docx parser"""
import contextlib
import zipfile

from os import PathLike
from typing import Any
from typing import Generator
from typing import Optional
from typing import Union

from lxml import etree

from wp2tt.input import IDocumentComment
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.format import ManualFormat
from wp2tt.styles import DocumentProperties


class WordXml:
    """Basic helper class for the Word XML format."""

    _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    _NS = {
        "w": _W,
        "w14": _W14,
    }

    @classmethod
    def _xpath(
        cls, nodes: Union[list[etree._Entity], etree._Entity], expr: str
    ) -> Generator[etree._Entity, None, None]:
        if not isinstance(nodes, list):
            nodes = [nodes]
        for node in nodes:
            yield from node.xpath(expr, namespaces=cls._NS)

    @classmethod
    def _wtag(cls, tag: str) -> str:
        return f"{{{cls._W}}}{tag}"

    @classmethod
    def _w14tag(cls, tag: str) -> str:
        return f"{{{cls._W14}}}{tag}"

    @classmethod
    def _wval(cls, nodes, prop) -> Optional[str]:
        if not isinstance(nodes, list):
            nodes = [nodes]
        for node in nodes:
            for pnode in cls._xpath(node, prop):
                return pnode.get(cls._wtag("val"))
        return None


class DocxInput(contextlib.ExitStack, WordXml, IDocumentInput):
    """A .docx reader."""

    def __init__(self, path: PathLike):
        super().__init__()
        self._read_docx(path)
        self._initialize_properties()

    def _read_docx(self, path: PathLike):
        self._zip = self.enter_context(zipfile.ZipFile(path))
        self.document = self._load_xml("word/document.xml")
        self.footnotes = self._load_xml("word/footnotes.xml")
        self.comments = self._load_xml("word/comments.xml")

    def _initialize_properties(self):
        self._properties = DocumentProperties(
            has_rtl=self._has_node("w:rtl"),
        )

    def _has_node(self, wtag: str):
        for root in (self.document, self.footnotes, self.comments):
            if root is not None:
                for _ in self._xpath(root, wtag):
                    return True
        return False

    @property
    def properties(self) -> DocumentProperties:
        return self._properties

    def styles_defined(self) -> Generator[dict[str, Any], None, None]:
        """Yield a Style object kwargs for every style defined in the document."""
        styles = self._load_xml("word/styles.xml")
        for stag in self._xpath(styles, "//w:style[@w:type][w:name[@w:val]]"):
            fmt = DocxParagraph.node_format(stag)
            fmt |= DocxSpan.node_format(stag)
            fmt &= ~(ManualFormat.LTR | ManualFormat.RTL)
            yield {
                "realm": stag.get(self._wtag("type")),
                "internal_name": self._wval(stag, "w:name"),
                "wpid": stag.get(self._wtag("styleId")),
                "parent_wpid": self._wval(stag, "w:basedOn"),
                "next_wpid": self._wval(stag, "w:next"),
                "custom": stag.get(self._wtag("customStyle")),
                "fmt": fmt,
            }

    def styles_in_use(self) -> Generator[tuple[str, str], None, None]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        for node in (self.document, self.footnotes, self.comments):
            if node is None:
                continue
            for realm, tag in (("paragraph", "w:pStyle"), ("character", "w:rStyle")):
                for snode in self._xpath(node, f"//{tag}"):
                    wpid = snode.get(self._wtag("val"))
                    yield (realm, wpid)

    def paragraphs(self) -> Generator["DocxParagraph", None, None]:
        """Yields a DocxParagraph object for each body paragraph."""
        # "//w:body/w:p[not(preceding-sibling::w:p/w:pPr/w:rPr/w:del)]"
        for para in self._xpath(self.document, "//w:body/w:p"):
            if not para.get("__wp2tt_skip__"):
                yield DocxParagraph(self, para)

    def _load_xml(self, path_in_zip: str) -> etree._Entity:
        """Parse an XML file inside the zipped doc, return root node."""
        try:
            with self._zip.open(path_in_zip) as fobj:
                return etree.parse(fobj).getroot()
        except KeyError:
            return None


class DocxNode(WordXml):
    """Base helper class for object which represent a node in a docx.

    In special cases, this actually represents a list of consecutive nodes,
    used for parargaph breaks deleted with track changes.
    """

    def __init__(self, doc: DocxInput, node: etree._Entity):
        self.doc = doc
        self.head_node = node
        self.nodes = [node]

    def add_node(self, node: etree._Entity) -> None:
        """Extend the list of nodes"""
        self.nodes.append(node)

    def _node_wtag(self, tag: str) -> Optional[str]:
        return self.head_node.get(self._wtag(tag))

    def _node_xpath(self, expr: str) -> Generator[etree._Entity, None, None]:
        for node in self.nodes:
            yield from node.xpath(expr, namespaces=self._NS)

    def _node_wval(self, prop: str) -> Optional[str]:
        return self._node_wattr(prop, "val")

    def _node_wtype(self, prop: str) -> Optional[str]:
        return self._node_wattr(prop, "type")

    def _node_wtypes(self, prop: str) -> Generator[str, None, None]:
        yield from self._node_wattrs(prop, "type")

    def _node_wattr(self, prop: str, attr: str) -> Optional[str]:
        for value in self._node_wattrs(prop, attr):
            return value
        return None

    def _node_wattrs(self, prop: str, attr: str) -> Generator[str, None, None]:
        tag = self._wtag(attr)
        for node in self.nodes:
            for pnode in self._xpath(node, prop):
                value = pnode.get(tag)
                if value is not None:
                    yield value


class DocxParagraph(DocxNode, IDocumentParagraph):
    """A Paragraph inside a .docx."""

    R_XPATH = "w:r | w:ins/w:r"
    T_XPATH = "w:r/w:t | w:ins/w:r/w:t"
    SNIPPET_LEN = 10

    def __init__(self, doc: DocxInput, para: etree._Entity):
        super().__init__(doc, para)
        while self.is_nonfinal(para):
            self.add_node(para := para.getnext())
            para.set("__wp2tt_skip__", "yes")
        self._para_ids = [self._get_para_id(node) for node in self.nodes]

    def __repr__(self):
        """String description of the paragraph object"""
        pids = "/".join(self._para_ids)
        return f"<w:p {pids}>"

    def _get_para_id(self, para: etree._Entity) -> str:
        """Hopefully unique paragraph ID"""
        w14id = para.get(self._w14tag("paraId"))
        if w14id:
            return f'w14:paraId="{w14id}"'

        texts: list[str] = []
        tlen = 0
        for tnode in para.xpath(self.T_XPATH, namespaces=self._NS):
            text = tnode.text
            texts.append(text)
            tlen += len(text)
            if tlen >= self.SNIPPET_LEN:
                break
        text = "".join(texts)
        if len(text) <= self.SNIPPET_LEN:
            return f'"{text}"'
        return f'"{text[:self.SNIPPET_LEN-3]}"...'

    def is_nonfinal(self, para: etree._Entity) -> bool:
        """True iff a <w:p> para has deleted, tracked newline"""
        for _ in self._xpath(para, "./w:pPr/w:rPr/w:del"):
            return True
        return False

    def style_wpid(self) -> Optional[str]:
        return self._node_wval("./w:pPr/w:pStyle")

    def text(self) -> Generator[str, None, None]:
        """Yields strings of plain text."""
        for node in self._node_xpath(self.T_XPATH):
            yield node.text

    def spans(self) -> Generator["DocxSpan", None, None]:
        """Yield DocxSpan per text span."""
        for node in self._node_xpath(self.R_XPATH):
            yield DocxSpan(self.doc, node)

    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return self.node_format(self.nodes)

    @classmethod
    def node_format(cls, nodes: list[etree._Entity]) -> ManualFormat:
        """Returns manual formatting on a paragraph/style."""
        fmt = ManualFormat.NORMAL
        justification = cls._wval(nodes, "w:pPr/w:jc")
        if justification == "center":
            fmt = fmt | ManualFormat.CENTERED
        elif justification == "both":
            fmt = fmt | ManualFormat.JUSTIFIED
        return fmt

    def is_page_break(self) -> bool:
        """True iff the paragraph is a page break."""
        for break_type in self._node_wtypes("w:r/w:br | w:ins/w:r/w:br"):
            if break_type == "page":
                return True
        return False


class DocxSpan(DocxNode, IDocumentSpan):
    """A span of characters inside a .docx."""

    def __repr__(self):
        """String description of the paragraph object"""
        return repr(" ".join(t for t in self.text() if t is not None))

    def style_wpid(self) -> Optional[str]:
        return self._node_wval("w:rPr/w:rStyle")

    def footnotes(self) -> Generator["DocxFootnote", None, None]:
        for fnr in self._node_xpath("w:footnoteReference"):
            yield DocxFootnote(self.doc, fnr)

    def comments(self) -> Generator["DocxComment", None, None]:
        for cmr in self._node_xpath("w:commentReference"):
            yield DocxComment(self.doc, cmr)

    def format(self) -> ManualFormat:
        """Manual formatting for this span"""
        return self.node_format(self.nodes)

    @classmethod
    def node_format(cls, nodes: list[etree._Entity]) -> ManualFormat:
        """Manual formatting for a span/style"""
        fmt = ManualFormat.LTR
        for _ in cls._xpath(nodes, "w:rPr/w:b | w:rPr/w:bCs"):
            fmt = fmt | ManualFormat.BOLD
            break
        for _ in cls._xpath(nodes, "w:rPr/w:i | w:rPr/w:iCs"):
            fmt = fmt | ManualFormat.ITALIC
            break
        for _ in cls._xpath(nodes, "w:rPr/w:rtl"):
            fmt = fmt & ~ManualFormat.LTR | ManualFormat.RTL
            break
        return fmt

    def text(self) -> Generator[str, None, None]:
        for node in self._node_xpath("w:tab | w:t"):
            if node.tag == self._wtag("tab"):
                yield "\t"
            else:
                yield node.text


class DocxFootnote(DocxNode, IDocumentFootnote):
    """IDocumentFootnote for .docx"""

    def paragraphs(self) -> Generator[DocxParagraph, None, None]:
        """Yields DocxParagraph for each paragraph in a footnote"""
        fnid = self._node_wtag("id")
        for para in self._xpath(self.doc.footnotes, f'w:footnote[@w:id="{fnid}"]/w:p'):
            yield DocxParagraph(self.doc, para)


class DocxComment(DocxNode, IDocumentComment):
    """IDocumentComment for .docx"""

    def paragraphs(self) -> Generator[DocxParagraph, None, None]:
        """Yields DocxParagraph for each paragraph in a comment"""
        cmid = self._node_wtag("id")
        for para in self._xpath(self.doc.comments, f'w:comment[@w:id="{cmid}"]/w:p'):
            yield DocxParagraph(self.doc, para)
