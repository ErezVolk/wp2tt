"""MS Word .docx parser."""

import contextlib
import typing as t

from pathlib import Path
from pathlib import PurePosixPath
from os import PathLike

from lxml import etree

from wp2tt.input import IDocumentComment
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentFormula
from wp2tt.input import IDocumentImage
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.input import IDocumentTable
from wp2tt.input import IDocumentTableRow
from wp2tt.input import IDocumentTableCell
from wp2tt.format import ManualFormat
from wp2tt.mathml import MathConverter
from wp2tt.styles import DocumentProperties
from wp2tt.zip import ZipDocument


class WordXml:
    """Basic helper class for the Word XML format."""

    _A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    _M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    _R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    _REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    _WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    _NS: t.Mapping[str, str] = {
        "a": _A,
        "m": _M,
        "r": _R,
        "rel": _REL,
        "w": _W,
        "w14": _W14,
        "wp": _WP,
    }

    @classmethod
    def xpath(
        cls, nodes: list[etree._Entity] | etree._Entity, expr: str,
    ) -> t.Iterable[etree._Entity]:
        """Wrap etree.xpath, with namespaces and multiple nodes."""
        if not isinstance(nodes, list):
            nodes = [nodes]
        for node in nodes:
            yield from node.xpath(expr, namespaces=cls._NS)

    @classmethod
    def _mtag(cls, tag: str) -> str:
        return f"{{{cls._M}}}{tag}"

    @classmethod
    def _wtag(cls, tag: str) -> str:
        return f"{{{cls._W}}}{tag}"

    @classmethod
    def _w14tag(cls, tag: str) -> str:
        return f"{{{cls._W14}}}{tag}"

    @classmethod
    def _rtag(cls, tag: str) -> str:
        return f"{{{cls._R}}}{tag}"

    @classmethod
    def _wval(
        cls, nodes: list[etree._Entity] | etree._Entity, prop: str,
    ) -> str | None:
        if not isinstance(nodes, list):
            nodes = [nodes]
        for node in nodes:
            for pnode in cls.xpath(node, prop):
                return pnode.get(cls._wtag("val"))
        return None


class DocxInput(contextlib.ExitStack, WordXml, IDocumentInput):
    """A .docx reader."""

    _wpid_prefix: str | None = None
    _name_prefix: str | None = None

    def __init__(self, path: PathLike) -> None:
        super().__init__()
        self._read_docx(path)
        self._initialize_properties()

    def _read_docx(self, path: PathLike) -> None:
        self.zip = self.enter_context(ZipDocument(path))
        self.document = self.zip.load_xml("word/document.xml")
        self.footnotes = self.zip.load_xml("word/footnotes.xml")
        self.comments = self.zip.load_xml("word/comments.xml")
        self.relationships = self.zip.load_xml("word/_rels/document.xml.rels")

    def _initialize_properties(self) -> None:
        self._properties = DocumentProperties(
            has_rtl=self._has_node("w:rtl"),
        )

    def _has_node(self, wtag: str) -> bool:
        for root in (self.document, self.footnotes, self.comments):
            if root is not None:
                for _ in self.xpath(root, wtag):
                    return True
        return False

    @property
    def properties(self) -> DocumentProperties:
        """Access this document's properties."""
        return self._properties

    def set_nth(self, nth: int) -> None:
        """Set prefix for non-clash in multi-file docs or something."""
        if nth > 1:
            self._wpid_prefix = f"d{nth}_"
            self._name_prefix = f"(D{nth}) "

    def export_wpid(self, wpid: str | None) -> str | None:
        """Make wpid unique in multi-input scenario."""
        if wpid is None or self._wpid_prefix is None:
            return wpid
        return f"{self._wpid_prefix}{wpid}"

    def export_name(self, name: str | None) -> str | None:
        """Make style name unique in multi-input scenario."""
        if name is None or self._name_prefix is None:
            return name
        return f"{self._name_prefix}{name}"

    def styles_defined(self) -> t.Iterable[dict[str, t.Any]]:
        """Yield a Style object kwargs for every style defined in the document."""
        styles = self.zip.load_xml("word/styles.xml")
        for stag in self.xpath(styles, "//w:style[@w:type][w:name[@w:val]]"):
            fmt = DocxSpan.node_format(stag)
            fmt &= ~(ManualFormat.LTR | ManualFormat.RTL)
            fmt |= DocxParagraph.node_format(stag)
            yield {
                "realm": stag.get(self._wtag("type")),
                "internal_name": self.export_name(self._wval(stag, "w:name")),
                "wpid": self.export_wpid(stag.get(self._wtag("styleId"))),
                "parent_wpid": self.export_wpid(self._wval(stag, "w:basedOn")),
                "next_wpid": self.export_wpid(self._wval(stag, "w:next")),
                "custom": stag.get(self._wtag("customStyle")),
                "fmt": fmt,
            }

    _TAG_TO_REALM: t.Mapping[str, str] = {
        "pStyle": "paragraph",
        "rStyle": "character",
        "tblStyle": "table",
    }
    _TAG_EXPRS = " or ".join(f"self::w:{tag}" for tag in _TAG_TO_REALM)
    TAG_XPATH = f"//*[{_TAG_EXPRS}]"
    WTAG_TO_REALM: t.Mapping[str, str] = {
        WordXml._wtag(tag): realm for tag, realm in _TAG_TO_REALM.items()
    }

    def styles_in_use(self) -> t.Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        for node in (self.document, self.footnotes, self.comments):
            if node is None:
                continue
            for snode in self.xpath(node, self.TAG_XPATH):
                wpid = self.export_wpid(snode.get(self._wtag("val")))
                yield (self.WTAG_TO_REALM[snode.tag], wpid)

    def paragraphs(self) -> t.Iterable["DocxParagraph | DocxTable"]:
        """Yield a DocxParagraph object for each body paragraph."""
        # ex. "//w:body/w:p[not(preceding-sibling::w:p/w:pPr/w:rPr/w:del)]"
        for node in self.xpath(self.document, "//w:body/*[self::w:p or self::w:tbl]"):
            if node.tag == self._wtag("tbl"):
                yield DocxTable(self, node)
            elif not node.get("__wp2tt_skip__"):
                yield DocxParagraph(self, node)


class DocxNode(WordXml):
    """Base helper class for object which represent a node in a docx.

    In special cases, this actually represents a list of consecutive nodes,
    used for parargaph breaks deleted with track changes.
    """

    def __init__(self, doc: DocxInput, node: etree._Entity) -> None:
        self.doc = doc
        self.head_node = node
        self.nodes = [node]

    def add_node(self, node: etree._Entity) -> None:
        """Extend the list of nodes."""
        self.nodes.append(node)

    def _node_wtag(self, tag: str) -> str | None:
        return self.head_node.get(self._wtag(tag))

    def _node_xpath(self, expr: str) -> t.Iterable[etree._Entity]:
        for node in self.nodes:
            yield from node.xpath(expr, namespaces=self._NS)

    def _node_wval(self, prop: str) -> str | None:
        return self._node_wattr(prop, "val")

    def _node_wtype(self, prop: str) -> str | None:
        return self._node_wattr(prop, "type")

    def _node_wtypes(self, prop: str) -> t.Iterable[str]:
        yield from self._node_wattrs(prop, "type")

    def _node_wattr(self, prop: str, attr: str) -> str | None:
        for value in self._node_wattrs(prop, attr):
            return value
        return None

    def _node_wattrs(self, prop: str, attr: str) -> t.Iterable[str]:
        tag = self._wtag(attr)
        for node in self.nodes:
            for pnode in self.xpath(node, prop):
                value = pnode.get(tag)
                if value is not None:
                    yield value


class DocxParagraph(DocxNode, IDocumentParagraph):
    """A Paragraph inside a .docx."""

    R_XPATH = "w:r | w:ins/w:r | m:oMath"
    T_XPATH = "w:r/w:t | w:ins/w:r/w:t"
    SNIPPET_LEN = 10

    def __init__(self, doc: DocxInput, para: etree._Entity) -> None:
        super().__init__(doc, para)
        while self.is_nonfinal(para):
            self.add_node(para := para.getnext())
            para.set("__wp2tt_skip__", "yes")
        self._para_ids = [self._get_para_id(node) for node in self.nodes]

    def __repr__(self) -> str:
        """Describe the paragraph object."""
        pids = "/".join(self._para_ids)
        return f"<w:p {pids}>"

    def _get_para_id(self, para: etree._Entity) -> str:
        """Create a hopefully unique paragraph ID."""
        w14id = para.get(self._w14tag("paraId"))
        if w14id:
            return f'w14:paraId="{w14id}"'

        texts: list[str] = []
        tlen = 0
        for tnode in para.xpath(self.T_XPATH, namespaces=self._NS):
            text = tnode.text
            if not text:  # May also be None
                continue
            texts.append(text)
            tlen += len(text)
            if tlen >= self.SNIPPET_LEN:
                break
        text = "".join(texts)
        if len(text) <= self.SNIPPET_LEN:
            return f'"{text}"'
        return f'"{text[:self.SNIPPET_LEN-3]}"...'

    def is_nonfinal(self, para: etree._Entity) -> bool:
        """Check if a <w:p> para has deleted, tracked newline."""
        for _ in self.xpath(para, "./w:pPr/w:rPr/w:del"):
            return True
        return False

    def style_wpid(self) -> str | None:
        """Get MS Word's internal ID for this style."""
        return self.doc.export_wpid(self._node_wval("./w:pPr/w:pStyle"))

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        for node in self._node_xpath(self.T_XPATH):
            if node.text:
                yield node.text

    def chunks(self) -> t.Iterable["DocxSpan | DocxImage | DocxFormula"]:
        """Yield DocxSpan per text span."""
        for node in self._node_xpath(self.R_XPATH):
            if node.tag == self._mtag("oMath"):
                yield DocxFormula(node)
            else:
                for drawing in self.xpath(node, "w:drawing[//a:blip[@r:embed]]"):
                    yield DocxImage(self.doc, drawing)
                    break
                else:
                    yield DocxSpan(self.doc, node)

    def format(self) -> ManualFormat:
        """Return manual formatting on this paragraph."""
        return self.node_format(self.nodes)

    @classmethod
    def node_format(cls, nodes: list[etree._Entity]) -> ManualFormat:
        """Return manual formatting on a paragraph/style."""
        fmt = ManualFormat.LTR
        justification = cls._wval(nodes, "w:pPr/w:jc")
        if justification == "center":
            fmt = fmt | ManualFormat.CENTERED
        elif justification == "both":
            fmt = fmt | ManualFormat.JUSTIFIED
        for _ in cls.xpath(nodes, "w:pPr/w:bidi"):
            fmt = (fmt & ~ManualFormat.LTR) | ManualFormat.RTL
        return fmt

    def is_page_break(self) -> bool:
        """Check if the paragraph is a page break."""
        for break_type in self._node_wtypes("w:r/w:br | w:ins/w:r/w:br"):
            if break_type == "page":
                return True
        return False


class DocxSpan(DocxNode, IDocumentSpan):
    """A span of characters inside a .docx."""

    def __repr__(self) -> str:
        """Describe the paragraph object."""
        return repr(" ".join(t for t in self.text() if t is not None))

    def style_wpid(self) -> str | None:
        """Get this Span's style."""
        return self.doc.export_wpid(self._node_wval("w:rPr/w:rStyle"))

    def footnotes(self) -> t.Iterable["DocxFootnote"]:
        """Yield foornotes in this span."""
        for fnr in self._node_xpath("w:footnoteReference"):
            yield DocxFootnote(self.doc, fnr)

    def comments(self) -> t.Iterable["DocxComment"]:
        """Yield foornotes in this span."""
        for cmr in self._node_xpath("w:commentReference"):
            yield DocxComment(self.doc, cmr)

    def format(self) -> ManualFormat:
        """Get manual formatting for this span."""
        return self.node_format(self.nodes)

    @classmethod
    def node_format(cls, nodes: list[etree._Entity]) -> ManualFormat:
        """Get manual formatting for a span/style."""
        fmt = ManualFormat.LTR
        for _ in cls.xpath(nodes, "w:rPr/w:b | w:rPr/w:bCs"):
            fmt = fmt | ManualFormat.BOLD
        for _ in cls.xpath(nodes, "w:rPr/w:i | w:rPr/w:iCs"):
            fmt = fmt | ManualFormat.ITALIC
        for vnode in cls.xpath(nodes, "w:rPr/w:vertAlign"):
            vval = vnode.get(cls._wtag("val"))
            if vval == "subscript":
                fmt |= ManualFormat.SUBSCRIPT
            elif vval == "superscript":
                fmt |= ManualFormat.SUPERSCRIPT
        for vnode in cls.xpath(nodes, "w:rPr/w:position"):
            vval = float(vnode.get(cls._wtag("val")))
            if vval < 0:
                fmt |= ManualFormat.LOWERED
            elif vval > 0:
                fmt |= ManualFormat.RAISED
        for _ in cls.xpath(nodes, "w:rPr/w:rtl"):
            fmt = fmt & ~ManualFormat.LTR | ManualFormat.RTL
        return fmt

    def text(self) -> t.Iterable[str]:
        """Yield chunks of text."""
        for node in self._node_xpath("w:tab | w:t"):
            if node.tag == self._wtag("tab"):
                yield "\t"
            elif node.text:
                yield node.text


class DocxImage(DocxNode, IDocumentImage):
    """An image inside a .docx."""

    descr: str | None = None
    target: PurePosixPath

    def __init__(self, doc: DocxInput, drawing: etree._Entity) -> None:
        super().__init__(doc, drawing)
        for prop in self._node_xpath("./wp:inline/wp:docPr[@descr]"):
            self.descr = prop.get("descr")

        rels = self.doc.relationships
        for blip in self._node_xpath(".//a:blip[@r:embed]"):
            rid = blip.get(self._rtag("embed"))
            for rel in self.doc.xpath(rels, f'//rel:Relationship[@Id="{rid}"]'):
                self.target = PurePosixPath("word") / rel.get("Target")

    def alt_text(self) -> str | None:
        """Get alt-text for image."""
        return self.descr

    def suffix(self) -> str:
        """Get image suffix (file extension)."""
        return self.target.suffix

    def save(self, path: PathLike) -> None:
        """Extract image."""
        with self.doc.zip.open(str(self.target)) as ifo, Path(path).open("wb") as ofo:
            ofo.write(ifo.read())


class DocxFormula(IDocumentFormula):
    """A formula inside a .docx."""

    def __init__(self, node: etree._Entity) -> None:
        self.node = node

    def raw(self) -> bytes:
        """Get formula's Word XML."""
        return etree.tostring(self.node, pretty_print=True)

    def mathml(self) -> str:
        """Convert formula to MathML."""
        mathml = MathConverter.omml_to_mathml(self.node)
        encoded = etree.tostring(mathml, pretty_print=True)
        return encoded.decode()


class DocxTable(DocxNode, IDocumentTable):
    """A table inside a .docx."""

    def __init__(self, doc: DocxInput, node: etree._Entity) -> None:
        super().__init__(doc, node)
        self.orows = [DocxTableRow(doc, row) for row in self.xpath(node, "./w:tr")]
        self.n_header_rows = sum(row.is_header() for row in self.orows)
        self.n_rows = len(self.orows)
        self.n_cols = max(
            sum(cell.shape[1] for cell in row.cells()) for row in self.orows
        )

    def style_wpid(self) -> str | None:
        """Return the wpid for this table's style."""
        return self.doc.export_wpid(self._node_wval("w:tblPr/w:tblStyle"))

    def format(self) -> ManualFormat:
        """Get table formatting (RTL is all we care about)."""
        for _ in self._node_xpath("./w:tblPr/w:bidiVisual"):
            return ManualFormat.RTL
        return ManualFormat.LTR

    @property
    def shape(self) -> tuple[int, int]:
        """Return (number of rows, number of columns)."""
        return (self.n_rows, self.n_cols)

    @property
    def header_rows(self) -> int:
        """Return number of header rows."""
        return self.n_header_rows

    def rows(self) -> t.Iterable["DocxTableRow"]:
        """Iterate the rows of the table."""
        yield from self.orows


class DocxTableRow(DocxNode, IDocumentTableRow):
    """A table row."""

    def __init__(self, doc: DocxInput, node: etree._Entity) -> None:
        super().__init__(doc, node)
        self.ocells = [DocxTableCell(doc, cell) for cell in self.xpath(node, "./w:tc")]

    def is_header(self) -> bool:
        """Check if this row is a header row."""
        for _ in self._node_xpath("./w:trPr/w:tblHeader"):
            return True
        return False

    def cells(self) -> t.Iterable["DocxTableCell"]:
        """Yield all cells in the row."""
        yield from self.ocells


class DocxTableCell(DocxNode, IDocumentTableCell):
    """A table cell."""

    def __init__(self, doc: DocxInput, node: etree._Entity) -> None:
        super().__init__(doc, node)
        try:
            self.span = int(self._wval(node, "./w:tcPr/w:gridSpan") or "1")
        except ValueError:
            self.span = 1

    @property
    def shape(self) -> tuple[int, int]:
        """Return (number of rows, number of columns)."""
        return (1, self.span)

    def contents(self) -> DocxParagraph:
        """Get the contents of this cell."""
        for pnode in self.xpath(self.nodes, "./w:p"):
            return DocxParagraph(self.doc, pnode)
        raise RuntimeError("Table cell without a paragraph")


class DocxFootnote(DocxNode, IDocumentFootnote):
    """IDocumentFootnote for .docx file."""

    def paragraphs(self) -> t.Iterable[DocxParagraph]:
        """Yield DocxParagraph for each paragraph in a footnote."""
        fnid = self._node_wtag("id")
        for para in self.xpath(self.doc.footnotes, f'w:footnote[@w:id="{fnid}"]/w:p'):
            yield DocxParagraph(self.doc, para)


class DocxComment(DocxNode, IDocumentComment):
    """IDocumentComment for .docx file."""

    def paragraphs(self) -> t.Iterable[DocxParagraph]:
        """Yield DocxParagraph for each paragraph in a comment."""
        cmid = self._node_wtag("id")
        for para in self.xpath(self.doc.comments, f'w:comment[@w:id="{cmid}"]/w:p'):
            yield DocxParagraph(self.doc, para)
