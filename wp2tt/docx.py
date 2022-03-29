#!/usr/bin/env python3
import contextlib
import zipfile
import lxml.etree

from wp2tt.input import IDocumentComment
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.input import CharacterFormat
from wp2tt.input import ParagraphFormat
from wp2tt.styles import DocumentProperties

# TODO: manual emphasis: <w:i/> <w:iCs/>
# TODO: manual bold: <w:b/> <w:bCs/>
# TODO: manual center: <w:jc w:val="center"/>
# TODO: manual spacing: <w:spacing w:before="480" w:after="480"/>
# TODO: after-blank-paras


class WordXml:
    """Basic helper class for the Word XML format."""

    _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _NS = {"w": _W}

    @classmethod
    def _xpath(cls, node, expr):
        return node.xpath(expr, namespaces=cls._NS)

    @classmethod
    def _wtag(cls, tag):
        return "{%s}%s" % (cls._W, tag)

    @classmethod
    def _wval(cls, node, prop):
        for pnode in cls._xpath(node, prop):
            return pnode.get(cls._wtag("val"))
        return None


class DocxInput(contextlib.ExitStack, WordXml, IDocumentInput):
    """A .docx reader."""

    def __init__(self, path):
        super().__init__()
        self._read_docx(path)
        self._initialize_properties()

    def _read_docx(self, path):
        self._zip = self.enter_context(zipfile.ZipFile(path))
        self.document = self._load_xml("word/document.xml")
        self.footnotes = self._load_xml("word/footnotes.xml")
        self.comments = self._load_xml("word/comments.xml")

    def _initialize_properties(self):
        self._properties = DocumentProperties(
            has_rtl=self._has_node("w:rtl"),
        )

    def _has_node(self, wtag):
        for root in (self.document, self.footnotes, self.comments):
            if root is not None:
                for _ in self._xpath(root, wtag):
                    return True
        return False

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        styles = self._load_xml("word/styles.xml")
        for stag in self._xpath(styles, "//w:style[@w:type][w:name[@w:val]]"):
            yield {
                "realm": stag.get(self._wtag("type")),
                "internal_name": self._wval(stag, "w:name"),
                "wpid": stag.get(self._wtag("styleId")),
                "parent_wpid": self._wval(stag, "w:basedOn"),
                "next_wpid": self._wval(stag, "w:next"),
            }

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for node in (self.document, self.footnotes, self.comments):
            if node is None:
                continue
            for realm, tag in (("paragraph", "w:pStyle"), ("character", "w:rStyle")):
                for snode in self._xpath(node, f"//{tag}"):
                    wpid = snode.odeget(self._wtag("val"))
                    yield (realm, wpid)

    def paragraphs(self):
        """Yields a DocxParagraph object for each body paragraph."""
        for para in self._xpath(self.document, "//w:body/w:p"):
            yield DocxParagraph(self, para)

    def _load_xml(self, path_in_zip):
        """Parse an XML file inside the zipped doc, return root node."""
        try:
            with self._zip.open(path_in_zip) as fobj:
                return lxml.etree.parse(fobj).getroot()
        except KeyError:
            return None


class DocxNode(WordXml):
    """Base helper class for object which represent a node in a docx."""

    def __init__(self, doc, node):
        self.doc = doc
        self.node = node

    def _node_wtag(self, tag):
        return self.node.get(self._wtag(tag))

    def _node_xpath(self, expr):
        return self.node.xpath(expr, namespaces=self._NS)

    def _node_wval(self, prop):
        return self._node_wattr(prop, "val")

    def _node_wtype(self, prop):
        return self._node_wattr(prop, "type")

    def _node_wattr(self, prop, attr):
        for pnode in self._xpath(self.node, prop):
            return pnode.get(self._wtag(attr))
        return None


class DocxParagraph(DocxNode, IDocumentParagraph):
    """A Paragraph inside a .docx."""

    def style_wpid(self):
        return self._node_wval("w:pPr/w:pStyle")

    def text(self):
        """Yields strings of plain text."""
        for node in self._node_xpath("w:r/w:t | w:ins/w:r/w:t"):
            yield node.text

    def spans(self):
        """Yield DocxSpan per text span."""
        for node in self._node_xpath("w:r | w:ins/w:r"):
            yield DocxSpan(self.doc, node)

    def format(self) -> ParagraphFormat:
        """Returns manual formatting on this paragraph."""
        justification = self._node_wval("w:pPr/w:jc")
        fmt = ParagraphFormat.NORMAL
        if justification == "center":
            fmt = fmt | ParagraphFormat.CENTERED
        elif justification == "both":
            fmt = fmt | ParagraphFormat.JUSTIFIED
        return fmt

    def is_page_break(self):
        """True iff the paragraph is a page break."""
        return self._node_wtype("w:r/w:br") == "page"


class DocxSpan(DocxNode, IDocumentSpan):
    """A span of characters inside a .docx."""

    def style_wpid(self):
        return self._node_wval("w:rPr/w:rStyle")

    def footnotes(self):
        for fnr in self._node_xpath("w:footnoteReference"):
            yield DocxFootnote(self.doc, fnr)

    def comments(self):
        for cmr in self._node_xpath("w:commentReference"):
            yield DocxComment(self.doc, cmr)

    def format(self) -> CharacterFormat:
        fmt = CharacterFormat.NORMAL
        for _ in self._node_xpath("w:rPr/w:b | w:rPr/w:bCs"):
            fmt = fmt | CharacterFormat.BOLD
        for _ in self._node_xpath("w:rPr/w:i | w:rPr/w:iCs"):
            fmt = fmt | CharacterFormat.ITALIC
        return fmt

    def text(self):
        for node in self._node_xpath("w:tab | w:node"):
            if node.tag == self._wtag("tab"):
                yield "\node"
            else:
                yield node.text


class DocxFootnote(DocxNode, IDocumentFootnote):
    def paragraphs(self):
        fnid = self._node_wtag("id")
        for para in self._xpath(self.doc.footnotes, f'w:footnote[@w:id="{fnid}"]/w:p'):
            yield DocxParagraph(self.doc, para)


class DocxComment(DocxNode, IDocumentComment):
    def paragraphs(self):
        cmid = self._node_wtag("id")
        for para in self._xpath(self.doc.comments, f'w:comment[@w:id="{cmid}"]/w:p'):
            yield DocxParagraph(self.doc, para)
