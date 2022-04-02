#!/usr/bin/env python
""".odt file parsing"""
import contextlib
import logging
import zipfile

import lxml.etree

from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.styles import DocumentProperties


class OoXml:
    """Basic helper class for the OpenOffice XML format."""

    _NS = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }

    _FAMILY_TO_REALM = {
        "paragraph": "paragraph",
        "text": "character",
    }

    def _xpath(self, node, expr):
        return node.xpath(expr, namespaces=self._NS)

    def _ootag(self, tag):
        namespace, tag = tag.split(":", 1)
        prefix = self._NS[namespace]
        return f"{{{prefix}}}{tag}"

    def _ooget(self, node, tag):
        return node.get(self._ootag(tag))


class XodtInput(contextlib.ExitStack, OoXml, IDocumentInput):
    """A reader for .odt and .fodt."""

    def __init__(self, path: str, zipped: bool):
        super().__init__()
        self._zipped = zipped
        if zipped:
            self._zip = self._open_zip(path)
        else:
            self._flat = self._open_flat(path)
        self._content = self._load_xml("content.xml")
        self._initialize_properties()

    def _initialize_properties(self):
        self._properties = DocumentProperties(
            has_rtl=self._has_node(
                '//style:paragraph-properties[@style:writing-mode="rl-tb"]'
            )
        )

    def _has_node(self, ootag):
        for _ in self._xpath(self._content, ootag):
            return True
        return False

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        styles = self._load_xml("styles.xml")
        for snode in self._xpath(styles, "//office:styles/style:style"):
            yield self._style_kwargs(snode)
        for snode in self._xpath(self._content, "//office:automatic-styles/style:style"):
            yield self._style_kwargs(snode, automatic=True)

    def _style_kwargs(self, node, **extras):
        name = self._ooget(node, "style:name")
        style_kwargs = dict(extras) if extras else {}
        style_kwargs.update(
            {
                "realm": self._FAMILY_TO_REALM[self._ooget(node, "style:family")],
                "internal_name": self._ooget(node, "style:display-name") or name,
                "wpid": name,
                "parent_wpid": self._ooget(node, "style:parent-style-name"),
                "next_wpid": self._ooget(node, "style:next-style-name"),
            }
        )
        return style_kwargs

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for realm, tag in (("paragraph", "text:p"), ("character", "text:span")):
            for sname in self._xpath(self._content, f"//{tag}"):
                wpid = self._ooget(sname, "text:style-name")
                yield (realm, wpid)

    def paragraphs(self):
        """Yields a OdtParagraph object for each body paragraph."""
        for para in self._xpath(self._content, "//office:body/office:text/text:p"):
            yield OdtParagraph(self, para)

    def _open_zip(self, path):
        return self.enter_context(zipfile.ZipFile(path))

    @classmethod
    def _open_flat(cls, path):
        with open(path, "r", encoding="utf8") as fobj:
            return lxml.etree.parse(fobj).getroot()

    def _load_xml(self, path_in_zip):
        """Parse an XML file inside the zipped doc, return root node."""
        if not self._zipped:
            return self._flat

        try:
            with self._zip.open(path_in_zip) as fobj:
                return lxml.etree.parse(fobj).getroot()
        except KeyError:
            return None


class OdtNode(OoXml):
    """Base helper class for object which represent a node in a docx."""

    def __init__(self, doc, node):
        self.doc = doc
        self.node = node

    def _node_ooget(self, tag):
        return self._ooget(self.node, tag)

    def _node_xpath(self, expr):
        return self.node.xpath(expr, namespaces=self._NS)


class OdtParagraph(OdtNode, IDocumentParagraph):
    """A Paragraph inside a .docx."""

    def style_wpid(self):
        return self._node_ooget("text:style-name")

    def text(self):
        """Yields strings of plain text."""
        yield from self.node.itertext()

    def spans(self):
        """Yield OdtSpan per text span."""
        for event, node in lxml.etree.iterwalk(self.node, events=("start", "end")):
            if event == "start":
                if node.tag == self._ootag("text:tab"):
                    yield OdtTabSpan(self.doc, node)
                elif node.tag == self._ootag("text:span"):
                    yield OdtSpanSpan(self.doc, node)
                elif node.tag == self._ootag("text:p"):
                    yield OdtHeadSpan(self.doc, node)
                else:
                    logging.debug(
                        "Not sure what to do with a <%s> %r", node.tag, node.text[:8]
                    )
                    yield OdtHeadSpan(self.doc, node)
            else:
                yield OdtTailSpan(self.doc, node)


class OdtSpanBase(OdtNode, IDocumentSpan):
    """Base for .odt span classes"""
    def style_wpid(self):
        return None

    def footnotes(self):
        pass

    def comments(self):
        pass

    def text(self):
        pass


class OdtHeadSpan(OdtSpanBase):
    """Beginning of a span"""
    def text(self):
        if self.node.text:
            yield self.node.text


class OdtTabSpan(OdtSpanBase):
    """A tab character"""
    def text(self):
        yield "\t"


class OdtSpanSpan(OdtSpanBase):
    """A proper text span"""
    def style_wpid(self):
        return self._node_ooget("text:style-name")

    def footnotes(self):
        for fnr in self._node_xpath('text:note[@text:node-class="footnote"]'):
            yield OdtFootnote(self.doc, fnr)

    def comments(self):
        pass

    def text(self):
        if self.node.text:
            yield self.node.text


class OdtTailSpan(OdtSpanBase):
    """End of a span"""
    def text(self):
        if self.node.tail:
            yield self.node.tail


class OdtFootnote(OdtNode, IDocumentFootnote):
    """Footnote in .odt"""
    def paragraphs(self):
        for para in self._node_xpath("text:note-body/text-p"):
            yield OdtParagraph(self.doc, para)
