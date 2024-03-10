#!/usr/bin/env python3
"""Read Markdown document"""
# pylint: disable=unused-argument
import logging
from pathlib import Path
import contextlib

from typing import Iterable

import mistune
from lxml import etree

from wp2tt.input import IDocumentComment
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.styles import DocumentProperties


class MarkdownUnRenderer:
    """Mistune callback to convert Markdown to XML"""

    NAMELESS_P_PRE = '<p wpid="normal">'
    NAMELESS_P_POST = "</p>"

    def __init__(self, **kwargs):
        self.options = kwargs

    def placeholder(self):
        """Mistune element"""
        return ""

    def header(self, text, level, raw=None):
        """Mistune element"""
        return f'<p wpid="header">{text}</p>'

    def text(self, text):
        """Mistune element"""
        return text

    def paragraph(self, text):
        """Mistune element"""
        return f"{self.NAMELESS_P_PRE}{text}{self.NAMELESS_P_POST}"

    def emphasis(self, text):
        """Mistune element"""
        return f'<s wpid="emphasis">{text}</s>'

    def double_emphasis(self, text):
        """Mistune element"""
        return f'<s wpid="doule emphasis">{text}</s>'

    def autolink(self, link, is_email=False):
        """Mistune element"""
        return f'<s wpid="link">{link}</s>'

    def link(self, link, title, content):
        """Mistune element"""
        return f'<s wpid="link" title="{title}">{content}</s>'

    def list_item(self, text):
        """Mistune element"""
        if text.startswith(self.NAMELESS_P_PRE) and text.endswith(self.NAMELESS_P_POST):
            text = text[len(self.NAMELESS_P_PRE) : -len(self.NAMELESS_P_POST)]
        return f'<p wpid="list item">{text}</p>'

    def list(self, text, ordered=True):
        """Mistune element"""
        return text

    def block_html(self, html):
        """Mistune element"""
        logging.warning("HTML is corrently ignored in Markdown")
        return ""

    def block_code(self, code, language=None):
        """Mistune element"""
        raise NotImplementedError("block_code()")

    def block_quote(self, text):
        """Mistune element"""
        raise NotImplementedError("block_quote()")

    def hrule(self):
        """Mistune element"""
        raise NotImplementedError("hrule()")

    def table(self, header, body):
        """Mistune element"""
        raise NotImplementedError("table()")

    def table_row(self, content):
        """Mistune element"""
        raise NotImplementedError("table_row()")

    def table_cell(self, content, **flags):
        """Mistune element"""
        raise NotImplementedError("table_cell()")

    def codespan(self, text):
        """Mistune element"""
        raise NotImplementedError("codespan()")

    def image(self, src, title, alt_text):
        """Mistune element"""
        raise NotImplementedError("image()")

    def linebreak(self):
        """Mistune element"""
        raise NotImplementedError("linebreak()")

    def newline(self):
        """Mistune element"""
        raise NotImplementedError("newline()")

    def strikethrough(self, text):
        """Mistune element"""
        raise NotImplementedError("strikethrough()")

    def inline_html(self, text):
        """Mistune element"""
        raise NotImplementedError("inline_html()")

    def footnote_ref(self, key, index):
        """Mistune element"""
        raise NotImplementedError("footnote_ref()")

    def footnote_item(self, key, text):
        """Mistune element"""
        raise NotImplementedError("footnote_item()")

    def footnotes(self, text):
        """Mistune element"""
        raise NotImplementedError("footnotes()")


class MarkdownInput(IDocumentInput, contextlib.ExitStack):
    """A Markdown reader."""

    def __init__(self, path: Path):
        super().__init__()
        self._read_markdown(path)
        self._properties = DocumentProperties(has_rtl=False)

    def _read_markdown(self, path: Path):
        renderer = MarkdownUnRenderer()
        parse = mistune.Markdown(renderer=renderer)
        with open(path, "r", encoding="utf8") as mdfo:
            xml = parse(mdfo.read())
        self._root = etree.fromstring(f"<document>{xml}</document>")
        print(
            etree.tostring(
                self._root, pretty_print=True, encoding="utf-8", xml_declaration=True
            ).decode("utf-8")
        )

    @property
    def properties(self):
        return self._properties

    def styles_defined(self) -> Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        for realm, wpid in self.styles_in_use():
            yield {
                "realm": realm or "",
                "internal_name": wpid or "",
                "wpid": wpid or "",
            }

    def xpath(self, expr) -> Iterable[etree._Entity]:
        """Wrapper for `lxml.xpath()`"""
        yield from self._root.xpath(expr)

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for node in self.xpath("//p[@wpid]"):
            yield "paragraph", node.get("wpid")
        for node in self.xpath("//s[@wpid]"):
            yield "character", node.get("wpid")
        for node in self.xpath("//li"):
            yield "paragraph", "list item"
            break

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each body paragraph."""
        for para in self._root:
            if para.tag in ("p"):
                yield MarkdownParagraph(para)


class MarkdownParagraph(IDocumentParagraph):
    """A Paragraph inside a document."""

    def __init__(self, node):
        self.node = node

    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        return self.node.get("wpid")

    def text(self):
        """Yields strings of plain text."""
        for span in self.spans():
            yield from span.text()

    def chunks(self):
        """Yield a MarkdownSpan per text span."""
        yield MarkdownHeadSpan(self.node)
        for span in self.node.xpath("s"):
            yield MarkdownSpanSpan(span)
            yield MarkdownTailSpan(span)
        yield MarkdownTailSpan(self.node)


class MarkdownSpanBase(IDocumentSpan):
    """Base class for our span types"""
    def __init__(self, node):
        self.node = node

    def style_wpid(self):
        return None

    def footnotes(self) -> Iterable["IDocumentFootnote"]:
        yield from ()

    def comments(self) -> Iterable["IDocumentComment"]:
        yield from ()


class MarkdownHeadSpan(MarkdownSpanBase):
    """Head of XML node"""
    def text(self):
        """Yields strings of plain text."""
        if self.node.text:
            yield self.node.text


class MarkdownTailSpan(MarkdownSpanBase):
    """Tail of XML node"""
    def text(self):
        """Yields strings of plain text."""
        if self.node.tail:
            yield self.node.tail


class MarkdownSpanSpan(MarkdownSpanBase):
    """A span of characters inside a document."""

    def style_wpid(self):
        """Returns the wpid for this span's style."""
        return self.node.get("wpid")

    def text(self):
        """Yields strings of plain text."""
        yield self.node.text or ""


class MarkdownFootnote(IDocumentFootnote):
    """A footnote."""

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each footnote paragraph."""
        raise NotImplementedError()
