#!/usr/bin/env python3
# TODO: mistune, "renderer"
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.styles import DocumentProperties


class MarkdownInput(IDocumentInput):
    """A Markdown reader."""
    def __init__(self, path):
        super().__init__()
        self._read_markdown(path)
        self._properties = DocumentProperties()

    def _read_markdown(self, path):
        raise NotImplementedError()

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError()

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError()

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each body paragraph."""
        raise NotImplementedError()


class MarkdownParagraph(IDocumentParagraph):
    """A Paragraph inside a document."""

    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        raise NotImplementedError()

    def text(self):
        """Yields strings of plain text."""
        raise NotImplementedError()

    def spans(self):
        """Yield a MarkdownSpan per text span."""
        raise NotImplementedError()


class MarkdownSpan(IDocumentSpan):
    """A span of characters inside a document."""
    def style_wpid(self):
        """Returns the wpid for this span's style."""
        raise NotImplementedError()

    def footnotes(self):
        """Yields a MarkdownFootnote object for each footnote in this span."""
        raise NotImplementedError()

    def comments(self):
        """Yields a MarkdownComment object for each comment in this span."""
        if False:
            yield

    def text(self):
        """Yields strings of plain text."""
        raise NotImplementedError()


class MarkdownFootnote(IDocumentFootnote):
    """A footnote."""

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each footnote paragraph."""
        raise NotImplementedError()
