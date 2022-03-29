#!/usr/bin/env python
"""Base class for input sources (document formats)."""
import enum


class ManualFormat(enum.Flag):
    """Manual Formatting"""

    NORMAL = 0

    CENTERED = enum.auto()
    JUSTIFIED = enum.auto()
    NEW_PAGE = enum.auto()
    SPACED = enum.auto()

    BOLD = enum.auto()
    ITALIC = enum.auto()


class IDocumentInput(object):
    """A document."""

    @property
    def properties(self):
        """A DocumentProperties object."""
        raise NotImplementedError()

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError()

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError()

    def paragraphs(self):
        """Yields an IDocumentParagraph object for each body paragraph."""
        raise NotImplementedError()


class IDocumentParagraph(object):
    """A Paragraph inside a document."""

    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    def is_page_break(self):
        """True iff the paragraph is a page break."""
        return False

    def text(self):
        """Yields strings of plain text."""
        raise NotImplementedError()

    def spans(self):
        """Yield an IDocumentSpan per text span."""
        raise NotImplementedError()


class IDocumentSpan(object):
    """A span of characters inside a document."""

    def style_wpid(self):
        """Returns the wpid for this span's style."""
        raise NotImplementedError()

    def footnotes(self):
        """Yields an IDocumentFootnote object for each footnote in this span."""
        raise NotImplementedError()

    def comments(self):
        """Yields an IDocumentComment object for each comment in this span."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this span."""
        return ManualFormat.NORMAL

    def text(self):
        """Yields strings of plain text."""
        raise NotImplementedError()


class IDocumentFootnote(object):
    """A footnote."""

    def paragraphs(self):
        """Yields an IDocumentParagraph object for each footnote paragraph."""
        raise NotImplementedError()


class IDocumentComment(object):
    """A comment (balloon)."""

    def paragraphs(self):
        """Yields an IDocumentParagraph object for each comment paragraph."""
        raise NotImplementedError()
