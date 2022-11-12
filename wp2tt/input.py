#!/usr/bin/env python
"""Base class for input sources (document formats)."""
from abc import abstractmethod
from abc import ABC

from typing import Generator

from wp2tt.format import ManualFormat
from wp2tt.styles import DocumentProperties


class IDocumentInput(ABC):
    """A document."""

    @property
    @abstractmethod
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        raise NotImplementedError()

    def styles_defined(self) -> Generator[dict[str, str], None, None]:
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError()

    def styles_in_use(self) -> Generator[tuple[str, str], None, None]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError()

    def paragraphs(self) -> Generator["IDocumentParagraph", None, None]:
        """Yields an IDocumentParagraph object for each body paragraph."""
        raise NotImplementedError()


class IDocumentParagraph(ABC):
    """A Paragraph inside a document."""

    @abstractmethod
    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    def is_page_break(self) -> bool:
        """True iff the paragraph is a page break."""
        return False

    @abstractmethod
    def text(self) -> Generator[str, None, None]:
        """Yields strings of plain text."""
        raise NotImplementedError()

    @abstractmethod
    def spans(self) -> Generator["IDocumentSpan", None, None]:
        """Yield an IDocumentSpan per text span."""
        raise NotImplementedError()


class IDocumentSpan(ABC):
    """A span of characters inside a document."""

    def style_wpid(self) -> str | None:
        """Returns the wpid for this span's style."""
        return None

    def footnotes(self) -> Generator["IDocumentFootnote", None, None]:
        """Yields an IDocumentFootnote object for each footnote in this span."""
        raise NotImplementedError()

    def comments(self) -> Generator["IDocumentComment", None, None]:
        """Yields an IDocumentComment object for each comment in this span."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this span."""
        return ManualFormat.NORMAL

    @abstractmethod
    def text(self) -> Generator[str, None, None]:
        """Yields strings of plain text."""
        raise NotImplementedError()


class IDocumentFootnote(ABC):
    """A footnote."""

    def paragraphs(self) -> Generator[IDocumentParagraph, None, None]:
        """Yields an IDocumentParagraph object for each footnote paragraph."""
        raise NotImplementedError()


class IDocumentComment(ABC):
    """A comment (balloon)."""

    def paragraphs(self) -> Generator[IDocumentParagraph, None, None]:
        """Yields an IDocumentParagraph object for each comment paragraph."""
        raise NotImplementedError()
