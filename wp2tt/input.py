#!/usr/bin/env python
"""Base class for input sources (document formats)."""
from abc import abstractmethod
from abc import abstractproperty
from abc import ABC
import enum

from typing import Generator
from typing import Tuple

from wp2tt.styles import DocumentProperties
from wp2tt.styles import Style


class ManualFormat(enum.Flag):
    """Manual Formatting"""

    NORMAL = 0

    CENTERED = enum.auto()
    JUSTIFIED = enum.auto()
    NEW_PAGE = enum.auto()
    SPACED = enum.auto()

    BOLD = enum.auto()
    ITALIC = enum.auto()


class IDocumentSpan(ABC):
    """A span of characters inside a document."""

    @abstractmethod
    def style_wpid(self) -> str:
        """Returns the wpid for this span's style."""
        raise NotImplementedError()

    @abstractmethod
    def footnotes(self):
        """Yields an IDocumentFootnote object for each footnote in this span."""
        raise NotImplementedError()

    @abstractmethod
    def comments(self):
        """Yields an IDocumentComment object for each comment in this span."""
        raise NotImplementedError()

    @abstractmethod
    def format(self) -> ManualFormat:
        """Returns manual formatting on this span."""
        return ManualFormat.NORMAL

    @abstractmethod
    def text(self) -> Generator[str, None, None]:
        """Yields strings of plain text."""
        raise NotImplementedError()


class IDocumentParagraph(ABC):
    """A Paragraph inside a document."""

    @abstractmethod
    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        raise NotImplementedError()

    @abstractmethod
    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    @abstractmethod
    def is_page_break(self) -> bool:
        """True iff the paragraph is a page break."""
        return False

    @abstractmethod
    def text(self) -> Generator[str, None, None]:
        """Yields strings of plain text."""
        raise NotImplementedError()

    @abstractmethod
    def spans(self) -> Generator[IDocumentSpan, None, None]:
        """Yield an IDocumentSpan per text span."""
        raise NotImplementedError()


class IDocumentInput(ABC):
    """A document."""

    @property
    @abstractproperty
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        raise NotImplementedError()

    def styles_defined(self) -> Generator[Style, None, None]:
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError()

    def styles_in_use(self) -> Generator[Tuple[str, str], None, None]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError()

    def paragraphs(self) -> Generator[IDocumentParagraph, None, None]:
        """Yields an IDocumentParagraph object for each body paragraph."""
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
