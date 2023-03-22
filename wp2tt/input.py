#!/usr/bin/env python
"""Base class for input sources (document formats)."""
from abc import abstractmethod
from abc import ABC
from os import PathLike

from typing import Iterable

from wp2tt.format import ManualFormat
from wp2tt.styles import DocumentProperties


class IDocumentInput(ABC):
    """A document."""

    def set_nth(self, nth: int):
        """When supported, make all IDs globally unique"""
        if nth > 1:
            print(f"{type(self).__name__}.set_nth(): Not implemented")

    @property
    @abstractmethod
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        raise NotImplementedError()

    def styles_defined(self) -> Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError()

    def styles_in_use(self) -> Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError()

    def paragraphs(self) -> Iterable["IDocumentParagraph | IDocumentTable"]:
        """Yields an IDocumentParagraph object for each body paragraph."""
        raise NotImplementedError()


class IDocumentParagraph(ABC):
    """A Paragraph inside a document."""

    @abstractmethod
    def style_wpid(self) -> str | None:
        """Returns the wpid for this paragraph's style."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    def is_page_break(self) -> bool:
        """True iff the paragraph is a page break."""
        return False

    @abstractmethod
    def text(self) -> Iterable[str]:
        """Yields strings of plain text."""
        raise NotImplementedError()

    @abstractmethod
    def chunks(self) -> Iterable["IDocumentSpan | IDocumentImage | IDocumentFormula"]:
        """Yield all elements in the paragraph."""
        raise NotImplementedError()

    def spans(self) -> Iterable["IDocumentSpan"]:
        """Yield an IDocumentSpan per text span."""
        for chunk in self.chunks():
            if isinstance(chunk, IDocumentSpan):
                yield chunk


class IDocumentSpan(ABC):
    """A span of characters inside a document."""

    def style_wpid(self) -> str | None:
        """Returns the wpid for this span's style."""
        return None

    def footnotes(self) -> Iterable["IDocumentFootnote"]:
        """Yields an IDocumentFootnote object for each footnote in this span."""
        raise NotImplementedError()

    def comments(self) -> Iterable["IDocumentComment"]:
        """Yields an IDocumentComment object for each comment in this span."""
        raise NotImplementedError()

    def format(self) -> ManualFormat:
        """Returns manual formatting on this span."""
        return ManualFormat.NORMAL

    @abstractmethod
    def text(self) -> Iterable[str]:
        """Yields strings of plain text."""
        raise NotImplementedError()


class IDocumentImage(ABC):
    """An image inside a document."""

    def alt_text(self) -> str | None:
        """Alternative text, if it exists"""
        return None

    @abstractmethod
    def suffix(self) -> str:
        """Extension (e.g., ".jpeg")."""
        raise NotImplementedError()

    @abstractmethod
    def save(self, path: PathLike) -> None:
        """Writes to a file."""
        raise NotImplementedError()


class IDocumentFormula(ABC):
    """A formula inside a document."""
    @abstractmethod
    def raw(self) -> bytes:
        """Return the original formula"""
        raise NotImplementedError()

    @abstractmethod
    def mathml(self) -> str:
        """Return formula as MathML."""
        raise NotImplementedError()


class IDocumentTable(ABC):
    """A table inside a document"""

    def style_wpid(self) -> str | None:
        """Returns the wpid for this table's style."""
        return None

    def format(self) -> ManualFormat:
        """Returns manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    @property
    @abstractmethod
    def shape(self) -> tuple[int, int]:
        """(number of rows, number of columns)"""

    @property
    def header_rows(self) -> int:
        """Number of header rows"""
        return 0

    @abstractmethod
    def rows(self) -> Iterable["IDocumentTableRow"]:
        """Iterates the rows of the table"""
        raise NotImplementedError()


class IDocumentTableRow(ABC):
    """A row in a table inside a document"""

    @abstractmethod
    def cells(self) -> Iterable["IDocumentTableCell"]:
        """Iterates the cells in the row"""
        raise NotImplementedError()


class IDocumentTableCell(ABC):
    """A cell in a table inside a document"""

    @property
    def shape(self) -> tuple[int, int]:
        """Number of rows and columns this cell spans"""
        return (1, 1)

    @abstractmethod
    def contents(self) -> IDocumentParagraph:
        """Get the contents of this cell"""


class IDocumentFootnote(ABC):
    """A footnote."""

    def paragraphs(self) -> Iterable[IDocumentParagraph]:
        """Yields an IDocumentParagraph object for each footnote paragraph."""
        raise NotImplementedError()


class IDocumentComment(ABC):
    """A comment (balloon)."""

    def paragraphs(self) -> Iterable[IDocumentParagraph]:
        """Yields an IDocumentParagraph object for each comment paragraph."""
        raise NotImplementedError()
