"""Base class for input sources (document formats)."""
from abc import abstractmethod
from abc import ABC
from os import PathLike
import typing as t

from wp2tt.format import ManualFormat
from wp2tt.styles import DocumentProperties


class IDocInput(ABC):
    """A document."""

    def set_nth(self, nth: int) -> None:
        """When supported, make all IDs globally unique."""
        if nth > 1:
            print(f"{type(self).__name__}.set_nth(): Not implemented")

    @property
    @abstractmethod
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        raise NotImplementedError

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        raise NotImplementedError

    def styles_in_use(self) -> t.Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        raise NotImplementedError

    def paragraphs(self) -> t.Iterable["IDocParagraph | IDocTable"]:
        """Yield an IDocParagraph object for each body paragraph."""
        raise NotImplementedError


class IDocTextSource(ABC):
    """Something that may contain text."""

    __is_empty: bool

    @abstractmethod
    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        raise NotImplementedError

    def is_empty(self) -> bool:
        """Return True iff the paragraph has no text."""
        try:
            return self.__is_empty
        except AttributeError:
            self.__is_empty = all(text == "" for text in self.text())
            return self.__is_empty


class IDocParagraph(IDocTextSource):
    """A Paragraph inside a document."""

    Chunk: t.TypeAlias = (
        "IDocSpan | IDocImage | IDocFormula | IDocBookmark | IDocHyperlink"
    )

    @abstractmethod
    def style_wpid(self) -> str | None:
        """Return the wpid for this paragraph's style."""
        raise NotImplementedError

    def format(self) -> ManualFormat:
        """Return manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    def is_page_break(self) -> bool:
        """Return True iff the paragraph is a page break."""
        return False

    @abstractmethod
    def chunks(self) -> t.Iterable[Chunk]:
        """Yield all elements in the paragraph."""
        raise NotImplementedError

    def spans(self) -> t.Iterable["IDocSpan"]:
        """Yield an IDocSpan per text span."""
        for chunk in self.chunks():
            if isinstance(chunk, IDocSpan):
                yield chunk


class IDocSpan(IDocTextSource):
    """A span of characters inside a document."""

    def style_wpid(self) -> str | None:
        """Return the wpid for this span's style."""
        return None

    def footnotes(self) -> t.Iterable["IDocFootnote"]:
        """Yield an IDocFootnote object for each footnote in this span."""
        raise NotImplementedError

    def comments(self) -> t.Iterable["IDocComment"]:
        """Yield an IDocComment object for each comment in this span."""
        raise NotImplementedError

    def format(self) -> ManualFormat:
        """Return manual formatting on this span."""
        return ManualFormat.NORMAL


class IDocImage(ABC):
    """An image inside a document."""

    def alt_text(self) -> str | None:
        """Get alternative text, if it exists."""
        return None

    @abstractmethod
    def suffix(self) -> str:
        """Extension (e.g., ".jpeg")."""
        raise NotImplementedError

    @abstractmethod
    def save(self, path: PathLike) -> None:
        """Write to a file."""
        raise NotImplementedError


class IDocFormula(ABC):
    """A formula inside a document."""

    @abstractmethod
    def raw(self) -> bytes:
        """Return the original formula."""
        raise NotImplementedError

    @abstractmethod
    def mathml(self) -> str:
        """Return formula as MathML."""
        raise NotImplementedError


class IDocTable(ABC):
    """A table inside a document."""

    def style_wpid(self) -> str | None:
        """Return the wpid for this table's style."""
        return None

    def format(self) -> ManualFormat:
        """Return manual formatting on this paragraph."""
        return ManualFormat.NORMAL

    @property
    @abstractmethod
    def shape(self) -> tuple[int, int]:
        """Pandas-like (number of rows, number of columns)."""

    @property
    def header_rows(self) -> int:
        """Number of header rows."""
        return 0

    @abstractmethod
    def rows(self) -> t.Iterable["IDocTableRow"]:
        """Iterate the rows of the table."""
        raise NotImplementedError


class IDocTableRow(ABC):
    """A row in a table inside a document."""

    @abstractmethod
    def cells(self) -> t.Iterable["IDocTableCell"]:
        """Iterate the cells in the row."""
        raise NotImplementedError


class IDocTableCell(ABC):
    """A cell in a table inside a document."""

    @property
    def shape(self) -> tuple[int, int]:
        """Number of rows and columns this cell spans."""
        return (1, 1)

    @abstractmethod
    def contents(self) -> IDocParagraph:
        """Get the contents of this cell."""


class IDocBookmark(ABC):
    """A bookmark in a document."""

    @property
    @abstractmethod
    def name(self) -> str:
        """The name of this bookmark."""


class IDocHyperlink(ABC):
    """A hyperlink in a document."""

    @property
    @abstractmethod
    def span(self) -> IDocSpan:
        """The apparance of this hyperlink."""

    @property
    @abstractmethod
    def target(self) -> str:
        """The target of this hyperlink."""

    def text(self) -> t.Iterable[str]:
        """Yield strings of plain text."""
        yield from self.span.text()


class IDocFootnote(ABC):
    """A footnote."""

    @abstractmethod
    def paragraphs(self) -> t.Iterable[IDocParagraph]:
        """Yield an IDocParagraph object for each footnote paragraph."""
        raise NotImplementedError


class IDocComment(ABC):
    """A comment (balloon)."""

    @abstractmethod
    def paragraphs(self) -> t.Iterable[IDocParagraph]:
        """Yield an IDocParagraph object for each comment paragraph."""
        raise NotImplementedError
