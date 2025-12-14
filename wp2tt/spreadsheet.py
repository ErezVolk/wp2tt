"""Simple spreadsheet reader (via pandas)."""
from abc import abstractmethod
import argparse
import contextlib
import logging
from pathlib import Path
import re
import typing as t

import pandas as pd

from wp2tt.input import IDocumentComment
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.input import IDocumentTable
from wp2tt.input import IDocumentTableCell
from wp2tt.input import IDocumentTableRow
from wp2tt.styles import DocumentProperties

log = logging.getLogger(__name__)


class Wpids:
    """Known paragraph IDs."""

    TABLE_STYLE = "Spreadsheet"
    HEADER_STYLE = "Spreadsheet Header"
    BODY_STYLE = "Spreadsheet Body"

    @classmethod
    def rtl(cls, style):
        """Nice name for RTL version of a style"""
        return f"{style} (RTL)"

    @classmethod
    def number(cls, style):
        """Nice name for Number version of a style"""
        return f"{style} (Number)"

    @classmethod
    def column(cls, name):
        """Nice name for column style"""
        return f"Spreadsheet Column ({name})"


class _SpreadsheetInput(contextlib.ExitStack, IDocumentInput):
    """Simple ODS reader (via pandas)"""
    _frame: pd.DataFrame
    _props: DocumentProperties = DocumentProperties()
    _column_wpids: list[str] | None = None

    def __init__(self, path: Path, args: argparse.Namespace | None = None):
        super().__init__()
        self._frame = self._read_spreadsheet(path)
        if not args:
            return
        cols = [str(col) for col in self._frame.columns]
        if args.max_table_cols:
            cols = cols[:args.max_table_cols]
        elif args.table_cols:
            log.debug("Table column indexes: %s", " ".join(args.table_cols))
            cols = [cols[idx - 1] for idx in args.table_cols]
        else:
            return
        log.debug("Using columns: %s", ", ".join(cols))
        self._frame = self._frame[cols]

    @abstractmethod
    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        """Read the file, according to the subclass"""

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return self._props

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """Styles defined"""
        yield {
            "realm": "table",
            "wpid": Wpids.TABLE_STYLE,
            "internal_name": Wpids.TABLE_STYLE,
        }
        for name in [Wpids.HEADER_STYLE, Wpids.BODY_STYLE]:
            yield from self._para_styles(name)

    def _para_styles(
        self,
        name: str,
        parent: str | None = None,
    ) -> t.Iterable[dict[str, str]]:
        """Helper for `self.styles_defined()`"""
        child = {"realm": "paragraph", "wpid": name, "internal_name": name}
        if parent is not None:
            child["parent_wpid"] = parent
        yield child
        for func in [Wpids.rtl, Wpids.number]:
            alt_name = func(name)
            yield dict(child) | {
                "wpid": alt_name,
                "internal_name": alt_name,
                "parent_wpid": name,
            }

    def styles_in_use(self) -> t.Iterable[tuple[str, str]]:
        """Basic styles"""
        for style_dict in self.styles_defined():
            yield (style_dict["realm"], style_dict["wpid"])

    def paragraphs(self) -> t.Iterable["DataFrameTable"]:
        """Just the one table, for now."""
        yield DataFrameTable(self._frame, self._column_wpids)


class OdsInput(_SpreadsheetInput):
    """ODS reader"""

    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        return pd.read_excel(path)


class CsvInput(_SpreadsheetInput):
    """CSV reader"""
    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        frame = pd.read_csv(path).fillna("")
        self._column_wpids = [Wpids.column(col) for col in frame.columns]
        return frame

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """CSV files get per-columns styles"""
        yield from super().styles_defined()
        if self._column_wpids is not None:
            for wpid in self._column_wpids:
                yield from self._para_styles(wpid, parent=Wpids.BODY_STYLE)


class DataFrameTable(IDocumentTable):
    """A DataFrame as a document table"""
    def __init__(self, frame: pd.DataFrame, column_wpids: list[str] | None):
        self._frame = frame
        self._column_wpids = column_wpids

    def style_wpid(self) -> str | None:
        return Wpids.TABLE_STYLE

    @property
    def shape(self) -> tuple[int, int]:
        """(number of rows, number of columns)"""
        return (len(self._frame) + 1, len(self._frame.columns))

    @property
    def header_rows(self) -> int:
        return 1

    def rows(self) -> t.Iterable[IDocumentTableRow]:
        """Iterates the rows of the table"""
        yield DataFrameRow(self._frame.columns, Wpids.HEADER_STYLE)
        for _, row in self._frame.iterrows():
            yield DataFrameRow(row, self._column_wpids)


class DataFrameRow(IDocumentTableRow):
    """A row in a table inside a document"""
    def __init__(self, items, wpids: list[str] | str | None = None):
        self._items = items
        if wpids is None:
            self._wpids = [Wpids.BODY_STYLE] * len(self._items)
        elif isinstance(wpids, str):
            self._wpids = [wpids] * len(self._items)
        else:
            self._wpids = wpids

    def cells(self) -> t.Iterable[IDocumentTableCell]:
        """Iterates the cells in the row"""
        for item, wpid in zip(self._items, self._wpids):
            if item is None:
                item = ""
            elif pd.api.types.is_number(item):
                if not isinstance(item, int) and item.is_integer():
                    item = int(item)
            yield SimpleCell(str(item), wpid)


class SimpleCell(IDocumentTableCell):
    """A cell that only holds a single piece of text"""
    def __init__(self, contents, wpid):
        self._contents = contents.strip()
        if re.search(r"[\u0591-\u05F4\u0600-\u06FF]", contents):
            self._wpid = Wpids.rtl(wpid)
        elif re.fullmatch(r"\d*\.?\d+", contents):
            self._wpid = Wpids.number(wpid)
        else:
            self._wpid = wpid

    def contents(self) -> IDocumentParagraph:
        """Get the contents of this cell"""
        return SimpleParagraph(self._contents, self._wpid)


class SimpleParagraph(IDocumentParagraph):
    """A paragraph that only holds a single piece of text"""
    def __init__(self, contents, wpid):
        self._contents = contents
        self._wpid = wpid

    def style_wpid(self):
        return self._wpid

    def text(self) -> t.Iterable[str]:
        yield self._contents

    def chunks(self) -> t.Iterable[IDocumentSpan]:
        yield SimpleSpan(self._contents)


class SimpleSpan(IDocumentSpan):
    """A span that only holds a single piece of text"""
    def __init__(self, contents):
        self._contents = contents

    def text(self) -> t.Iterable[str]:
        yield self._contents

    def footnotes(self) -> t.Iterable[IDocumentFootnote]:
        yield from ()

    def comments(self) -> t.Iterable[IDocumentComment]:
        yield from ()
