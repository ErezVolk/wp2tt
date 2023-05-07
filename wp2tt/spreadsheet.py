"""Simple spreadsheet reader (via pandas)"""
from abc import abstractmethod
import argparse
import contextlib
import logging
from pathlib import Path
import re
from typing import Iterable

from pandas_ods_reader import read_ods
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


class Wpids:
    TABLE_STYLE = "Spreadsheet"
    HEADER_STYLE = "Spreadsheet Header"
    BODY_STYLE = "Spreadsheet Body"

    @classmethod
    def rtl(cls, style):
        return f"{style} (RTL)"

    @classmethod
    def number(cls, style):
        return f"{style} (Number)"

    @classmethod
    def column(cls, name):
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
        cols = self._frame.columns
        if args.max_table_cols:
            cols = cols[:args.max_table_cols]
        elif args.table_cols:
            logging.debug(
                "Table column indexes: %s",
                " ".join(repr(col) for col in args.table_cols)
            )
            cols = [cols[idx - 1] for idx in args.table_cols]
        else:
            return
        logging.debug("Using columns: %s", ", ".join(cols))
        self._frame = self._frame[cols]

    @abstractmethod
    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        """Read the file, according to the subclass"""

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return self._props

    def styles_defined(self) -> Iterable[dict[str, str]]:
        """Styles defined"""
        yield {"realm": "table", "wpid": Wpids.TABLE_STYLE, "internal_name": Wpids.TABLE_STYLE}
        for name in [Wpids.HEADER_STYLE, Wpids.BODY_STYLE]:
            yield from self._para_styles(name)

    def _para_styles(self, name: str, parent: str | None = None) -> Iterable[dict[str, str]]:
        """Helper for `self.styles_defined()`"""
        yield {"realm": "paragraph", "wpid": name, "internal_name": name, "parent_wpid": parent}
        yield {"realm": "paragraph", "wpid": Wpids.rtl(name), "internal_name": Wpids.rtl(name), "parent_wpid": name}
        yield {"realm": "paragraph", "wpid": Wpids.number(name), "internal_name": Wpids.number(name), "parent_wpid": name}

    def styles_in_use(self) -> Iterable[tuple[str, str]]:
        """Basic styles"""
        for style_dict in self.styles_defined():
            yield {style_dict["realm"], style_dict["wpid"]}

    def paragraphs(self) -> Iterable["DataFrameTable"]:
        """Just the one table, for now."""
        yield DataFrameTable(self._frame, self._column_wpids)


class OdsInput(_SpreadsheetInput):
    """ODS reader"""

    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        return read_ods(path)


class CsvInput(_SpreadsheetInput):
    """CSV reader"""
    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        frame = pd.read_csv(path)
        self._column_wpids = [Wpids.column(col) for col in frame.columns]
        return frame

    def styles_defined(self) -> Iterable[dict[str, str]]:
        """CSV files get per-columns styles"""
        yield from super().styles_defined()
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

    def rows(self) -> Iterable[IDocumentTableRow]:
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

    def cells(self) -> Iterable[IDocumentTableCell]:
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

    def text(self) -> Iterable[str]:
        yield self._contents

    def chunks(self) -> Iterable[IDocumentSpan]:
        yield SimpleSpan(self._contents)


class SimpleSpan(IDocumentSpan):
    """A span that only holds a single piece of text"""
    def __init__(self, contents):
        self._contents = contents

    def text(self) -> Iterable[str]:
        yield self._contents

    def footnotes(self) -> Iterable[IDocumentFootnote]:
        yield from ()

    def comments(self) -> Iterable[IDocumentComment]:
        yield from ()
