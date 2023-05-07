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


class _SpreadsheetInput(contextlib.ExitStack, IDocumentInput):
    """Simple ODS reader (via pandas)"""
    _frame: pd.DataFrame
    _props: DocumentProperties = DocumentProperties()

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
        """Not supported yet"""
        return [
            {"realm": realm, "wpid": wpid, "internal_name": wpid}
            for realm, wpid in self.styles_in_use()
        ]

    def styles_in_use(self) -> Iterable[tuple[str, str]]:
        """Not supported yet"""
        return [
            ("table", Wpids.TABLE_STYLE),
            ("paragraph", Wpids.HEADER_STYLE),
            ("paragraph", Wpids.BODY_STYLE),
            ("paragraph", Wpids.rtl(Wpids.HEADER_STYLE)),
            ("paragraph", Wpids.rtl(Wpids.BODY_STYLE)),
            ("paragraph", Wpids.number(Wpids.HEADER_STYLE)),
            ("paragraph", Wpids.number(Wpids.BODY_STYLE)),
        ]

    def paragraphs(self) -> Iterable["DataFrameTable"]:
        """Just the one table, for now."""
        yield DataFrameTable(self._frame)


class OdsInput(_SpreadsheetInput):
    """ODS reader"""

    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        return read_ods(path)


class CsvInput(_SpreadsheetInput):
    """CSV reader"""
    def _read_spreadsheet(self, path: Path) -> pd.DataFrame:
        return pd.read_csv(path)


class DataFrameTable(IDocumentTable):
    """A DataFrame as a document table"""
    def __init__(self, frame: pd.DataFrame):
        self._frame = frame

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
            yield DataFrameRow(row)


class DataFrameRow(IDocumentTableRow):
    """A row in a table inside a document"""
    def __init__(self, items, wpid=None):
        self._wpid = wpid or Wpids.BODY_STYLE
        self._items = items

    def cells(self) -> Iterable[IDocumentTableCell]:
        """Iterates the cells in the row"""
        for item in self._items:
            if item is None:
                item = ""
            elif pd.api.types.is_number(item):
                if not isinstance(item, int) and item.is_integer():
                    item = int(item)
            yield SimpleCell(str(item), self._wpid)


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
