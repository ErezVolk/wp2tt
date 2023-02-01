"""Simple ODS reader (via pandas)"""
import argparse
import contextlib
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


class OdsInput(contextlib.ExitStack, IDocumentInput):
    """Simple ODS reader (via pandas)"""
    _frame: pd.DataFrame
    _props: DocumentProperties = DocumentProperties()

    def __init__(self, path: Path, args: argparse.Namespace | None = None):
        super().__init__()
        self._frame = read_ods(path)
        if args and args.max_table_cols:
            self._frame = self._frame[self._frame.columns[:args.max_table_cols]]

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
        ]

    def paragraphs(self) -> Iterable["DataFrameTable"]:
        """Just the one table, for now."""
        yield DataFrameTable(self._frame)


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
            try:
                if item.is_integer():
                    item = int(item)
            except AttributeError:
                pass
            yield SimpleCell(str(item), self._wpid)


class SimpleCell(IDocumentTableCell):
    """A cell that only holds a single piece of text"""
    def __init__(self, contents, wpid):
        self._contents = contents
        if re.search(r"[\u0591-\u05F4\u0600-\u06FF]", contents):
            self._wpid = Wpids.rtl(wpid)
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
