"""Simple ODS reader (via pandas)"""
import argparse
import contextlib
from pathlib import Path
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


class OdsInput(contextlib.ExitStack, IDocumentInput):
    """Simple ODS reader (via pandas)"""
    frame: pd.DataFrame
    props: DocumentProperties = DocumentProperties()

    def __init__(self, path: Path, args: argparse.Namespace | None = None):
        super().__init__()
        self.frame = read_ods(path)
        if args and args.max_table_cols:
            self.frame = self.frame[self.frame.columns[:args.max_table_cols]]

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return self.props

    def styles_defined(self) -> Iterable[dict[str, str]]:
        """Not supported yet"""
        yield from ()

    def styles_in_use(self) -> Iterable[tuple[str, str]]:
        """Not supported yet"""
        yield from ()

    def paragraphs(self) -> Iterable["DataFrameTable"]:
        """Just the one table, for now."""
        yield DataFrameTable(self.frame)


class DataFrameTable(IDocumentTable):
    """A DataFrame as a document table"""
    def __init__(self, frame: pd.DataFrame):
        self.frame = frame

    @property
    def shape(self) -> tuple[int, int]:
        """(number of rows, number of columns)"""
        return (len(self.frame) + 1, len(self.frame.columns))

    @property
    def header_rows(self) -> int:
        return 1

    def rows(self) -> Iterable[IDocumentTableRow]:
        """Iterates the rows of the table"""
        yield DataFrameRow(self.frame.columns)
        for _, row in self.frame.iterrows():
            yield DataFrameRow(row)


class DataFrameRow(IDocumentTableRow):
    """A row in a table inside a document"""
    def __init__(self, items):
        self.items = items

    def cells(self) -> Iterable[IDocumentTableCell]:
        """Iterates the cells in the row"""
        for item in self.items:
            yield SimpleCell(str(item))


class SimpleCell(IDocumentTableCell):
    """A cell that only holds a single piece of text"""
    def __init__(self, contents):
        self._contents = contents

    def contents(self) -> IDocumentParagraph:
        """Get the contents of this cell"""
        return SimpleParagraph(self._contents)


class SimpleParagraph(IDocumentParagraph):
    """A paragraph that only holds a single piece of text"""
    def __init__(self, contents):
        self._contents = contents

    def style_wpid(self):
        return None

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
