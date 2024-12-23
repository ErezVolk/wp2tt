"""Utility classes."""
import argparse
import contextlib
import logging
import typing as t
from collections.abc import Sequence
from pathlib import Path

from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentTable
from wp2tt.docx import DocxInput
from wp2tt.markdown import MarkdownInput
from wp2tt.spreadsheet import CsvInput
from wp2tt.spreadsheet import OdsInput
from wp2tt.odt import XodtInput
from wp2tt.styles import DocumentProperties


class ProxyInput(IDocumentInput, contextlib.ExitStack):
    """Just a proxy IDocumentInput."""

    args: argparse.Namespace | None

    def __init__(self, args: argparse.Namespace | None = None) -> None:
        super().__init__()
        self.args = args


class MultiInput(ProxyInput):
    """Input from multiple files."""

    _args: argparse.Namespace | None

    def __init__(
        self,
        paths: Sequence[Path],
        args: argparse.Namespace | None = None,
    ) -> None:
        super().__init__(args)
        self._paths = paths
        self._inputs: list[IDocumentInput] = []
        for nth, path in enumerate(paths, 1):
            one = ByExtensionInput(path, self.args)
            if args and args.per_file_styles:
                one.set_nth(nth)
            self._inputs.append(one)
            self.enter_context(one)

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return DocumentProperties()  # Worst case scenario

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        known = set()
        for doc in self._inputs:
            for style in doc.styles_defined():
                if str(style) not in known:
                    yield style
                    known.add(str(style))

    def styles_in_use(self) -> t.Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        total = 0
        for path, doc in zip(self._paths, self._inputs, strict=True):
            in_file = 0
            for style in doc.styles_in_use():
                yield style
                in_file += 1
                total += 1
            logging.debug("%u style(s) in %r", in_file, path)
        logging.debug("%u style(s) in %u docs", total, len(self._paths))

    def paragraphs(self) -> t.Iterable[IDocumentParagraph | IDocumentTable]:
        """Yield an IDocumentParagraph object for each body paragraph."""
        total = 0
        for path, doc in zip(self._paths, self._inputs, strict=True):
            in_file = 0
            logging.debug("Paragraphs in %r", path)
            for para in doc.paragraphs():
                yield para
                in_file += 1
                total += 1
            logging.debug("%u paragraphs(s) in %r", in_file, path)
        logging.debug("%u paragraphs(s) in %u docs", total, len(self._paths))


class ByExtensionInput(ProxyInput):
    """An input, based on the file's extension."""

    _input: IDocumentInput

    def __init__(self, path: Path, args: argparse.Namespace | None = None) -> None:
        super().__init__(args)
        ext = path.suffix.lower()
        if ext == ".docx":
            self._input = DocxInput(path)
        elif ext == ".odt":
            self._input = XodtInput(path, zipped=True)
        elif ext == ".fodt":
            self._input = XodtInput(path, zipped=False)
        elif ext == ".md":
            self._input = MarkdownInput(path)
        elif ext in ".ods":
            self._input = OdsInput(path, args)
        elif ext in ".csv":
            self._input = CsvInput(path, args)
        else:
            raise RuntimeError(f"Unknown file extension for {path}")
        self.enter_context(self._input)

    def set_nth(self, nth: int) -> None:
        """Make sure style names in this document are prefixed."""
        self._input.set_nth(nth)

    @property
    def properties(self) -> DocumentProperties:
        """A DocumentProperties object."""
        return self._input.properties

    def styles_defined(self) -> t.Iterable[dict[str, str]]:
        """Yield a Style object kwargs for every style defined in the document."""
        yield from self._input.styles_defined()

    def styles_in_use(self) -> t.Iterable[tuple[str, str | None]]:
        """Yield a pair (realm, wpid) for every style used in the document."""
        yield from self._input.styles_in_use()

    def paragraphs(self) -> t.Iterable["IDocumentParagraph | IDocumentTable"]:
        """Yield an IDocumentParagraph object for each body paragraph."""
        yield from self._input.paragraphs()
