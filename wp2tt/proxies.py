#!/usr/bin/env python3
"""Utility classes"""
from collections.abc import Sequence
import contextlib
import logging
from pathlib import Path

from typing import Generator
from typing import List

from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.docx import DocxInput
from wp2tt.markdown import MarkdownInput
from wp2tt.odt import XodtInput
from wp2tt.styles import DocumentProperties


class MultiInput(IDocumentInput, contextlib.ExitStack):
    """Input from multiple files."""

    def __init__(self, paths: Sequence[Path]):
        super().__init__()
        self._paths = paths
        self._inputs: List[IDocumentInput] = [
            self.enter_context(ByExtensionInput(path))
            for path in paths
        ]

    @property
    def properties(self):
        """A DocumentProperties object."""
        return DocumentProperties()  # Worst case scenario

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        known = set()
        for doc in self._inputs:
            for style in doc.styles_defined():
                if str(style) not in known:
                    yield style
                    known.add(str(style))

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        total = 0
        for path, doc in zip(self._paths, self._inputs):
            in_file = 0
            for style in doc.styles_in_use():
                yield style
                in_file += 1
                total += 1
            logging.debug("%u style(s) in %r", in_file, path)
        logging.debug("%u style(s) in %u docs", total, len(self._paths))

    def paragraphs(self) -> Generator[IDocumentParagraph, None, None]:
        """Yields an IDocumentParagraph object for each body paragraph."""
        total = 0
        for path, doc in zip(self._paths, self._inputs):
            in_file = 0
            logging.debug("Paragraphs in %r", path)
            for para in doc.paragraphs():
                yield para
                in_file += 1
                total += 1
            logging.debug("%u paragraphs(s) in %r", in_file, path)
        logging.debug("%u paragraphs(s) in %u docs", total, len(self._paths))


class ByExtensionInput(IDocumentInput, contextlib.ExitStack):
    """An input, based on the file's extension."""

    _input: IDocumentInput

    def __init__(self, path: Path):
        super().__init__()
        ext = path.suffix.lower()
        if ext == ".docx":
            self._input = DocxInput(path)
        elif ext == ".odt":
            self._input = XodtInput(path, zipped=True)
        elif ext == ".fodt":
            self._input = XodtInput(path, zipped=False)
        elif ext == ".md":
            self._input = MarkdownInput(path)
        else:
            raise RuntimeError(f"Unknown file extension for {path}")
        self.enter_context(self._input)

    @property
    def properties(self):
        """A DocumentProperties object."""
        return self._input.properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        yield from self._input.styles_defined()

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        yield from self._input.styles_in_use()

    def paragraphs(self):
        """Yields an IDocumentParagraph object for each body paragraph."""
        yield from self._input.paragraphs()
