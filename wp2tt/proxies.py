#!/usr/bin/env python3
import contextlib
import os
from wp2tt.input import IDocumentInput
from wp2tt.docx import DocxInput
from wp2tt.markdown import MarkdownInput
from wp2tt.odt import XodtInput
from wp2tt.styles import DocumentProperties


class MultiInput(contextlib.ExitStack, IDocumentInput):
    """Input from multiple files."""

    def __init__(self, paths):
        super().__init__()
        self._inputs = [
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
        for doc in self._inputs:
            yield from doc.styles_in_use()

    def paragraphs(self):
        """Yields an IDocumentParagraph object for each body paragraph."""
        for doc in self._inputs:
            yield from doc.paragraphs()


class ByExtensionInput(contextlib.ExitStack, IDocumentInput):
    """An input, based on the file's extension."""

    def __init__(self, path):
        super().__init__()
        _, ext = os.path.splitext(path)
        ext = ext.lower()
        if ext == '.docx':
            self._input = DocxInput(path)
        elif ext == '.odt':
            self._input = XodtInput(path, zipped=True)
        elif ext == '.fodt':
            self._input = XodtInput(path, zipped=False)
        elif ext == '.md':
            self._input = MarkdownInput(path)
        else:
            raise RuntimeError('Unknown file extension for %r', path)
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
