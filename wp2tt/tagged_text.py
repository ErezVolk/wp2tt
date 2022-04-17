#!/usr/bin/env python3
"""Tagged text creation"""
import collections
import contextlib
import itertools
import logging
from pathlib import Path
import re

from typing import Iterator
from typing import List
from typing import Mapping
from typing import Optional

from wp2tt.output import IOutput
from wp2tt.styles import DocumentProperties
from wp2tt.styles import Style


class InDesignTaggedTextOutput(IOutput, contextlib.ExitStack):
    """Writes to a tagged text file"""

    def __init__(
        self,
        filename: Path,
        debug=False,
        properties: Optional[DocumentProperties] = None,
    ):
        super().__init__()
        self._filename = filename
        self._styles: List[Style] = []
        self._headers_written = False
        self._shades: Mapping[str, Iterator[int]] = collections.defaultdict(itertools.count)
        self._debug = debug
        if properties is None:
            self._properties = DocumentProperties()
        else:
            self._properties = properties

        self._filename.parent.mkdir(parents=True, exist_ok=True)
        self._fo = self.enter_context(open(self._filename, "w", encoding="UTF-16LE"))
        ufo_fn = self._filename.with_suffix(".utf8")
        if self._debug:
            self._ufo = self.enter_context(open(ufo_fn, "w", encoding="UTF-8"))
        elif ufo_fn.is_file():
            ufo_fn.unlink()

    def _writeln(self, line="") -> None:
        self._write(line)
        self._write("\n")

    def define_style(self, style: Style) -> None:
        """Add a style definition."""
        if style in self._styles:
            return

        if style.parent_style and style.parent_style.used:
            self.define_style(style.parent_style)

        self._styles.append(style)

        if style.next_style and style.next_style.used:
            self.define_style(style.next_style)
        if self._headers_written:
            self._write_style_definition(style)

    def _write_headers(self) -> None:
        if self._headers_written:
            return

        self._writeln("<UNICODE-MAC>")
        self._write(r"<Version:13.1>")
        if self._properties.has_rtl:
            self._write(r"<FeatureSet:Indesign-R2L>")
        self._write(r"<ColorTable:=")
        self._write(r"<Black:COLOR:CMYK:Process:0,0,0,1>")
        self._write(r"<Cyan:COLOR:CMYK:Process:1,0,0,0>")
        self._write(r"<Magenta:COLOR:CMYK:Process:0,1,0,0>")
        self._write(r"<Yellow:COLOR:CMYK:Process:0,0,1,0>")
        self._write(r">")
        self._writeln()
        for style in self._styles:
            self._write_style_definition(style)
            self._writeln()
        self._headers_written = True

    def _write_style_definition(self, style: Style) -> None:
        logging.debug("InDesign: %s", style)
        id_realm = style.realm[:4].capitalize()
        if style.idtt:
            idtt = [style.idtt]
        else:
            fullness = 50.0 + 50.0 / (1.05 ** next(self._shades[style.realm]))

            if style.realm == "paragraph":
                idtt = [
                    "<pShadingColor:Yellow>",
                    "<pShadingOn:1>",
                    "<pShadingTint:",
                    str(int(100 - fullness)),
                    ">",
                ]
            elif style.realm == "character":
                idtt = ["<cColor:Magenta>", "<cColorTint:", str(int(fullness)), ">"]

        self._write("<Define")
        self._write(id_realm)
        self._write("Style:")
        self._write(self._idname(style))
        for elem in idtt:
            self._write(elem)

        if style.parent_style and style.parent_style.used:
            self._write("<BasedOn:")
            self._write(self._idname(style.parent_style))
            self._write(">")
        if style.next_style and style.next_style.used:
            self._write("<Nextstyle:")
            self._write(self._idname(style.next_style))
            self._write(">")
        self._write(">")

    def define_text_variable(self, name: str, value: str) -> None:
        self._write("<DefineTextVariable:")
        self._write_escaped(name)
        self._write("=<TextVarType:CustomText>")
        self._write("<tvString:")
        self._write_escaped(value)
        self._write(">")
        self._write(">")

    def enter_paragraph(self, style: Optional[Style] = None) -> None:
        """Start a paragraph with a specified style."""
        self._write_headers()
        self._set_style("Para", style)

    def _set_style(self, realm: str, style: Optional[Style]) -> None:
        if style:
            self.define_style(style)
        self._write("<")
        self._write(realm)
        self._write("Style:")
        if style:
            self._write(self._idname(style))
        self._write(">")

    @classmethod
    def _idname(cls, style: Style) -> str:
        return re.sub(r"\s*/\s*", r"\:", style.name)

    def leave_paragraph(self) -> None:
        """Finalize paragraph."""
        self._writeln()

    def set_character_style(self, style: Optional[Style] = None) -> None:
        """Start a span using a specific character style."""
        self._set_style("Char", style)

    def enter_footnote(self) -> None:
        """Add a footnote reference and enter the footnote."""
        self._write("<FootnoteStart:>")

    def leave_footnote(self) -> None:
        """Close footnote, go ack to main text."""
        self._write("<FootnoteEnd:>")

    def write_text(self, text: str) -> None:
        """Add some plain text."""
        if text:
            self._write_escaped(text)

    def _write_escaped(self, text: str) -> None:
        return self._write(self._escape(text))

    @classmethod
    def _escape(cls, text: str) -> str:
        return re.sub(r"([<>])", r"\\\1", text)

    def _write(self, string: str) -> None:
        if string:
            self._fo.write(string)
            if self._debug:
                self._ufo.write(string)
