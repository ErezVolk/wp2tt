#!/usr/bin/env python3
"""Tagged text creation"""
import collections
import contextlib
import io
import itertools
import logging
import re

from typing import Iterator
from typing import Mapping

from wp2tt.output import IOutput
from wp2tt.styles import DocumentProperties
from wp2tt.styles import Style
from wp2tt.styles import OptionalStyle


class InDesignTaggedTextOutput(IOutput, contextlib.ExitStack):
    """Writes to a tagged text file"""

    BASIC_TABLE_STYLE = r"\[Basic Table\]"
    REALM_TO_MNEM = {
        "character": "Char",
        "paragraph": "Para",
        "table": "Table",
    }

    _curr_char_style: OptionalStyle = None
    _in_table: bool = False

    def __init__(self, properties: DocumentProperties | None = None):
        super().__init__()
        self._buffer = io.StringIO()
        self._styles: list[Style] = []
        self._headers_written = False
        self._shades: Mapping[str, Iterator[int]] = collections.defaultdict(itertools.count)
        if properties is None:
            self._properties = DocumentProperties()
        else:
            self._properties = properties

    def _writeln(self, line="") -> None:
        self._write(line)
        self._write("\n")

    def define_style(self, style: Style) -> None:
        """Add a style definition."""
        if style in self._styles:
            return

        if style.parent_style is not None:
            if style.parent_style.used:
                self.define_style(style.parent_style)

        self._styles.append(style)

        if style.next_style is not None:
            if style.next_style.used:
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
        idtt: list[str] = []
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
        self._write(self.REALM_TO_MNEM[style.realm])
        self._write("Style:")
        self._write(self._idname(style))
        for elem in idtt:
            self._write(elem)

        if style.parent_style is not None and style.parent_style.used:
            self._write("<BasedOn:")
            self._write(self._idname(style.parent_style))
            self._write(">")
        if style.next_style is not None and style.next_style.used:
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

    def enter_table(self, rows: int, cols: int, rtl: bool = False, style: OptionalStyle = None):
        """Start a table"""
        self._set_style("Table", style)
        direction = "RTL" if rtl else "LTR"
        self._write(f"<TableStart:{rows},{cols}:0:0:{direction}>")
        self._in_table = True

    def leave_table(self, style: OptionalStyle = None) -> None:
        """Finish a table"""
        self._write("<TableEnd:>")
        self._in_table = False

    def enter_table_row(self):
        """Start a table row."""
        self._write("<RowStart:>")

    def leave_table_row(self) -> None:
        """Finalize table row."""
        self._write("<RowEnd:>")

    def enter_table_cell(self):
        """Start a table cell."""
        self._write("<CellStart:>")

    def leave_table_cell(self) -> None:
        """Finalize table cell."""
        self._write("<CellEnd:>")

    def enter_paragraph(self, style: OptionalStyle = None) -> None:
        """Start a paragraph with a specified style."""
        self._write_headers()
        self._set_style("Para", style)
        if self._curr_char_style:
            self._set_style("Char", self._curr_char_style)

    def _set_style(self, mnem: str, style: OptionalStyle) -> None:
        if style:
            self.define_style(style)
        self._write("<")
        self._write(mnem)
        self._write("Style:")
        if style:
            self._write(self._idname(style))
        self._write(">")

    @classmethod
    def _idname(cls, style: Style) -> str:
        return re.sub(r"\s*/\s*", r"\:", style.name)

    def leave_paragraph(self) -> None:
        """Finalize paragraph."""
        if self._curr_char_style:
            self._set_style("Char", None)
        if not self._in_table:
            self._writeln()

    def set_character_style(self, style: OptionalStyle = None) -> None:
        """Start a span using a specific character style."""
        self._set_style("Char", style)
        self._curr_char_style = style

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
            self._buffer.write(string)

    @property
    def contents(self) -> str:
        """The actual tagged text"""
        return self._buffer.getvalue()
