"""Simple output classes."""
import re

from abc import ABC
from abc import abstractmethod

from wp2tt.styles import Style
from wp2tt.styles import OptionalStyle


class IOutput(ABC):
    """Interface for things that write InDesign Tagged Text."""

    @abstractmethod
    def define_style(self, style: Style) -> None:
        """Add a style definition."""
        raise NotImplementedError

    @abstractmethod
    def define_text_variable(self, name: str, value: str) -> None:
        """Add a text variable."""
        raise NotImplementedError

    @abstractmethod
    def enter_paragraph(self, style: OptionalStyle = None) -> None:
        """Start a paragraph with a specified style."""
        raise NotImplementedError

    @abstractmethod
    def leave_paragraph(self) -> None:
        """Finalize paragraph."""
        raise NotImplementedError

    @abstractmethod
    def enter_table(
        self,
        rows: int,
        cols: int,
        header_rows: int = 0,
        style: OptionalStyle = None,
        *,
        rtl: bool = False,
    ) -> None:
        """Start a table with a specified style."""
        raise NotImplementedError

    @abstractmethod
    def leave_table(self) -> None:
        """Finalize table."""
        raise NotImplementedError

    @abstractmethod
    def enter_table_row(self) -> None:
        """Start a table row."""
        raise NotImplementedError

    @abstractmethod
    def leave_table_row(self) -> None:
        """Finalize table row."""
        raise NotImplementedError

    @abstractmethod
    def enter_table_cell(self, rows: int = 1, cols: int = 1) -> None:
        """Start a table cell."""
        raise NotImplementedError

    @abstractmethod
    def leave_table_cell(self) -> None:
        """Finalize table cell."""
        raise NotImplementedError

    @abstractmethod
    def set_character_style(self, style: OptionalStyle = None) -> None:
        """Start a span using a specific character style."""
        raise NotImplementedError

    @abstractmethod
    def enter_footnote(self) -> None:
        """Add a footnote reference and enter the footnote."""
        raise NotImplementedError

    @abstractmethod
    def leave_footnote(self) -> None:
        """Close footnote, go ack to main text."""

    @abstractmethod
    def write_text(self, text: str) -> None:
        """Add some plain text."""

    @abstractmethod
    def write_bookmark(self, name: str) -> None:
        """Write a bookmark."""

    @abstractmethod
    def finalize(self) -> None:
        """Write any footers and stuff."""


class WhitespaceStripper(IOutput):
    """A proxy IOutput which strips all initial and final whitespace.

    Good foor footnotes.
    """

    def __init__(self, writer: IOutput) -> None:
        self.writer = writer
        self.begun = False

    def define_style(self, style: Style) -> None:
        """Add a style definition."""
        return self.writer.define_style(style)

    def define_text_variable(self, name: str, value: str) -> None:
        """Add a text variable."""
        return self.writer.define_text_variable(name, value)

    def enter_paragraph(self, style: OptionalStyle = None) -> None:
        """Start a paragraph with a specified style."""
        return self.writer.enter_paragraph(style)

    def leave_paragraph(self) -> None:
        """Finalize paragraph."""
        return self.writer.leave_paragraph()

    def set_character_style(self, style: OptionalStyle = None) -> None:
        """Start a span using a specific character style."""
        return self.writer.set_character_style(style)

    def enter_footnote(self) -> None:
        """Add a footnote reference and enter the footnote."""
        return self.writer.enter_footnote()

    def leave_footnote(self) -> None:
        """Close footnote, go ack to main text."""
        return self.writer.leave_footnote()

    def write_text(self, text: str) -> None:
        """Add some plain text."""
        if not self.begun:
            # Trim initial whitespace
            text = re.sub(r"^\s+", r"", text)
            self.begun = bool(text)

        return self.writer.write_text(text)

    def write_bookmark(self, name: str) -> None:
        """Write a bookmark."""
        return self.writer.write_bookmark(name)

    def enter_table(
        self,
        rows: int,
        cols: int,
        header_rows: int = 0,
        style: OptionalStyle = None,
        *,
        rtl: bool = False,
    ) -> None:
        """Start a table with a specified style."""
        return self.writer.enter_table(rows, cols, header_rows, style, rtl=rtl)

    def leave_table(self) -> None:
        """Finalize table."""
        return self.writer.leave_table()

    def enter_table_row(self) -> None:
        """Start a table row."""
        return self.writer.enter_table_row()

    def leave_table_row(self) -> None:
        """Finalize table row."""
        return self.writer.leave_table_row()

    def enter_table_cell(self, rows: int = 1, cols: int = 1) -> None:
        """Start a table cell."""
        return self.writer.enter_table_cell(rows, cols)

    def leave_table_cell(self) -> None:
        """Finalize table cell."""
        return self.writer.leave_table_cell()
