"""Simple output classes."""
import re
import typing as t

from abc import ABC
from abc import abstractmethod

if t.TYPE_CHECKING:
    from wp2tt.styles import Style
    from wp2tt.styles import OptionalStyle


class IOutput(ABC):
    """Interface for things that write InDesign Tagged Text."""

    @abstractmethod
    def define_style(self, style: Style) -> None:
        """Add a style definition."""
        del style
        raise NotImplementedError

    @abstractmethod
    def define_text_variable(self, name: str, value: str) -> None:
        """Add a text variable."""
        del name, value
        raise NotImplementedError

    @abstractmethod
    def enter_paragraph(self, style: OptionalStyle = None) -> None:
        """Start a paragraph with a specified style."""
        del style
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
        del rows, cols, header_rows, style, rtl
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
        del rows, cols
        raise NotImplementedError

    @abstractmethod
    def leave_table_cell(self) -> None:
        """Finalize table cell."""
        raise NotImplementedError

    @abstractmethod
    def set_character_style(self, style: OptionalStyle = None) -> None:
        """Start a span using a specific character style."""
        del style
        raise NotImplementedError

    @abstractmethod
    def enter_footnote(self) -> None:
        """Add a footnote reference and enter the footnote."""
        raise NotImplementedError

    @abstractmethod
    def leave_footnote(self) -> None:
        """Close footnote, go ack to main text."""
        raise NotImplementedError

    @abstractmethod
    def write_text(self, text: str) -> None:
        """Add some plain text."""


class ProxyOutput(IOutput):
    """A proxy IOutput base class."""

    def __init__(self, writer: IOutput) -> None:
        self.writer = writer

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
        return self.writer.write_text(text)

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


class WhitespaceStripper(ProxyOutput):
    """A proxy IOutput which strips all initial and final whitespace.

    Good foor footnotes.
    """

    begun = False

    def write_text(self, text: str) -> None:
        """Add some plain text."""
        if not self.begun:
            # Trim initial whitespace
            text = re.sub(r"^\s+", r"", text)
            self.begun = bool(text)

        return self.writer.write_text(text)
