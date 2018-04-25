#!/usr/bin/env python3
import re


class IOutput(object):
    """Interface for things that write InDesign Tagged Text."""

    def define_style(self, style):
        """Add a style definition."""
        raise NotImplementedError()

    def define_text_variable(self, name, value):
        """Add a text variable."""
        raise NotImplementedError()

    def enter_paragraph(self, style=None):
        """Start a paragraph with a specified style."""
        raise NotImplementedError()

    def leave_paragraph(self):
        """Finalize paragraph."""
        raise NotImplementedError()

    def set_character_style(self, style=None):
        """Start a span using a specific character style."""
        raise NotImplementedError()

    def enter_footnote(self):
        """Add a footnote reference and enter the footnote."""
        raise NotImplementedError()

    def leave_footnote(self):
        """Close footnote, go ack to main text."""
        raise NotImplementedError()

    def write_text(self, text):
        """Add some plain text."""
        raise NotImplementedError()


class WhitespaceStripper(IOutput):
    """A proxy IOutput which strips all initial and final whitespace.

    Good foor footnotes.
    """
    def __init__(self, writer):
        self.writer = writer
        self.begun = False
        self.pending = ''

    def define_style(self, style):
        """Add a style definition."""
        return self.writer.define_style(style)

    def define_text_variable(self, name, value):
        """Add a text variable."""
        return self.writer.define_text_variable(name, value)

    def enter_paragraph(self, style=None):
        """Start a paragraph with a specified style."""
        return self.writer.enter_paragraph(style)

    def leave_paragraph(self):
        """Finalize paragraph."""
        return self.writer.leave_paragraph()

    def set_character_style(self, style=None):
        """Start a span using a specific character style."""
        return self.writer.set_character_style(style)

    def enter_footnote(self):
        """Add a footnote reference and enter the footnote."""
        return self.writer.enter_footnote()

    def leave_footnote(self):
        """Close footnote, go ack to main text."""
        return self.writer.leave_footnote()

    def write_text(self, text):
        """Add some plain text."""
        if not self.begun:
            # Trim initial whitespace
            text = re.sub(r'^\s+', r'', text)
            self.begun = bool(text)

        if self.begun:
            # Defer any possibly final whitespace
            parts = re.split(r'(\s+$)', self.pending + text)
            text = parts[0]
            self.pending = parts[1] if len(parts) > 1 else ''

        return self.writer.write_text(text)


