"""Manual Formatting."""
import enum


class ManualFormat(enum.Flag):
    """Manual Formatting."""

    NORMAL = 0

    CENTERED = enum.auto()
    JUSTIFIED = enum.auto()
    NEW_PAGE = enum.auto()
    SPACED = enum.auto()
    RTL = enum.auto()
    LTR = enum.auto()
    INDENTED = enum.auto()

    BOLD = enum.auto()
    ITALIC = enum.auto()
    SUBSCRIPT = enum.auto()
    SUPERSCRIPT = enum.auto()
    RAISED = enum.auto()
    LOWERED = enum.auto()
    HIGHLIGHT = enum.auto()
