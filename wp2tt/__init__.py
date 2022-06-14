#!/usr/bin/env python3
"""A utility to convert word processor files (.docx, .odt) to InDesign's Tagged Text."""
import argparse
import collections
import configparser
import contextlib
import itertools
import logging
from pathlib import Path
import re
import shlex
import shutil
import sys

from typing import Dict
from typing import Iterator
from typing import List
from typing import Mapping
from typing import Optional
from typing import Set
from typing import Union

import attr

from wp2tt.version import WP2TT_VERSION
from wp2tt.ini import ini_fields
from wp2tt.ini import ConfigSection
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.format import ManualFormat
from wp2tt.styles import Style
from wp2tt.styles import Rule
from wp2tt.proxies import ByExtensionInput
from wp2tt.proxies import MultiInput
from wp2tt.output import IOutput
from wp2tt.output import WhitespaceStripper
from wp2tt.tagged_text import InDesignTaggedTextOutput


def main():
    """Entry point"""
    WordProcessorToInDesignTaggedText().run()


class StopMarkerFound(Exception):
    """We raise this to stop the presses."""


class BadReferenceInRule(Exception):
    """We raise this for bad ruels."""


class ParseDict(argparse.Action):
    """Helper class to convert KEY=VALUE pairs to a dict."""

    def __call__(self, parser, namespace, values, *args, **kwargs):
        setattr(namespace, self.dest, dict(val.split("=", 1) for val in values))


@attr.s(slots=True, frozen=True)
class ManualFormatCustomStyle:
    """Manual formatting, possibly applied to a custom style"""

    fmt = attr.ib(type=ManualFormat)
    unadorned = attr.ib(type=Optional[str])


@attr.s(slots=True)
class State:
    """Context of styles"""

    curr_char_style = attr.ib(default=None, type=Optional[Style])
    prev_para_style = attr.ib(default=None, type=Optional[Style])
    is_empty = attr.ib(default=True)
    is_post_empty = attr.ib(default=False)
    is_post_break = attr.ib(default=False)
    para_char_fmt = attr.ib(default=ManualFormat.NORMAL)


class WordProcessorToInDesignTaggedText:
    """Read a word processor file. Write an InDesign Tagged Text file.

    What's not to like?"""

    SETTING_FILE_ENCODING = "UTF-8"
    CONFIG_SECTION_NAME = "General"
    SPECIAL_GROUP = "(autogenerated)"
    DEFAULT_BASE = SPECIAL_GROUP + "/(Basic Style)"
    FOOTNOTE_REF_STYLE = SPECIAL_GROUP + "/(Footnote Reference in Text)"
    COMMENT_REF_STYLE = SPECIAL_GROUP + "/(Comment Reference)"

    IGNORED_STYLES = {
        "character": ["annotation reference"],
    }
    SPECIAL_STYLE = {
        "character": {
            COMMENT_REF_STYLE: {
                "idtt": "<pShadingColor:Cyain><pShadingOn:1><pShadingTint:100>",
            }
        },
        "paragraph": {
            "annotation text": {
                "name": SPECIAL_GROUP + "/(Comment Text)",
                "idtt": "<cSize:6><cColor:Cyan><cColorTint:100>",
            }
        },
    }

    args: argparse.Namespace
    base_names: Dict[str, str]
    base_styles: Dict[str, Style]
    comment_ref_style: Style
    config: Mapping[str, str]
    doc: IDocumentInput
    footnote_ref_style: Style
    format_mask: ManualFormat
    manual_styles: Dict[ManualFormatCustomStyle, Style]
    output_fn: Path
    parser: argparse.ArgumentParser
    rerunner_fn: Path
    rules: List[Rule]
    settings: configparser.ConfigParser
    settings_fn: Path
    settings_touched: bool
    state: State = State()
    stop_marker: str
    stop_marker_found: bool
    style_sections_used: Set[str]
    styles: Dict[str, Style]
    writer: IOutput

    def run(self):
        """Main entry point."""
        self.parse_command_line()
        self.configure_logging()
        self.read_settings()
        self.read_input()
        self.write_idtt()
        self.report_statistics()
        self.write_settings()
        self.write_rerunner()

    def parse_command_line(self):
        """Find out what we're supposed to do."""
        parser = self.parser = argparse.ArgumentParser(
            description=f"Word Processor to InDesign Tagged Text, v{WP2TT_VERSION}"
        )
        parser.add_argument("input", type=Path, help="Input word processor file")
        parser.add_argument(
            "output", type=Path, nargs="?", help="InDesign Tagged Text file"
        )
        parser.add_argument(
            "-a",
            "--append",
            type=Path,
            metavar="INPUT",
            nargs="*",
            help="Concatenate more input file(s) to the same output",
        )
        parser.add_argument(
            "-s",
            "--stop-at",
            metavar="TEXT",
            required=False,
            help="Stop importing when TEXT is found",
        )
        parser.add_argument(
            "-c",
            "--base-character-style",
            metavar="NAME",
            default=self.DEFAULT_BASE,
            help="Base all character styles on this.",
        )
        parser.add_argument(
            "-p",
            "--base-paragraph-style",
            metavar="NAME",
            default=self.DEFAULT_BASE,
            help="Base all paragraph styles on this.",
        )
        parser.add_argument(
            "-v",
            "--style-to-variable",
            metavar="STYLE=VARIABLE",
            nargs="+",
            action=ParseDict,
            help="Map paragraph styles to document variables.",
        )

        group = parser.add_mutually_exclusive_group()
        group.add_argument(
            "-m",
            "--manual",
            action="store_true",
            help="Create styles for some manual formatting.",
        )
        group.add_argument(
            "-M",
            "--manual-light",
            action="store_true",
            help="Like --manual, but only for character styles",
        )

        parser.add_argument(
            "-f",
            "--fresh-start",
            action="store_true",
            help="Do not read any existing settings.",
        )
        parser.add_argument(
            "-d",
            "--debug",
            action="store_true",
            help="Print interesting debug information.",
        )
        parser.add_argument(
            "-C",
            "--convert-comments",
            action="store_true",
            help="Convert comments to balloons.",
        )
        parser.add_argument(
            "--no-rerunner",
            action="store_true",
            help="Do not (over)write the rerruner script.",
        )
        parser.add_argument(
            "--direction",
            choices=["RTL", "LTR"],
            default="RTL",
            help="Default text direction.",
        )
        parser.add_argument(
            "--vav",
            action="store_true",
            help="Convert to VAV WITH HOLAM ligature.",
        )
        self.args = parser.parse_args()

        if self.args.output:
            self.output_fn = self.args.output
        else:
            self.output_fn = self.args.input.with_suffix(".txt")

        self.settings_fn = self.output_fn.with_suffix(".ini")
        self.rerunner_fn = Path(f"{self.output_fn}.rerun")
        self.stop_marker = self.args.stop_at
        self.format_mask = ~ManualFormat[self.args.direction]

    def configure_logging(self):
        """Set logging level and format."""
        logging.basicConfig(
            format="%(asctime)s %(message)s",
            level=logging.DEBUG if self.args.debug else logging.INFO,
        )

    def read_settings(self):
        """Read and parse the ini file."""
        self.settings = configparser.ConfigParser()
        self.settings_touched = False
        self.style_sections_used = set()
        if self.settings_fn.is_file() and not self.args.fresh_start:
            logging.info("Reading %s", self.settings_fn)
            self.settings.read(self.settings_fn, encoding=self.SETTING_FILE_ENCODING)

        self.load_rules()
        self.config = self.ensure_setting_section(self.CONFIG_SECTION_NAME)
        if self.stop_marker:
            self.config["stop_marker"] = self.stop_marker
        else:
            self.stop_marker = self.config.get("stop_marker")

    def load_rules(self):
        """Convert Rule sections into Rule objects."""
        self.rules = []
        for section_name in self.settings.sections():
            if not section_name.lower().startswith("rule:"):
                continue
            section = self.settings[section_name]
            self.rules.append(
                Rule(
                    mnemonic=f"R{len(self.rules) + 1}",
                    description=section_name[5:],
                    **{
                        name: section[ini_name]
                        for name, ini_name in ini_fields(Rule)
                        if ini_name in section
                    },
                )
            )
            logging.debug(self.rules[-1])

    def write_settings(self):
        """When done, write the settings file for the next time."""
        if self.settings_touched and self.settings_fn.is_file():
            logging.debug("Backing up %s", self.settings_fn)
            shutil.copy(self.settings_fn, self.settings_fn.with_suffix(".bak"))

        logging.info("Writing %s", self.settings_fn)
        with open(self.settings_fn, "w", encoding=self.SETTING_FILE_ENCODING) as fobj:
            self.settings.write(fobj)

    def write_rerunner(self):
        """Write a script to regenerate the output."""
        if self.args.no_rerunner:
            return

        logging.info("Writing %s", self.rerunner_fn)
        with open(self.rerunner_fn, "w", encoding=self.SETTING_FILE_ENCODING) as fobj:
            fobj.write("#!/bin/bash\n" "# AUTOGENERATED FILE, DO NOT EDIT.\n" "\n")
            cli = [
                self.quote_fn(sys.argv[0]),
                self.quote_fn(self.args.input),
            ]
            if self.args.output:
                cli.append(self.quote_fn(self.output_fn))
            cli.append('"$@"')  # Has to come before the dashes
            if self.stop_marker:
                cli.extend(["--stop-at", shlex.quote(self.stop_marker)])
            if self.args.base_character_style != self.DEFAULT_BASE:
                cli.extend(
                    [
                        "--base-character-style",
                        shlex.quote(self.args.base_character_style),
                    ]
                )
            if self.args.base_paragraph_style != self.DEFAULT_BASE:
                cli.extend(
                    [
                        "--base-paragraph-style",
                        shlex.quote(self.args.base_paragraph_style),
                    ]
                )
            if self.args.style_to_variable:
                cli.append("--style-to-variable")
                cli.extend(
                    shlex.quote(f"{k}={v}")
                    for k, v in self.args.style_to_variable.items()
                )
            if self.args.manual:
                cli.append("--manual")
            elif self.args.manual_light:
                cli.append("--manual-light")
            if self.args.vav:
                cli.append("--vav")
            if self.args.debug:
                cli.append("--debug")
            if self.args.append:
                cli.append("--append")
                cli.extend([self.quote_fn(path) for path in self.args.append])
            log_fn = Path(f"{self.rerunner_fn}.output")
            cli.extend(["2>&1", "|tee", log_fn.absolute()])
            fobj.write(" ".join(map(str, cli)))
            fobj.write("\n")
        self.rerunner_fn.chmod(0o755)

    def read_input(self):
        """Unzip and parse the input files."""
        logging.info("Reading %s", self.args.input)
        with self.create_reader() as self.doc:
            self.scan_style_definitions()
            self.scan_style_mentions()
        self.link_styles()
        self.link_rules()

    def create_reader(self) -> IDocumentInput:
        """Create the approriate document reader object"""
        if self.args.append:
            return MultiInput([self.args.input] + self.args.append)
        return ByExtensionInput(self.args.input)

    def scan_style_definitions(self) -> None:
        """Create a Style object for everything in the document."""
        self.styles = {}
        self.create_special_styles()
        counts: Mapping[str, Iterator[int]] = collections.defaultdict(
            lambda: itertools.count(start=1)
        )
        for style_kwargs in self.doc.styles_defined():
            if style_kwargs.get("automatic"):
                group = self.SPECIAL_GROUP
                num = next(counts[style_kwargs["realm"]])
                style_kwargs["name"] = f"{group}/automatic-{num}"
            self.found_style_definition(**style_kwargs)

    def create_special_styles(self) -> None:
        """Add any internal styles (i.e., not imported from the doc)."""
        self.base_names = {
            "character": self.args.base_character_style,
            "paragraph": self.args.base_paragraph_style,
        }
        self.base_styles = {
            realm: self.found_style_definition(
                realm=realm,
                internal_name=name,
                wpid=name,
                used=True,
                automatic=True,
            )
            for realm, name in self.base_names.items()
        }
        self.footnote_ref_style = self.found_style_definition(
            realm="character",
            internal_name=self.FOOTNOTE_REF_STYLE,
            wpid=self.FOOTNOTE_REF_STYLE,
            idtt="<cColor:Magenta><cColorTint:100><cPosition:Superscript>",
            automatic=True,
        )
        self.comment_ref_style = self.found_style_definition(
            realm="character",
            internal_name=self.COMMENT_REF_STYLE,
            wpid=self.COMMENT_REF_STYLE,
            parent_wpid=self.FOOTNOTE_REF_STYLE,
            idtt="<cColor:Cyan><cColorTint:100>",
            automatic=True,
        )
        self.manual_styles = {}

    def scan_style_mentions(self) -> None:
        """Mark which styles are actually used."""
        for realm, wpid in self.doc.styles_in_use():
            style_key = self.style_key(realm=realm, wpid=wpid)
            if style_key not in self.styles:
                logging.debug("Used but not defined? %r", style_key)
            elif not self.styles[style_key].used:
                logging.debug("Style used: %r", style_key)
                self.styles[style_key].used = True

    def link_styles(self) -> None:
        """A sort of alchemy-relationship thing."""
        for style in self.styles.values():
            style.parent_style = self.style_or_none(style.realm, style.parent_wpid)
            style.next_style = self.style_or_none(style.realm, style.next_wpid)

    def style_or_none(self, realm: str, wpid: str) -> Optional[Style]:
        """Given a realm/wpid pair, return our internal Style object"""
        if not wpid:
            return None
        return self.styles[self.style_key(realm=realm, wpid=wpid)]

    def link_rules(self) -> None:
        """A sort of alchemy-relationship thing."""
        for rule in self.rules:
            try:
                rule.turn_this_style = self.find_style_by_ini_ref(
                    rule.turn_this, required=True
                )
                rule.into_this_style = self.find_style_by_ini_ref(
                    rule.into_this,
                    required=True,
                    inherit_from=rule.turn_this_style,
                )
                if rule.when_following is not None:
                    rule.when_following_styles = [
                        self.find_style_by_ini_ref(ini_ref)
                        for ini_ref in re.findall(r"\[.*?\]", rule.when_following)
                    ]
            except BadReferenceInRule:
                logging.warning("Ignoring rule with bad references: %s", rule)
                rule.valid = False

    def find_style_by_ini_ref(
        self, ini_ref: str, required=False, inherit_from=None
    ) -> Optional[Style]:
        """Returns a style, given type of string we use for ini file section names."""
        if not ini_ref:
            if required:
                logging.debug("MISSING REQUIRED SOMETHING")
                raise BadReferenceInRule()
            return None
        mobj = re.match(
            r"^\[(?P<realm>\w+):(?P<internal_name>.+)\]$", ini_ref, re.IGNORECASE
        )
        if not mobj:
            logging.debug("Malformed %r", ini_ref)
            raise BadReferenceInRule()
        try:
            realm = mobj.group("realm").lower()
            internal_name = mobj.group("internal_name")
            return next(
                style
                for style in self.styles.values()
                if style.realm.lower() == realm and style.internal_name == internal_name
            )
        except StopIteration as exc:
            if not inherit_from:
                logging.debug("ERROR: Unknown %r", ini_ref)
                raise BadReferenceInRule() from exc
        return self.add_style(
            realm=realm,
            wpid=ini_ref,
            internal_name=internal_name,
            parent_wpid=inherit_from.wpid,
            parent_style=inherit_from,
        )

    def found_style_definition(
        self, realm: str, internal_name: str, wpid: str, **kwargs
    ) -> Style:
        """Called when encountering a style definition.

        Generate a Tagged Text style definition.
        """
        if realm not in self.base_names:
            logging.error("What about %s:%r [%r]?", realm, wpid, internal_name)
            self.base_names[realm] = self.args.base_character_style

        if parent_style := kwargs.get("parent_style"):
            kwargs.setdefault("parent_wpid", parent_style.wpid)
        elif wpid != self.base_names.get(realm):
            kwargs.setdefault("parent_wpid", self.base_names.get(realm))

        # Allow any special overrides (color, name, etc.)
        try:
            kwargs.update(self.SPECIAL_STYLE[realm][internal_name])
        except KeyError:
            pass

        section = self.get_setting_section(realm=realm, internal_name=internal_name)
        if not kwargs.get("name"):
            kwargs["name"] = internal_name
        kwargs.update(
            {
                name: section[ini_name]
                for name, ini_name in ini_fields(Style, writeable=True)
                if ini_name in section
            }
        )

        return self.add_style(
            realm=realm, wpid=wpid, internal_name=internal_name, **kwargs
        )

    def add_style(self, **kwargs) -> Style:
        """Create a new Style object and add to our map."""
        if self.args.style_to_variable:
            if kwargs["realm"] == "paragraph":
                kwargs.setdefault(
                    "variable", self.args.style_to_variable.get(kwargs["internal_name"])
                )

        kwargs.setdefault("name", kwargs["internal_name"])

        style = Style(**kwargs)
        logging.debug("Created %s", style)
        self.styles[self.style_key(style=style)] = style
        return style

    @classmethod
    def style_key(cls, style=None, realm=None, wpid=None) -> str:
        """The string which we use for `self.styles`.

        Note that this is based on the wpid, because
        that's the key docx files use for cross-reference.
        """
        if style:
            realm = style.realm
            wpid = style.wpid
        return f"{realm}:{wpid}"

    @classmethod
    def fix_section_name(
        cls,
        section_name: Optional[str] = None,
        realm: Optional[str] = None,
        internal_name: Optional[str] = None,
        style: Optional[Style] = None,
    ) -> str:
        """The name of the ini section for a given style.

        This uses `internal_name`, rather than `wpid` or `name`,
        because `wpid` can get ugly ("a2") and `name` should be
        modifyable.
        """
        if section_name is not None:
            return section_name

        if style:
            realm = style.realm
            internal_name = style.internal_name
        if realm:
            realm = realm.capitalize()
        return f"{realm}:{internal_name}"

    def write_idtt(self) -> None:
        """The main conversion loop: parse document, write tagged text"""
        logging.info("Writing %s", self.output_fn)
        self.set_state(State())
        with InDesignTaggedTextOutput(self.doc.properties) as self.writer:
            self.convert_document()
            self.write_output()

    def convert_document(self):
        """Convert a document, one paragraph at a time"""
        try:
            self.stop_marker_found = False
            self.state.is_post_empty = False
            self.state.is_post_break = False
            for para in self.doc.paragraphs():
                self.convert_paragraph(para)
            if self.stop_marker:
                logging.info("Note: Stop marker was never found")
                logging.debug("In other words, no %r", self.stop_marker)
        except StopMarkerFound as marker:
            logging.info(marker)

    def write_output(self):
        """Write the actual output file(s)"""
        text = self.writer.contents
        if self.args.vav:
            text = text.replace("\u05D5\u05B9", "\uFB4B")
        self.output_fn.parent.mkdir(parents=True, exist_ok=True)
        with open(self.output_fn, "w", encoding="UTF-16LE") as fo:
            fo.write(text)
        if self.args.debug:
            utf8_fn = self.output_fn.with_suffix(".utf8")
            with open(utf8_fn, "w", encoding="UTF-8") as fo:
                fo.write(text)

    def convert_paragraph(self, para: IDocumentParagraph) -> None:
        """Convert entire paragraph"""
        self.state.is_empty = True

        if self.stop_marker:
            self.check_for_stop_paragraph(para)
        style = self.apply_rules_to(self.get_paragraph_style(para))

        self.writer.enter_paragraph(style)
        for rng in para.spans():
            self.convert_range(rng)
        if style and style.variable:
            self.define_variable_from_paragraph(style.variable, para)
        self.writer.leave_paragraph()

        # For the next paragraph
        if para.is_page_break():
            # logging.debug("Paragraph %r has a page break", para)
            self.state.is_post_break = True
            self.state.is_empty = False
        elif self.state.is_post_break:
            if self.state.is_empty:
                logging.debug("Paragraph %r is empty, next is still post-break", para)
            else:
                self.state.is_post_break = False

        self.state.is_post_empty = self.state.is_empty
        self.state.prev_para_style = style

    def get_paragraph_style(self, para: IDocumentParagraph) -> Optional[Style]:
        """Return style to be used for a paragraph"""
        self.state.para_char_fmt = ManualFormat.NORMAL

        # The style object without any manual overrides
        unadorned = self.style("paragraph", para.style_wpid())
        if not self.args.manual:
            return unadorned

        # Manual formatting
        fmt = self.get_format(para)
        if self.state.is_post_break:
            fmt = fmt | ManualFormat.NEW_PAGE
        elif self.state.is_post_empty:
            fmt = fmt | ManualFormat.SPACED

        # Check for paragraph with a character style
        char_fmts = {self.get_format(rng) for rng in para.spans()}
        if len(char_fmts) == 1:
            self.state.para_char_fmt = char_fmts.pop()
            fmt = fmt | self.state.para_char_fmt

        return self.get_manual_style("paragraph", unadorned, fmt)

    def get_format(
        self, node: Union[IDocumentParagraph, IDocumentSpan]
    ) -> ManualFormat:
        """Return the marked part of a paragraph/span's format.

        i.e., masking out the default direction."""
        return node.format() & self.format_mask

    def get_manual_style(
        self, realm: str, unadorned: Optional[Style], fmt: ManualFormat
    ) -> Optional[Style]:
        """When using manual formatting, create/get a style"""
        if unadorned is not None:
            if not unadorned.custom:
                # Only look at unadorned style if it's custom
                unadorned = None

        if fmt and unadorned is not None:
            fmt &= ~unadorned.fmt

        if not fmt:  # No manual formatting
            # Only paragraphs get a named "NORMAL" style
            if unadorned is not None or realm != "paragraph":
                return unadorned

        uaid = unadorned.wpid if unadorned else None
        mfcs = ManualFormatCustomStyle(fmt, uaid)
        if mfcs in self.manual_styles:
            return self.manual_styles[mfcs]

        if fmt:
            fmtname = "_".join(f.name for f in ManualFormat if fmt & f and f.name)
        else:
            fmtname = fmt.name or "DEFAULT"
            logging.debug("%r -> %r", fmt, fmtname)

        # In the manual case, we may not have the basic style
        parent = unadorned or self.style(realm=realm, wpid=self.base_names[realm])

        if unadorned:
            name = f"{self.SPECIAL_GROUP}/{unadorned.name} ({fmtname})"
        else:
            name = f"{self.SPECIAL_GROUP}/({fmtname})"
        self.manual_styles[mfcs] = self.found_style_definition(
            realm=realm,
            internal_name=name,
            wpid=name,
            parent_style=parent,
            automatic=True,
        )
        return self.manual_styles[mfcs]

    def apply_rules_to(self, style: Optional[Style]) -> Optional[Style]:
        """Convert style according to user-defined rules"""
        if style:
            for rule in self.rules:
                if self.rule_applies_to(rule, style):
                    rule.applied += 1
                    return rule.into_this_style
        return style

    def rule_applies_to(self, rule: Rule, style: Style) -> bool:
        """True iff `rule` is should be applied on `style`."""
        if not rule.valid:
            return False
        if rule.turn_this_style is not style:
            return False
        if rule.when_first_in_doc:
            if style.count > 1:
                return False
        if rule.when_following_styles:
            if self.state.prev_para_style not in rule.when_following_styles:
                return False
        return True

    def check_for_stop_paragraph(self, para: IDocumentParagraph) -> None:
        """Looks for stop marker, raises StopMarkerFound() if found"""
        text = ""
        for chunk in para.text():
            text += chunk
            if text.startswith(self.stop_marker):
                raise StopMarkerFound(
                    "Stop marker found at the beginning of a paragraph"
                )
            if len(text) >= len(self.stop_marker):
                return

    def define_variable_from_paragraph(self, variable: str, para: IDocumentParagraph):
        """Save contents of a paragraph to a text variable"""
        self.writer.define_text_variable(
            variable, "".join(chunk for chunk in para.text())
        )

    def set_state(self, state: State) -> State:
        """Set current style configuration, return previous one."""
        prev = self.state
        self.state = state
        return prev

    def convert_range(self, rng: IDocumentSpan):
        """Convert all text and styles in a Range"""
        self.switch_character_style(self.get_character_style(rng))
        self.convert_range_text(rng)
        for footnote in rng.footnotes():
            self.convert_footnote(footnote)
            if self.state.is_post_break:
                logging.debug("Has footnote, not empty")
            self.state.is_empty = False
        if self.args.convert_comments:
            for cmt in rng.comments():
                self.convert_comment(cmt)

    NON_WHITESPACE = re.compile(r"\S")

    def get_character_style(self, rng: IDocumentSpan) -> Optional[Style]:
        """Returns Style object for a Range"""
        unadorned = self.style("character", rng.style_wpid())
        if not self.args.manual and not self.args.manual_light:
            return unadorned

        # Don't look at manual formatting for all-whitespace span
        # TODO: This breaks on "<LTR>Hey</LTR> <LTR>World</LTR>"
        # for text in rng.text():
        #     if self.NON_WHITESPACE.search(text):
        #         break
        # else:
        #     return unadorned

        # Manual formatting
        fmt = self.get_format(rng)
        if fmt == self.state.para_char_fmt:
            # Format already included in paragraph style
            fmt = ManualFormat.NORMAL
        return self.get_manual_style("character", unadorned, fmt)

    def convert_range_text(self, rng: IDocumentSpan):
        """Convert text in a Range object"""
        for text in rng.text():
            self.write_text(text)
            if self.NON_WHITESPACE.search(text):
                self.state.is_empty = False

    def switch_character_style(self, style):
        """Set current character style"""
        prev = self.state.curr_char_style
        if style is not prev:
            self.writer.set_character_style(style)
            self.state.curr_char_style = style
        return prev

    def write_text(self, text):
        """Add some plain text."""
        with contextlib.ExitStack() as stack:
            stack.callback(lambda: self.do_write_text(text))
            if not self.stop_marker:
                return
            offset = text.find(self.stop_marker)
            if offset < 0:
                return
            text = text[:offset]
            raise StopMarkerFound("Stop marker found")

    def do_write_text(self, text):
        """Actually send text to `self.writer`."""
        if not text:
            return
        self.writer.write_text(text)

    def convert_footnote(self, footnote, ref_style=None):
        """Convert one footnote to tagged text."""
        with self.FootnoteContext(self, ref_style):
            for par in footnote.paragraphs():
                self.convert_paragraph(par)

    def convert_comment(self, cmt):
        """Tagged Text doesn't support notes, so we convert them to footnotes."""
        return self.convert_footnote(cmt, ref_style=self.comment_ref_style)

    class FootnoteContext(contextlib.AbstractContextManager):
        """Context manager for style inside footnotes"""

        def __init__(self, outer, ref_style=None):
            super().__init__()
            self.outer = outer
            self.writer = outer.writer
            self.outer_character_style = outer.state.curr_char_style
            ref_style = ref_style or outer.footnote_ref_style

            self.outer.activate_style(ref_style)
            self.writer.set_character_style(ref_style)
            self.writer.enter_footnote()
            self.outer.writer = WhitespaceStripper(self.writer)
            self.outer_state = self.outer.set_state(State())

        def __exit__(self, *args):
            self.outer.set_state(self.outer_state)
            self.writer.leave_footnote()
            self.writer.set_character_style(self.outer_character_style)
            self.outer.writer = self.writer

    def report_statistics(self) -> None:
        """Print a nice summary"""
        realms = {s.realm for s in self.styles.values()}
        for realm in realms:
            logging.info(
                "Number of %s styles used: %u",
                realm.capitalize(),
                sum(1 for s in self.styles.values() if s.realm == realm and s.used),
            )
        for rule in self.rules:
            if rule.applied:
                logging.info("%s application(s) of %s", rule.applied, rule)

    def style(self, realm: str, wpid: Optional[str]) -> Optional[Style]:
        """Return a Style object"""
        if not wpid:
            return None

        style = self.styles[self.style_key(realm=realm, wpid=wpid)]
        if realm in self.IGNORED_STYLES:
            if style.internal_name in self.IGNORED_STYLES[realm]:
                return None

        self.activate_style(style)
        style.count += 1
        return style

    def activate_style(self, style: Style) -> None:
        """Define the style if needed."""
        section_name = self.fix_section_name(
            realm=style.realm, internal_name=style.internal_name
        )
        if section_name in self.style_sections_used:
            return

        if style.parent_style:
            if not style.parent_style.used:
                logging.debug(
                    "[%s] leads to missing %r", section_name, style.parent_wpid
                )
                style.parent_style = self.base_styles[style.realm]
                style.parent_wpid = style.parent_style.wpid
            self.activate_style(style.parent_style)

        self.update_setting_section(section_name, style)
        self.style_sections_used.add(section_name)

        if style.next_style and style.next_style.used:
            self.activate_style(style.next_style)
        elif style.next_wpid:
            logging.debug("[%s] leads to missing %r", section_name, style.next_wpid)

    def get_setting_section(
        self,
        section_name: Optional[str] = None,
        realm: Optional[str] = None,
        internal_name: Optional[str] = None,
        style: Optional[Style] = None,
    ) -> ConfigSection:
        """Return a section from the ini file.

        If no such section exists, return a new dict, but don't add it
        to the configuration.
        """
        actual_name = self.fix_section_name(
            section_name=section_name,
            realm=realm,
            internal_name=internal_name,
            style=style,
        )

        if self.settings.has_section(actual_name):
            return self.settings[actual_name]
        return {}

    def ensure_setting_section(
        self,
        section_name: Optional[str] = None,
        realm: Optional[str] = None,
        internal_name: Optional[str] = None,
        style: Optional[Style] = None,
    ) -> ConfigSection:
        """Return a section from the ini file, adding one if needed."""
        actual_name = self.fix_section_name(
            section_name=section_name,
            realm=realm,
            internal_name=internal_name,
            style=style,
        )
        if not self.settings.has_section(actual_name):
            self.settings[actual_name] = {}
        return self.settings[actual_name]

    def update_setting_section(self, section_name: str, style: Style) -> None:
        """Update key:value pairs in an ini section.

        Sets `self.settings_touched` if anything was changed.
        """
        section = self.ensure_setting_section(section_name)
        for name, ini_name in ini_fields(Style):
            value = getattr(style, name)
            self.settings_touched |= section.get(ini_name) != value
            if value:
                section[ini_name] = str(value or "")
            else:
                section.pop(ini_name, None)

    @classmethod
    def quote_fn(cls, path: Union[Path, str]) -> str:
        """Courtesy wrapper"""
        if isinstance(path, str):
            path = Path(path)
        return shlex.quote(str(path.absolute()))


# TODO:
# - DOCX: find default RTL
# - DOCX: <w:bookmarkStart w:name="X"/> / <w:instrText>PAGEREF ..
# - ODT: footnotes
# - ODT: comments
# - PUB: Support non-ME docs
# - PUB: Manual
# - many-to-one wp_name -> name
# - mapping csv
# - [paragraph rule] when_first_in_doc
# - [paragraph rule] when_matches_re
# - [paragraph style] keep_last_n_chars
# - character style rule (grep)
# - Non-unicode when not required?
# - Import MarkDown
# - Paragraph direction (w:r/w:rPr/w:rtl -> <pParaDir:1>; but what about the basic dir?)
# - For post edit/proof: Manual formatting consolidation, TBD
# - Para: global base -> body base, heading base
# - More rule context: after same, after different, first since...
# - Really need a test suite of some sort.
# - Manual format: collapse with existing styles
# - A flag to only create/update the ini file
# - Maybe add front matter (best done in Id? either that or jinja2!)
# - Something usable for those balloons (footnote+hl? endnote? convert to note in jsx?)
#   bold: w:b (w:bCs?); italic: w:i (w:iCs?); undeline w:u
#   font: <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New">
#   override style: <w:i w:val="0">
# - (f)odt import
# - Convert editing marks
# - idml import
# - Automatic header group
# - More complex BiDi
# - Endnotes
# - Linked styles?
# - Derivation rules?
# - Latent styles?
# - Digraph kerning (probably better in InDesign?)
