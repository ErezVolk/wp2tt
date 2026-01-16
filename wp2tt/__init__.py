#!/usr/bin/env python3
"""Convert word processor files (.docx, .odt) to InDesign's Tagged Text."""
import argparse
import collections
import contextlib
import dataclasses as dcl
from datetime import datetime
from datetime import timezone
from datetime import timedelta
import itertools
import logging
import os
from pathlib import Path
import re
import shlex
import shutil
import subprocess
import typing as t

import cairosvg

from wp2tt.cache import Cache
from wp2tt.format import ManualFormat
from wp2tt.ini import SettingsFile
from wp2tt.input import IDocComment
from wp2tt.input import IDocFootnote
from wp2tt.input import IDocFormula
from wp2tt.input import IDocImage
from wp2tt.input import IDocInput
from wp2tt.input import IDocParagraph
from wp2tt.input import IDocSpan
from wp2tt.input import IDocTable
from wp2tt.mathml import MathConverter
from wp2tt.output import IOutput
from wp2tt.output import WhitespaceStripper
from wp2tt.proxies import ByExtensionInput
from wp2tt.proxies import MultiInput
from wp2tt.proxies import ProxyInput
from wp2tt.styles import OptionalStyle
from wp2tt.styles import Rule
from wp2tt.styles import Style
from wp2tt.tagged_text import InDesignTaggedTextOutput
from wp2tt.usage import Wp2ttParser

log = logging.getLogger(__name__)


def main() -> None:
    """Entry point."""
    WordProcessorToInDesignTaggedText().run()


class StopMarkerFoundError(Exception):
    """We raise this to stop the presses."""


class BadReferenceInRuleError(Exception):
    """We raise this for bad ruels."""


@dcl.dataclass(frozen=True)
class ManualFormatCustomStyle:
    """Manual formatting, possibly applied to a custom style."""

    fmt: ManualFormat
    unadorned: str | None


@dcl.dataclass
class State:
    """Context of styles."""

    curr_char_style: OptionalStyle = None
    active_char_style: OptionalStyle = None
    prev_para_style: OptionalStyle = None
    is_empty: bool = True
    is_post_empty: bool = False
    is_post_break: bool = False
    para_char_fmt: ManualFormat = ManualFormat.NORMAL


class WordProcessorToInDesignTaggedText:
    """Read a word processor file. Write an InDesign Tagged Text file.

    What's not to like?
    """

    SPECIAL_GROUP = Wp2ttParser.SPECIAL_GROUP
    FOOTNOTE_REF_STYLE = SPECIAL_GROUP + "/(Footnote Reference in Text)"
    COMMENT_REF_STYLE = SPECIAL_GROUP + "/(Comment Reference)"
    COMMENT_STYLE = SPECIAL_GROUP + "/(Comment)"
    IMAGE_STYLE = SPECIAL_GROUP + "/(Image)"
    FORMULA_STYLE = SPECIAL_GROUP + "/(Formula)"
    TABLE_PARAGRAPH_STYLE = SPECIAL_GROUP + "/(Table Container)"

    IGNORED_STYLES: t.Mapping[str, list[str]] = {
        "character": ["annotation reference"],
    }
    STYLE_OVERRIDE: t.Mapping = {
        "character": {
            COMMENT_REF_STYLE: {
                "idtt": "<pShadingColor:Cyan><pShadingOn:1><pShadingTint:100>",
            },
        },
        "paragraph": {
            "annotation text": {
                "name": SPECIAL_GROUP + "/(Comment Text)",
                "idtt": "<cSize:6><cColor:Cyan><cColorTint:100>",
            },
        },
    }

    INDESIGN_ENV_VAR = "__CFBundleIdentifier"
    INDESIGN_ENV_VAL = "com.adobe.InDesign"

    args: argparse.Namespace
    base_names: dict[str, str]
    base_styles: dict[str, Style]
    cache: Cache
    comment_style: Style
    comment_ref_style: Style
    config: t.Mapping[str, str]
    doc: IDocInput
    footnote_ref_style: Style
    format_mask: ManualFormat
    formula_style: Style
    image_count = itertools.count(1)
    image_dir: Path
    image_style: Style
    manual_styles: dict[ManualFormatCustomStyle, Style]
    output_dir: Path
    output_fn: Path
    output_stem: str
    parser: Wp2ttParser
    rerunner_fn: Path
    rules: list[Rule]
    settings: SettingsFile
    settings_fn: Path
    state: State = State()
    stop_marker: str
    stop_marker_found: bool
    style_sections_used: set[str]
    styles: dict[str, Style]
    table_paragraph_style: Style
    writer: IOutput

    def run(self) -> None:
        """Do it."""
        self.parse_command_line()
        self.configure_logging()
        self.read_settings()
        self.read_input()
        self.write_idtt()
        self.report_statistics()
        self.write_settings()
        self.write_rerunner()

    def parse_command_line(self) -> None:
        """Find out what we're supposed to do."""
        parser = self.parser = Wp2ttParser()
        self.args = parser.parse_args()

        if self.args.output:
            self.output_fn = self.args.output
        else:
            self.output_fn = self.args.input.with_suffix(".txt")

        self.output_stem = self.output_fn.stem
        self.output_dir = self.output_fn.parent.resolve()
        self.settings_fn = self.output_fn.with_suffix(".ini")
        self.rerunner_fn = Path(f"{self.output_fn}.rerun")
        self.stop_marker = self.args.stop_at
        if self.args.direction == "IGNORE":
            self.format_mask = ManualFormat.NORMAL
        else:
            self.format_mask = ~ManualFormat[self.args.direction]

        if self.args.cache:
            self.cache = Cache(self.args.cache)
        elif not self.args.no_cache:
            self.cache = Cache(self.output_dir / "cache")
        else:
            self.cache = Cache()

    def configure_logging(self) -> None:
        """Set logging level and format."""
        logging.basicConfig(
            format="%(asctime)s %(message)s",
            level=logging.DEBUG if self.args.debug else logging.INFO,
        )

    def read_settings(self) -> None:
        """Read and parse the ini file."""
        self.style_sections_used = set()
        self.settings = SettingsFile(
            self.settings_fn,
            fresh_start=self.args.fresh_start,
        )

        self.load_rules()
        self.config = self.settings.ensure_section("General")
        if self.stop_marker:
            self.config["stop_marker"] = self.stop_marker
        else:
            self.stop_marker = self.config.get("stop_marker") or ""
            self.args.stop_at = self.stop_marker  # For rerunner

    def load_rules(self) -> None:
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
                        for name, ini_name in self.settings.fields(Rule)
                        if ini_name in section
                    },
                ),
            )
            log.debug(self.rules[-1])

    def write_settings(self) -> None:
        """When done, write the settings file for the next time."""
        self.settings.backup_and_write()

    def write_rerunner(self) -> None:
        """Write a script to regenerate the output."""
        if not self.args.no_rerunner:
            self.parser.write_rerunner(self.rerunner_fn, self.args)

    def read_input(self) -> None:
        """Unzip and parse the input files."""
        log.info("Reading %s", self.args.input)
        self.doc = self.create_reader()
        self.scan_style_definitions()
        self.scan_style_mentions()
        self.link_styles()
        self.link_rules()

    def create_reader(self) -> ProxyInput:
        """Create the approriate document reader object."""
        inputs = [self.args.input] + (self.args.append or [])
        if not self.args.no_input_rerunner:
            self.consider_input_rerunners(inputs)
        if len(inputs) > 1:
            return MultiInput(inputs, self.args)
        return ByExtensionInput(inputs[0], self.args)

    def consider_input_rerunners(self, inputs: list[Path]) -> None:
        """Look for input rerunners and run them."""
        for path in inputs:
            rerun = path.with_name(f"{path.name}.rerun")
            if not rerun.is_file():
                continue
            log.info("%s exists, trying to run it...", rerun)
            try:
                subprocess.run([str(rerun.resolve())], check=True)
            except PermissionError:
                log.warning("No permission to run %s", rerun)
            except subprocess.CalledProcessError as exc:
                log.error("%s exited with code %s", rerun, exc.returncode)
            except OSError as exc:
                log.error("Error running %s: %s", rerun, exc)

    def scan_style_definitions(self) -> None:
        """Create a Style object for everything in the document."""
        self.styles = {}
        self.create_special_styles()
        counts: t.Mapping[str, t.Iterator[int]] = collections.defaultdict(
            lambda: itertools.count(start=1),
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
            "table": InDesignTaggedTextOutput.BASIC_TABLE_STYLE,
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
        self.comment_style = self.found_style_definition(
            realm="paragraph",
            internal_name=self.COMMENT_STYLE,
            wpid=self.COMMENT_STYLE,
            parent_wpid=self.args.base_paragraph_style,
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
        self.image_style = self.found_style_definition(
            realm="character",
            internal_name=self.IMAGE_STYLE,
            wpid=self.IMAGE_STYLE,
            idtt="<cColor:Yellow><cColorTint:100>",
            automatic=True,
        )
        self.formula_style = self.found_style_definition(
            realm="character",
            internal_name=self.FORMULA_STYLE,
            wpid=self.FORMULA_STYLE,
            idtt="<cColor:Yellow><cColorTint:100>",
            automatic=True,
        )
        self.table_paragraph_style = self.found_style_definition(
            realm="paragraph",
            internal_name=self.TABLE_PARAGRAPH_STYLE,
            wpid=self.TABLE_PARAGRAPH_STYLE,
            automatic=True,
        )
        self.manual_styles = {}

    def scan_style_mentions(self) -> None:
        """Mark which styles are actually used."""
        realms = set()
        for realm, wpid in self.doc.styles_in_use():
            realms.add(realm)
            style_key = self.style_key(realm=realm, wpid=wpid)
            if style_key not in self.styles:
                log.debug("Used but not defined? %r", style_key)
            elif not self.styles[style_key].used:
                log.debug("Style used: %r", style_key)
                self.styles[style_key].used = True
        if "table" in realms:
            log.debug("Here be tables")
            self.table_paragraph_style.used = True

    def link_styles(self) -> None:
        """Do a sort of alchemy-relationship thing."""
        for style in self.styles.values():
            style.parent_style = self.style_or_none(
                style.realm, style.parent_wpid,
            )
            style.next_style = self.style_or_none(style.realm, style.next_wpid)

    def define_styles(self) -> None:
        """Write (partial) style definition section."""
        for style in self.styles.values():
            if style.used:
                self.writer.define_style(style)

    def style_or_none(self, realm: str, wpid: str | None) -> OptionalStyle:
        """Given a realm/wpid pair, return our internal Style object."""
        if not wpid:
            return None
        return self.styles[self.style_key(realm=realm, wpid=wpid)]

    def link_rules(self) -> None:
        """Do a sort of alchemy-relationship thing."""
        for rule in self.rules:
            try:
                rule.turn_this_style = self.find_style_by_ini_ref(
                    rule.turn_this, required=True,
                )
                rule.into_this_style = self.find_style_by_ini_ref(
                    rule.into_this,
                    required=True,
                    inherit_from=rule.turn_this_style,
                )
                if rule.when_following is not None:
                    wfs = [
                        self.find_style_by_ini_ref(ini_ref)
                        for ini_ref in re.findall(r"\[.*?\]", rule.when_following)
                    ]
                    rule.when_following_styles = [
                        style for style in wfs
                        if style is not None
                    ]
            except BadReferenceInRuleError:
                log.warning("Ignoring rule with bad references: %s", rule)
                rule.valid = False

    def find_style_by_ini_ref(
        self,
        ini_ref: str | None,
        *,
        required: bool = False,
        inherit_from: Style | None = None,
    ) -> OptionalStyle:
        """Return a style, given type of string we use for ini file section names."""
        if not ini_ref:
            if required:
                log.debug("MISSING REQUIRED SOMETHING")
                raise BadReferenceInRuleError
            return None
        mobj = re.match(
            r"^\[(?P<realm>\w+):(?P<internal_name>.+)\]$", ini_ref, re.IGNORECASE,
        )
        if not mobj:
            log.debug("Malformed %r", ini_ref)
            raise BadReferenceInRuleError
        realm = mobj.group("realm").lower()
        internal_name = mobj.group("internal_name")
        try:
            return next(
                style
                for style in self.styles.values()
                if style.realm.lower() == realm and style.internal_name == internal_name
            )
        except StopIteration as exc:
            if not inherit_from:
                log.debug("ERROR: Unknown %r", ini_ref)
                raise BadReferenceInRuleError from exc
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
        """Handle a style definition.

        Generate a Tagged Text style definition.
        """
        if realm not in self.base_names:
            log.error("What about %s:%r [%r]?", realm, wpid, internal_name)
            self.base_names[realm] = self.args.base_character_style

        if parent_style := kwargs.get("parent_style"):
            kwargs.setdefault("parent_wpid", parent_style.wpid)
        elif wpid != self.base_names.get(realm):
            kwargs.setdefault("parent_wpid", self.base_names.get(realm))

        # Allow any special overrides (color, name, etc.)
        with contextlib.suppress(KeyError):
            kwargs.update(self.STYLE_OVERRIDE[realm][internal_name])

        section = self.settings.get_section(realm=realm, internal_name=internal_name)
        if not kwargs.get("name"):
            kwargs["name"] = internal_name
        kwargs.update(
            {
                name: section[ini_name]
                for name, ini_name in self.settings.fields(Style, writeable=True)
                if ini_name in section
            },
        )

        return self.add_style(
            realm=realm, wpid=wpid, internal_name=internal_name, **kwargs,
        )

    def add_style(self, **kwargs) -> Style:
        """Create a new Style object and add to our map."""
        if self.args.style_to_variable:
            if kwargs["realm"] == "paragraph":
                kwargs.setdefault(
                    "variable",
                    self.args.style_to_variable.get(kwargs["internal_name"]),
                )

        kwargs.setdefault("name", kwargs["internal_name"])

        style = Style(**kwargs)
        log.debug("Created %s", style)
        self.styles[self.style_key(style=style)] = style
        return style

    @classmethod
    def style_key(
        cls,
        style: Style | None = None,
        realm: str | None = None,
        wpid: str | None = None,
    ) -> str:
        """Create the string which we use for `self.styles`.

        Note that this is based on the wpid, because
        that's the key docx files use for cross-reference.
        """
        if style:
            realm = style.realm
            wpid = style.wpid
        return f"{realm}:{wpid}"

    def write_idtt(self) -> None:
        """Run the main conversion loop: parse document, write tagged text."""
        log.info("Writing %s", self.output_fn)
        self.set_state(State())
        self.create_output()
        properties = self.doc.properties
        newline = "\u2028" if self.args.newline else "\n"
        with InDesignTaggedTextOutput(properties, newline) as self.writer:
            self.define_styles()
            self.convert_document()
            self.write_output()

    def create_output(self) -> None:
        """Create output directory, clean old image dirs."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        now = datetime.now(tz=timezone(timedelta(0)))
        self.image_dir = self.output_dir / now.strftime("img-%Y%m%d-%H%M")
        if self.args.remove_old_images_if_indesign:
            remove_old = os.environ.get(self.INDESIGN_ENV_VAR) == self.INDESIGN_ENV_VAL
        else:
            remove_old = self.args.remove_old_images
        if remove_old:
            for path in self.output_dir.glob("img-????????-????"):
                if path.is_dir():
                    log.debug("Removing %s", path)
                    shutil.rmtree(path, ignore_errors=True)

    def convert_document(self) -> None:
        """Convert a document, one paragraph at a time."""
        try:
            self.stop_marker_found = False
            self.state.is_post_empty = False
            self.state.is_post_break = False
            for node in self.doc.paragraphs():
                if isinstance(node, IDocTable):
                    self.convert_table(node)
                else:
                    self.convert_paragraph(node)
            if self.stop_marker:
                log.info("Note: Stop marker was never found")
                log.debug("In other words, no %r", self.stop_marker)
        except StopMarkerFoundError as marker:
            log.info(marker)

    def write_output(self) -> None:
        """Write the actual output file(s)."""
        assert isinstance(self.writer, InDesignTaggedTextOutput)
        text = self.writer.contents
        if self.args.maqaf:
            text = text.replace("=", "\u05BE")
        if self.args.vav:
            text = text.replace("\u05D5\u05B9", "\uFB4B")
        with Path(self.output_fn).open("w", encoding="UTF-16LE") as fobj:
            fobj.write(text)
        if self.args.debug:
            utf8_fn = self.output_fn.with_suffix(".utf8")
            with Path(utf8_fn).open("w", encoding="UTF-8") as fobj:
                fobj.write(text)

    def convert_table(self, table: IDocTable) -> None:
        """Convert entire table."""
        self.writer.enter_paragraph(self.table_paragraph_style)

        (rows, cols) = table.shape
        rtl = (table.format() & ManualFormat.RTL) == ManualFormat.RTL
        style = self.style("table", table.style_wpid())
        self.writer.enter_table(
            rows=rows,
            cols=cols,
            header_rows=table.header_rows,
            rtl=rtl,
            style=style,
        )
        for row in table.rows():
            self.writer.enter_table_row()
            for cell in row.cells():
                (cell_rows, cell_cols) = cell.shape
                self.writer.enter_table_cell(rows=cell_rows, cols=cell_cols)
                self.convert_paragraph(cell.contents())
                self.writer.leave_table_cell()
            self.writer.leave_table_row()
        self.writer.leave_table()

        self.writer.leave_paragraph()

    def convert_paragraph(
        self,
        para: IDocParagraph,
        default_style: OptionalStyle = None,
    ) -> None:
        """Convert entire paragraph."""
        self.state.is_empty = True

        if self.stop_marker:
            self.check_for_stop_paragraph(para)
        style = self.apply_rules_to(para) or default_style

        self.writer.enter_paragraph(style)
        for chunk in para.chunks():
            self.convert_chunk(chunk)
        if style and style.variable:
            self.define_variable_from_paragraph(style.variable, para)
        self.writer.leave_paragraph()

        # For the next paragraph
        if para.is_page_break():
            self.state.is_post_break = True
            self.state.is_empty = False
        elif self.state.is_post_break:
            if self.state.is_empty:
                log.debug("Paragraph %r is empty, next is still post-break", para)
            else:
                self.state.is_post_break = False

        self.state.is_post_empty = self.state.is_empty
        self.state.prev_para_style = style

    def get_paragraph_style(self, para: IDocParagraph) -> OptionalStyle:
        """Return style to be used for a paragraph."""
        self.state.para_char_fmt = ManualFormat.NORMAL

        # The style object without any manual overrides
        unadorned = self.style("paragraph", para.style_wpid())
        if not self.args.manual:
            return unadorned

        # Manual formatting
        fmt = para.format() & self.format_mask
        if self.state.is_post_break:
            fmt = fmt | ManualFormat.NEW_PAGE
        elif self.state.is_post_empty:
            fmt = fmt | ManualFormat.SPACED

        for span in para.spans():
            for text in span.text():
                if text[0].isspace():
                    fmt = fmt | ManualFormat.INDENTED
                break  # Just the first
            break  # Just the first

        # Check for paragraph with a character style
        char_fmts = {self.get_span_format(span) for span in para.spans()}
        if len(char_fmts) == 1:
            self.state.para_char_fmt = char_fmts.pop()
            fmt = fmt | self.state.para_char_fmt

        return self.get_manual_style("paragraph", unadorned, fmt)

    def get_span_format(
        self, node: IDocSpan,
    ) -> ManualFormat:
        """Return the marked part of a paragraph/span's format."""
        return node.format() & self.format_mask

    def get_manual_style(
        self, realm: str, unadorned: OptionalStyle, fmt: ManualFormat,
    ) -> OptionalStyle:
        """When using manual formatting, create/get a style."""
        # ? if unadorned is not None:
        # ?     if not unadorned.custom:
        # ?         # Only look at unadorned style if it's custom
        # ?         unadorned = None

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
            log.debug("%r -> %r", fmt, fmtname)

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

    def apply_rules_to(self, para: IDocParagraph) -> OptionalStyle:
        """Convert style according to user-defined rules."""
        style = self.get_paragraph_style(para)
        if style:
            for rule in self.rules:
                if self.rule_applies_to(rule, style, para.is_empty):
                    rule.applied += 1
                    return rule.into_this_style
        return style

    def rule_applies_to(
        self,
        rule: Rule,
        style: Style,
        is_empty: t.Callable[[], bool],
    ) -> bool:
        """Check if `rule` should be applied on `style`."""
        if not rule.valid:
            return False
        if rule.unless_empty:
            if is_empty():
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

    def check_for_stop_paragraph(self, para: IDocParagraph) -> None:
        """Look for stop marker, raises StopMarkerFoundError() if found."""
        text = ""
        for chunk in para.text():
            text += chunk
            if text.startswith(self.stop_marker):
                raise StopMarkerFoundError(
                    "Stop marker found at the beginning of a paragraph",
                )
            if len(text) >= len(self.stop_marker):
                return

    def define_variable_from_paragraph(
        self,
        variable: str,
        para: IDocParagraph,
    ) -> None:
        """Save contents of a paragraph to a text variable."""
        self.writer.define_text_variable(
            variable, "".join(chunk for chunk in para.text()),
        )

    def set_state(self, state: State) -> State:
        """Set current style configuration, return previous one."""
        prev = self.state
        self.state = state
        return prev

    def convert_chunk(self, chunk: IDocParagraph.Chunk) -> None:
        """Convert all text and styles in a Span."""
        if isinstance(chunk, IDocSpan):
            self.convert_span(chunk)
        elif isinstance(chunk, IDocImage):
            self.convert_image(chunk)
        elif isinstance(chunk, IDocFormula):
            self.convert_formula(chunk)
        else:
            log.warning("Unhandled chunk type %s", type(chunk))

    def convert_span(self, span: IDocSpan) -> None:
        """Convert all text and styles in a Span."""
        self.switch_character_style(self.get_character_style(span))
        self.convert_span_text(span)
        for footnote in span.footnotes():
            self.convert_footnote(footnote)
            if self.state.is_post_break:
                log.debug("Has footnote, not empty")
            self.state.is_empty = False
        if self.args.convert_comments:
            for cmt in span.comments():
                self.convert_comment(cmt)

    def get_character_style(self, span: IDocSpan) -> OptionalStyle:
        """Return Style object for a Span."""
        unadorned = self.style("character", span.style_wpid())
        if not self.args.manual and not self.args.manual_light:
            return unadorned

        # Manual formatting
        fmt = self.get_span_format(span)
        if fmt == self.state.para_char_fmt:
            # Format already included in paragraph style
            fmt = ManualFormat.NORMAL
        return self.get_manual_style("character", unadorned, fmt)

    def next_image_fn(self, infix: str, suffix: str) -> Path:
        """Generate next image filename."""
        count = next(self.image_count)
        self.image_dir.mkdir(parents=True, exist_ok=True)
        name = f"{self.output_stem}-{infix}-{count:03d}{suffix}"
        return self.image_dir / name

    def convert_image(self, img: IDocImage) -> None:
        """Save an image, keep a placeholder in the output."""
        path = self.find_image(img) or self.extract_image(img)
        self.write_image_link(path, self.image_style)

    def find_image(self, img: IDocImage) -> Path | None:
        """For an image with alt-text, try to look that up."""
        key = img.alt_text()
        if not key:
            return None

        return self.settings.find_image(key)

    def extract_image(self, img: IDocImage) -> Path:
        """Save image to file, optionally converting it."""
        suffix = img.suffix()
        path = self.next_image_fn("image", suffix)
        log.debug("Writing %s", path)
        img.save(path)

        if suffix == ".emf" and not self.args.no_emf2svg:
            svg = path.with_suffix(".svg")
            cached = self.cache.name(lambda: self.read_file(path), ".svg")
            if cached is not None and cached.is_file():
                path = self.cache.get(cached, svg)
            else:
                log.debug("Converting %s -> %s", path.name, svg.name)
                cmd = ["emf2svg-conv", "-i", str(path), "-o", str(svg)]
                subprocess.run(cmd, check=True)
                self.svg2png(svg, path)
                path = svg
                self.cache.put(path, cached)
        return path

    def write_image_link(self, path: Path, style: Style) -> None:
        """Write an image placeholder."""
        prev = self.switch_character_style(style)
        if path.is_relative_to(self.output_dir):
            self.do_write_text(str(path.relative_to(self.output_dir)))
        else:
            self.do_write_text(str(path))
        self.switch_character_style(prev)

    def convert_formula(self, formula: IDocFormula) -> None:
        """Convert a formula."""
        path = self.next_image_fn("formula", ".svg")

        cached = self.cache.name(get_contents=formula.raw, suffix=".svg")
        if cached is not None and cached.is_file():
            path = self.cache.get(cached, path)
        else:
            log.debug("Converting -> %s", path.name)
            mathml = formula.mathml()

            svg = MathConverter.mathml_to_svg(mathml, size=self.args.formula_font_size)
            with path.open("wb") as fobj:
                fobj.write(svg)

            if self.args.debug:
                with path.with_suffix(".raw").open("wb") as fobj:
                    fobj.write(formula.raw())
                with path.with_suffix(".mathml").open("w", encoding="utf-8") as fobj:
                    fobj.write(mathml)
            self.svg2png(svg, path)
            self.cache.put(path, cached)

        self.write_image_link(path, self.formula_style)

    def svg2png(self, svg: Path | bytes, path_like: Path) -> None:
        """Convert SVG to png."""
        if self.args.no_svg2png:
            return
        if isinstance(svg, Path):
            with svg.open("rb") as fobj:
                svg = fobj.read()
        path = path_like.with_suffix(".png")
        with path.open("wb") as fobj:
            png = cairosvg.svg2png(svg)
            if isinstance(png, bytes):
                fobj.write(png)
            else:
                log.warning("Cannot convert %s", path)

    def read_file(self, path: str | os.PathLike) -> bytes:
        """Read contents of a file."""
        with Path(path).open("rb") as fobj:
            return fobj.read()

    def convert_span_text(self, span: IDocSpan) -> None:
        """Convert text in a Span object."""
        switched = self.state.curr_char_style != self.state.active_char_style
        for raw in span.text():
            text = raw
            if text and not switched and self.state.is_empty:
                text = text.lstrip()
            self.write_text(text)
            if not text.isspace():
                self.state.is_empty = False

    def switch_character_style(self, style: OptionalStyle) -> OptionalStyle:
        """Set current character style."""
        prev = self.state.curr_char_style
        if style is not prev:
            self.writer.set_character_style(style)
            self.state.curr_char_style = style
        return prev

    def write_text(self, text: str) -> None:
        """Add some plain text."""
        with contextlib.ExitStack() as stack:
            stack.callback(lambda: self.do_write_text(text))
            if not self.stop_marker:
                return
            offset = text.find(self.stop_marker)
            if offset < 0:
                return
            text = text[:offset]
            raise StopMarkerFoundError("Stop marker found")

    def do_write_text(self, text: str) -> None:
        """Actually send text to `self.writer`."""
        if text:
            self.writer.write_text(text)
            self.state.active_char_style = self.state.curr_char_style

    def convert_footnote(
        self,
        footnote: IDocFootnote | IDocComment,
        ref_style: OptionalStyle = None,
        default_style: OptionalStyle = None,
    ) -> None:
        """Convert one footnote to tagged text."""
        with self.FootnoteContext(self, ref_style):
            for par in footnote.paragraphs():
                self.convert_paragraph(par, default_style=default_style)

    def convert_comment(self, cmt: IDocComment) -> None:
        """Tagged Text doesn't support notes, so convert them to footnotes."""
        if self.args.comment_prefix:
            text = "".join(
                text
                for paragraph in cmt.paragraphs()
                for span in paragraph.spans()
                for text in span.text()
            )
            if not text.startswith(self.args.comment_prefix):
                return
        self.convert_footnote(
            cmt,
            ref_style=self.comment_ref_style,
            default_style=self.comment_style,
        )

    class FootnoteContext(contextlib.AbstractContextManager):
        """Context manager for style inside footnotes."""

        def __init__(
            self,
            outer: "WordProcessorToInDesignTaggedText",
            ref_style: OptionalStyle = None,
        ) -> None:
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

        def __exit__(self, *args) -> None:
            self.outer.set_state(self.outer_state)
            self.writer.leave_footnote()
            self.writer.set_character_style(self.outer_character_style)
            self.outer.writer = self.writer

    def report_statistics(self) -> None:
        """Print a nice summary."""
        realms = {s.realm for s in self.styles.values()}
        for realm in realms:
            log.info(
                "Number of %s styles used: %u",
                realm.capitalize(),
                sum(1 for s in self.styles.values() if s.realm == realm and s.used),
            )
        for rule in self.rules:
            if rule.applied:
                log.info("%s application(s) of %s", rule.applied, rule)

    def style(self, realm: str, wpid: str | None) -> OptionalStyle:
        """Return a Style object."""
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
        section_name = self.settings.fix_section_name(
            realm=style.realm, internal_name=style.internal_name,
        )
        if section_name in self.style_sections_used:
            return

        if style.parent_style is not None:
            if not style.parent_style.used:
                log.debug(
                    "[%s] leads to missing %r",
                    section_name,
                    style.parent_wpid,
                )
                style.parent_style = self.base_styles[style.realm]
                style.parent_wpid = style.parent_style.wpid
            self.activate_style(style.parent_style)

        log.debug("Activating %s", style)
        self.settings.update_section(section_name, style)
        self.style_sections_used.add(section_name)

        if style.next_style is not None and style.next_style.used:
            self.activate_style(style.next_style)
        elif style.next_wpid:
            log.debug("[%s] leads to missing %r", section_name, style.next_wpid)

    @classmethod
    def quote_fn(cls, path: Path | str) -> str:
        """Wrap shlex.quote with Path knowledge."""
        if isinstance(path, str):
            path = Path(path)
        return shlex.quote(str(path.absolute()))
