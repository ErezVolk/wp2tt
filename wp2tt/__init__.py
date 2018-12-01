#!/usr/bin/env python3
"""A utility to convert word processor files (.docx, .odt) to InDesign's Tagged Text."""
import argparse
import collections
import configparser
import contextlib
import itertools
import logging
import os
import re
import shlex
import shutil
import sys

import attr

from wp2tt.version import WP2TT_VERSION
from wp2tt.ini import ini_fields
from wp2tt.styles import Style
from wp2tt.styles import Rule
from wp2tt.proxies import ByExtensionInput
from wp2tt.proxies import MultiInput
from wp2tt.output import WhitespaceStripper
from wp2tt.tagged_text import InDesignTaggedTextOutput


def main():
    WordProcessorToInDesignTaggedText().run()


class StopMarkerFound(Exception):
    """We raise this to stop the presses."""
    pass


class BadReferenceInRule(Exception):
    """We raise this for bad ruels."""
    pass


class ParseDict(argparse.Action):
    """Helper class to convert KEY=VALUE pairs to a dict."""
    def __call__(self, parser, namespace, values, option_string):
        setattr(namespace, self.dest, dict(val.split('=', 1) for val in values))


@attr.s
class State(object):
    curr_char_style = attr.ib(default=None)
    prev_para_style = attr.ib(default=None)
    curr_para_text = attr.ib(default='')
    prev_para_text = attr.ib(default=None)


class WordProcessorToInDesignTaggedText(object):
    """Read a word processor file. Write an InDesign Tagged Text file. What's not to like?"""
    SETTING_FILE_ENCODING = 'UTF-8'
    CONFIG_SECTION_NAME = 'General'
    SPECIAL_GROUP = '(autogenerated)'
    DEFAULT_BASE = SPECIAL_GROUP + '/(Basic Style)'
    FOOTNOTE_REF_STYLE = SPECIAL_GROUP + '/(Footnote Reference in Text)'
    COMMENT_REF_STYLE = SPECIAL_GROUP + '/(Comment Reference)'
    IGNORED_STYLES = {
        'character': ['annotation reference'],
    }
    SPECIAL_STYLE = {
        'character': {
            COMMENT_REF_STYLE: {
                'idtt': '<pShadingColor:Cyain><pShadingOn:1><pShadingTint:100>',
            }
        },
        'paragraph': {
            'annotation text': {
                'name': SPECIAL_GROUP + '/(Comment Text)',
                'idtt': '<cSize:6><cColor:Cyan><cColorTint:100>',
            }
        },
    }

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
        self.parser = argparse.ArgumentParser(
            description='Word Processor to InDesign Tagged Text Converter, v' + WP2TT_VERSION
        )
        self.parser.add_argument('input', help='Input .docx file')
        self.parser.add_argument('output', nargs='?',
                                 help='InDesign Tagged Text file')
        self.parser.add_argument('-a', '--append', metavar='INPUT', nargs='*',
                                 help='Concatenate more input file(s) to the same output')
        self.parser.add_argument('-s', '--stop-at', metavar='TEXT',
                                 required=False,
                                 help='Stop importing when TEXT is found')
        self.parser.add_argument('-c', '--base-character-style', metavar='NAME',
                                 default=self.DEFAULT_BASE,
                                 help='Base all character styles on this.')
        self.parser.add_argument('-p', '--base-paragraph-style', metavar='NAME',
                                 default=self.DEFAULT_BASE,
                                 help='Base all paragraph styles on this.')
        self.parser.add_argument('-v', '--style-to-variable', metavar='STYLE=VARIABLE', nargs='+',
                                 action=ParseDict,
                                 help='Map paragraph styles to document variables.')
        self.parser.add_argument('-f', '--fresh-start', action='store_true',
                                 help='Do not read any existing settings.')
        self.parser.add_argument('-d', '--debug', action='store_true',
                                 help='Print interesting debug information.')
        self.parser.add_argument('-C', '--convert-comments', action='store_true',
                                 help='Convert comments to balloons.')
        self.parser.add_argument('--no-rerunner', action='store_true',
                                 help='Do not (over)write the rerruner script.')
        self.args = self.parser.parse_args()

        if self.args.output:
            self.output_fn = self.args.output
        else:
            basename, dummy_ext = os.path.splitext(self.args.input)
            self.output_fn = basename + '.tagged.txt'

        self.settings_fn = self.output_fn + '.ini'
        self.rerunner_fn = self.output_fn + '.rerun'
        self.stop_marker = self.args.stop_at

    def configure_logging(self):
        """Set logging level and format."""
        logging.basicConfig(
            format='%(asctime)s %(message)s',
            level=logging.DEBUG if self.args.debug else logging.INFO
        )

    def read_settings(self):
        """Read and parse the ini file."""
        self.settings = configparser.ConfigParser()
        self.settings_touched = False
        self.style_sections_used = set()
        if os.path.isfile(self.settings_fn) and not self.args.fresh_start:
            logging.info('Reading %r', self.settings_fn)
            self.settings.read(self.settings_fn, encoding=self.SETTING_FILE_ENCODING)

        self.load_rules()
        self.config = self.ensure_setting_section(self.CONFIG_SECTION_NAME)
        if self.stop_marker:
            self.config['stop_marker'] = self.stop_marker
        else:
            self.stop_marker = self.config.get('stop_marker')

    def load_rules(self):
        """Convert Rule sections into Rule objects."""
        self.rules = []
        for section_name in self.settings.sections():
            if not section_name.lower().startswith('rule:'):
                continue
            section = self.settings[section_name]
            self.rules.append(Rule(
                mnemonic='R%s' % (len(self.rules) + 1),
                description=section_name[5:],
                **{
                    name: section[ini_name]
                    for name, ini_name in ini_fields(Rule)
                    if ini_name in section
                }
            ))
            logging.debug(self.rules[-1])

    def write_settings(self):
        """When done, write the settings file for the next time."""
        if self.settings_touched and os.path.isfile(self.settings_fn):
            logging.debug('Backing up %r', self.settings_fn)
            shutil.copy(self.settings_fn, self.settings_fn + '.bak')

        logging.info('Writing %r', self.settings_fn)
        with open(self.settings_fn, 'w', encoding=self.SETTING_FILE_ENCODING) as fo:
            self.settings.write(fo)

    def write_rerunner(self):
        """Write a script to regenerate the output."""
        if self.args.no_rerunner:
            return

        logging.info('Writing %r', self.rerunner_fn)
        with open(self.rerunner_fn, 'w', encoding=self.SETTING_FILE_ENCODING) as fo:
            fo.write(
                '#!/bin/bash\n'
                '# AUTOGENERATED FILE, DO NOT EDIT.\n'
                '\n'
            )
            cli = [
                shlex.quote(os.path.abspath(sys.argv[0])),
                shlex.quote(os.path.abspath(self.args.input)),
            ]
            if self.args.output:
                cli.append(shlex.quote(os.path.abspath(self.output_fn)))
            cli.append('"$@"')  # Has to come before the dashes
            if self.stop_marker:
                cli.extend(['--stop-at', shlex.quote(self.stop_marker)])
            if self.args.base_character_style != self.DEFAULT_BASE:
                cli.extend(['--base-character-style', shlex.quote(self.args.base_character_style)])
            if self.args.base_paragraph_style != self.DEFAULT_BASE:
                cli.extend(['--base-paragraph-style', shlex.quote(self.args.base_paragraph_style)])
            if self.args.style_to_variable:
                cli.append('--style-to-variable')
                cli.extend(
                    shlex.quote('%s=%s' % (k, v))
                    for k, v in self.args.style_to_variable.items()
                )
            if self.args.debug:
                cli.append('--debug')
            if self.args.append:
                cli.append('--append')
                cli.extend([
                    shlex.quote(os.path.abspath(path))
                    for path in self.args.append
                ])
            cli.extend(['2>&1', '|tee', os.path.abspath(self.rerunner_fn + '.output')])
            fo.write(' '.join(cli))
            fo.write('\n')
        os.chmod(self.rerunner_fn, 0o755)

    def read_input(self):
        """Unzip and parse the input files."""
        logging.info('Reading %r', self.args.input)
        with self.create_reader() as self.doc:
            self.scan_style_definitions()
            self.scan_style_mentions()
        self.link_styles()
        self.link_rules()

    def create_reader(self):
        if self.args.append:
            return MultiInput([self.args.input] + self.args.append)
        else:
            return ByExtensionInput(self.args.input)

    def scan_style_definitions(self):
        """Create a Style object for everything in the document."""
        self.styles = {}
        self.create_special_styles()
        counts = collections.defaultdict(lambda: itertools.count(start=1))
        for style_kwargs in self.doc.styles_defined():
            if style_kwargs.get('automatic'):
                style_kwargs['name'] = '%s/automatic-%u' % (
                    self.SPECIAL_GROUP, next(counts[style_kwargs['realm']])
                )
            self.found_style_definition(**style_kwargs)

    def create_special_styles(self):
        """Add any internal styles (i.e., not imported from the doc)."""
        self.base_names = {
            'character': self.args.base_character_style,
            'paragraph': self.args.base_paragraph_style,
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
            realm='character',
            internal_name=self.FOOTNOTE_REF_STYLE,
            wpid=self.FOOTNOTE_REF_STYLE,
            idtt='<cColor:Magenta><cColorTint:100><cPosition:Superscript>',
            automatic=True,
        )
        self.comment_ref_style = self.found_style_definition(
            realm='character',
            internal_name=self.COMMENT_REF_STYLE,
            wpid=self.COMMENT_REF_STYLE,
            parent_wpid=self.FOOTNOTE_REF_STYLE,
            idtt='<cColor:Cyan><cColorTint:100>',
            automatic=True,
        )

    def scan_style_mentions(self):
        """Mark which styles are actually used."""
        for realm, wpid in self.doc.styles_in_use():
            style_key = self.style_key(realm=realm, wpid=wpid)
            if style_key not in self.styles:
                logging.debug('Used but not defined? %r', style_key)
            elif not self.styles[style_key].used:
                logging.debug('Style used: %r', style_key)
                self.styles[style_key].used = True

    def link_styles(self):
        """A sort of alchemy-relationship thing."""
        for style in self.styles.values():
            style.parent_style = self.style_or_none(style.realm, style.parent_wpid)
            style.next_style = self.style_or_none(style.realm, style.next_wpid)

    def style_or_none(self, realm, wpid):
        if not wpid:
            return None
        return self.styles[self.style_key(realm=realm, wpid=wpid)]

    def link_rules(self):
        """A sort of alchemy-relationship thing."""
        for rule in self.rules:
            try:
                rule.turn_this_style = self.find_style_by_ini_ref(
                    rule.turn_this,
                    required=True
                )
                rule.into_this_style = self.find_style_by_ini_ref(
                    rule.into_this,
                    required=True,
                    inherit_from=rule.turn_this_style,
                )
                if rule.when_following is not None:
                    rule.when_following_styles = [
                        self.find_style_by_ini_ref(ini_ref)
                        for ini_ref in
                        re.findall(r'\[.*?\]', rule.when_following)
                    ]
            except BadReferenceInRule:
                logging.warn('Ignoring rule with bad references: %s', rule)
                rule.valid = False

    def find_style_by_ini_ref(self, ini_ref, required=False, inherit_from=None):
        """Returns a style, given type of string we use for ini file section names."""
        if not ini_ref:
            if required:
                logging.debug('MISSING REQUIRED SOMETHING')
                raise BadReferenceInRule()
            return None
        m = re.match(
            r'^\[(?P<realm>\w+):(?P<internal_name>.+)\]$',
            ini_ref,
            re.IGNORECASE
        )
        if not m:
            logging.debug('Malformed %r', ini_ref)
            raise BadReferenceInRule()
        try:
            realm = m.group('realm').lower()
            internal_name = m.group('internal_name')
            return next(
                style for style in self.styles.values()
                if style.realm.lower() == realm and style.internal_name == internal_name
            )
        except StopIteration:
            if not inherit_from:
                logging.debug('ERROR: Unknown %r', ini_ref)
                raise BadReferenceInRule()
        return self.add_style(
            realm=realm,
            wpid=ini_ref,
            internal_name=internal_name,
            parent_wpid=inherit_from.wpid,
            parent_style=inherit_from,
        )

    def found_style_definition(self, realm, internal_name, wpid, **kwargs):
        if realm not in self.base_names:
            logging.debug('What about %s:%r [%r]?', realm, wpid, internal_name)
            return

        if wpid != self.base_names.get(realm):
            kwargs.setdefault('parent_wpid', self.base_names.get(realm))

        # Allow any special overrides (color, name, etc.)
        kwargs.update(self.SPECIAL_STYLE.get(realm, {}).get(internal_name, {}))

        section = self.get_setting_section(realm=realm, internal_name=internal_name)
        if not kwargs.get('name'):
            kwargs['name'] = internal_name
        kwargs.update({
            name: section[ini_name]
            for name, ini_name in ini_fields(Style, writeable=True)
            if ini_name in section
        })

        return self.add_style(
            realm=realm,
            wpid=wpid,
            internal_name=internal_name,
            **kwargs
        )

    def add_style(self, **kwargs):
        """Create a new Style object and add to our map."""
        if self.args.style_to_variable:
            if kwargs['realm'] == 'paragraph':
                kwargs.setdefault('variable', self.args.style_to_variable.get(kwargs['internal_name']))

        kwargs.setdefault('name', kwargs['internal_name'])

        style = Style(**kwargs)
        self.styles[self.style_key(style=style)] = style
        return style

    def style_key(self, style=None, realm=None, wpid=None):
        """The string which we use for `self.styles`.

        Note that this is based on the wpid, because
        that's the key docx files use for cross-reference.
        """
        if style:
            realm = style.realm
            wpid = style.wpid
        return '%s:%s' % (realm, wpid)

    def section_name(self, realm=None, internal_name=None, style=None):
        """The name of the ini section for a given style.

        This uses `internal_name`, rather than `wpid` or `name`,
        because `wpid` can get ugly ("a2") and `name` should be
        modifyable.
        """
        if style:
            realm = style.realm
            internal_name = style.internal_name
        return '%s:%s' % (realm.capitalize(), internal_name)

    def write_idtt(self):
        logging.info('Writing %r', self.output_fn)
        self.set_state(State())
        with InDesignTaggedTextOutput(self.output_fn, self.args.debug, self.doc.properties) as self.writer:
            self.convert_document()

    def convert_document(self):
        try:
            self.stop_marker_found = False
            for p in self.doc.paragraphs():
                self.convert_paragraph(p)
            if self.stop_marker:
                logging.info('Note: Stop marker was never found')
                logging.debug('In other words, no %r', self.stop_marker)
        except StopMarkerFound as marker:
            logging.info(marker)

    def convert_paragraph(self, p):
        if self.stop_marker:
            self.check_for_stop_paragraph(p)
        style = self.apply_rules_to(self.style('paragraph', p.style_wpid()))
        with self.ParagraphContext(self, style):
            for r in p.spans():
                self.convert_range(r)
            if style and style.variable:
                self.define_variable_from_paragraph(style.variable, p)

    def apply_rules_to(self, style):
        for rule in self.rules:
            if self.rule_applies_to(rule, style):
                rule.applied += 1
                return rule.into_this_style
        return style

    def rule_applies_to(self, rule, style):
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

    def check_for_stop_paragraph(self, p):
        text = ''
        for chunk in p.text():
            text += chunk
            if text.startswith(self.stop_marker):
                raise StopMarkerFound('Stop marker found at the beginning of a paragraph')
            if len(text) >= len(self.stop_marker):
                return

    def define_variable_from_paragraph(self, variable, p):
        self.writer.define_text_variable(variable, ''.join(chunk for chunk in p.text()))

    def set_state(self, state):
        prev = getattr(self, 'state', None)
        self.state = state
        return prev

    class ParagraphContext(contextlib.AbstractContextManager):
        def __init__(self, outer, style):
            super().__init__()
            self.outer = outer
            self.writer = self.outer.writer
            self.style = style
            self.writer.enter_paragraph(self.style)

        def __exit__(self, *args):
            self.writer.leave_paragraph()
            self.outer.set_state(State(
                prev_para_style=self.style,
                prev_para_text=self.outer.state.curr_para_text,
            ))

    def convert_range(self, r):
        self.switch_character_style(self.style('character', r.style_wpid()))
        self.convert_range_text(r)
        for fn in r.footnotes():
            self.convert_footnote(fn)
        if self.args.convert_comments:
            for cmt in r.comments():
                self.convert_comment(cmt)

    def convert_range_text(self, r):
        for t in r.text():
            self.write_text(t)

    def switch_character_style(self, style):
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
            raise StopMarkerFound('Stop marker found')

    def do_write_text(self, text):
        if not text:
            return
        self.writer.write_text(text)
        self.state.curr_para_text += text

    def convert_footnote(self, fn, ref_style=None):
        with self.FootnoteContext(self, ref_style):
            for p in fn.paragraphs():
                self.convert_paragraph(p)

    def convert_comment(self, cmt):
        """Tagged Text doesn't support notes, so we convert them to footnotes."""
        return self.convert_footnote(cmt, ref_style=self.comment_ref_style)

    class FootnoteContext(contextlib.AbstractContextManager):
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

    def report_statistics(self):
        realms = {s.realm for s in self.styles.values()}
        for realm in realms:
            logging.info('Number of %s styles used: %u',
                         realm.capitalize(),
                         sum(1 for s in self.styles.values() if s.realm == realm and s.used))
        for rule in self.rules:
            if rule.applied:
                logging.info('%s application(s) of %s', rule.applied, rule)

    def style(self, realm, wpid):
        if not wpid:
            return None

        style = self.styles[self.style_key(realm=realm, wpid=wpid)]
        if realm in self.IGNORED_STYLES:
            if style.internal_name in self.IGNORED_STYLES[realm]:
                return None

        self.activate_style(style)
        style.count += 1
        return style

    def activate_style(self, style):
        section_name = self.section_name(style.realm, style.internal_name)
        if section_name in self.style_sections_used:
            return

        if style.parent_style:
            if not style.parent_style.used:
                logging.debug('[%s] leads to missing %r', section_name, style.parent_wpid)
                style.parent_style = self.base_styles[style.realm]
                style.parent_wpid = style.parent_style.wpid
            self.activate_style(style.parent_style)

        self.update_setting_section(section_name, style)
        self.style_sections_used.add(section_name)

        if style.next_style and style.next_style.used:
            self.activate_style(style.next_style)
        elif style.next_wpid:
            logging.debug('[%s] leads to missing %r', section_name, style.next_wpid)

    def get_setting_section(self, section_name=None, realm=None, internal_name=None, style=None):
        """Return a section from the ini file.

        If no such section exists, return a new dict, but don't add it
        to the configuration.
        """
        if not section_name:
            section_name = self.section_name(realm=realm, internal_name=internal_name, style=style)

        if self.settings.has_section(section_name):
            return self.settings[section_name]
        return {}

    def ensure_setting_section(self, section_name=None, realm=None, internal_name=None, style=None):
        """Return a section from the ini file, adding one if needed."""
        if not section_name:
            section_name = self.section_name(realm=realm, internal_name=internal_name, style=style)
        if not self.settings.has_section(section_name):
            self.settings[section_name] = {}
        return self.settings[section_name]

    def update_setting_section(self, section_name, style):
        """Update key:value pairs in an ini section.

        Sets `self.settings_touched` if anything was changed.
        """
        section = self.ensure_setting_section(section_name)
        for name, ini_name in ini_fields(Style):
            value = getattr(style, name)
            self.settings_touched |= section.get(ini_name) != value
            if value:
                section[ini_name] = str(value or '')
            else:
                section.pop(ini_name, None)


# TODO:
# - DOCX: <w:bookmarkStart w:name="X"/> / <w:instrText>PAGEREF ..
# - IDTT:
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
# - Manual format: autogenerate styles
# - Manual format: collapse with existing styles
# - A flag to only create/update the ini file
# - Maybe add front matter (best done in Id? either that or the template thingy (ooh, jinja2!))
# - Something usable for those balloons (footnote+hl? endnote? convert to note in jsx?)
# - Autocreate character styles from manual combos
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


