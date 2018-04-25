#!/usr/bin/env python3
import attr
from wp2tt.ini import ATTR_KEY
from wp2tt.ini import ATTR_VALUE_READONLY
from wp2tt.ini import ATTR_VALUE_HIDDEN

ATTR_READONLY = {ATTR_KEY: ATTR_VALUE_READONLY}
ATTR_NO_INI = {ATTR_KEY: ATTR_VALUE_HIDDEN}


@attr.s(slots=True)
class Style(object):
    """A character/paragraph style, normally found in the input file."""
    realm = attr.ib(metadata=ATTR_NO_INI)
    wpid = attr.ib(metadata=ATTR_READONLY)  # Used for xrefs in docx; localized
    internal_name = attr.ib(metadata=ATTR_NO_INI)  # Used in the section names
    name = attr.ib()  # What the user (and InDesign) see
    parent_wpid = attr.ib(default=None, metadata=ATTR_READONLY)
    next_wpid = attr.ib(default=None, metadata=ATTR_READONLY)
    automatic = attr.ib(default=None, metadata=ATTR_READONLY)
    idtt = attr.ib(default='')
    variable = attr.ib(default=None)

    used = attr.ib(default=None, metadata=ATTR_NO_INI)
    count = attr.ib(default=0, metadata=ATTR_NO_INI)

    parent_style = attr.ib(default=None, metadata=ATTR_NO_INI)
    next_style = attr.ib(default=None, metadata=ATTR_NO_INI)

    def __str__(self):
        return '<%s %r>' % (self.realm, self.name)


@attr.s(slots=True)
class Rule(object):
    """A derivation rule for Styles."""
    mnemonic = attr.ib(metadata=ATTR_NO_INI)
    description = attr.ib(metadata=ATTR_NO_INI)
    turn_this = attr.ib(default=None)
    into_this = attr.ib(default=None)
    when_following = attr.ib(default=None)
    when_first_in_doc = attr.ib(default=None)

    turn_this_style = attr.ib(default=None, metadata=ATTR_NO_INI)
    into_this_style = attr.ib(default=None, metadata=ATTR_NO_INI)
    when_following_styles = attr.ib(default=None, metadata=ATTR_NO_INI)

    valid = attr.ib(default=True, metadata=ATTR_NO_INI)
    applied = attr.ib(default=0, metadata=ATTR_NO_INI)

    def __str__(self):
        return '%s %r' % (self.mnemonic, self.description)


@attr.s(slots=True)
class DocumentProperties(object):
    """Things we can tell about a document."""
    has_rtl = attr.ib(default=True)
    pure_ascii = attr.ib(default=False)
