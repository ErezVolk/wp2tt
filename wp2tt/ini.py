"""Ini file helper"""
import configparser
import logging
from pathlib import Path
import shutil

from typing import Iterable

import attr

from wp2tt.styles import ATTR_KEY
from wp2tt.styles import ATTR_VALUE_READONLY
from wp2tt.styles import ATTR_VALUE_HIDDEN
from wp2tt.styles import OptionalStyle
from wp2tt.styles import Style


ConfigSection = (configparser.SectionProxy | dict[str, str])


class SettingsFile(configparser.ConfigParser):
    """Settings .ini file"""
    ENCODING = "UTF-8"
    IMAGE_SECTION = "Images"
    BASE_KEY = "__base__"

    path: Path
    touched: bool = False
    images: configparser.SectionProxy | None = None
    base: Path

    def __init__(self, path: Path, fresh_start: bool = False):
        """Read ini file"""
        super().__init__()
        self.path = path
        self.base = self.path.parent.resolve()

        if fresh_start or not self.exists():
            return

        logging.info("Reading %s", self.path)
        self.read(path, encoding=self.ENCODING)
        try:
            self.images = self[self.IMAGE_SECTION]
            self.base = self.base / self.images[self.BASE_KEY]
            logging.debug("Image base: %s", self.base)
        except KeyError:
            pass

    def exists(self) -> bool:
        """Does the .ini file currently exist"""
        return self.path.is_file()

    def find_image(self, key) -> Path | None:
        """Look for an image"""
        try:
            return self.base / self.images[key]
        except (KeyError, TypeError):
            return None

    def fields(self, klass, writeable=False) -> Iterable[tuple[str, str]]:
        """Yields a pair (name, ini_name) for all attributes."""
        for field in attr.fields(klass):
            special = field.metadata.get(ATTR_KEY)
            if special == ATTR_VALUE_HIDDEN:
                continue
            ini_name = name = field.name
            if special == ATTR_VALUE_READONLY:
                if writeable:
                    continue
                ini_name += " (readonly)"
            yield (name, ini_name)

    def backup_and_write(self):
        """Write to disk, backing up first if modified"""
        if self.touched and self.exists():
            logging.debug("Backing up %s", self.path)
            shutil.copy(self.path, self.path.with_suffix(".bak"))

        logging.info("Writing %s", self.path)
        with open(self.path, "w", encoding=self.ENCODING) as fobj:
            self.write(fobj)

    def get_section(
        self,
        section_name: str | None = None,
        realm: str | None = None,
        internal_name: str | None = None,
        style: OptionalStyle = None,
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

        if self.has_section(actual_name):
            return self[actual_name]
        return {}

    def ensure_section(
        self,
        section_name: str | None = None,
        realm: str | None = None,
        internal_name: str | None = None,
        style: OptionalStyle = None,
    ) -> ConfigSection:
        """Return a section from the ini file, adding one if needed."""
        actual_name = self.fix_section_name(
            section_name=section_name,
            realm=realm,
            internal_name=internal_name,
            style=style,
        )
        if not self.has_section(actual_name):
            self[actual_name] = {}
        return self[actual_name]

    def update_section(self, section_name: str, style: Style) -> None:
        """Update key:value pairs in an ini section.

        Sets `self.touched` if anything was changed.
        """
        section = self.ensure_section(section_name)
        for name, ini_name in self.fields(Style):
            value = getattr(style, name)
            self.touched |= section.get(ini_name) != value
            if value:
                section[ini_name] = str(value or "")
            else:
                section.pop(ini_name, None)

    @classmethod
    def fix_section_name(
        cls,
        section_name: str | None = None,
        realm: str | None = None,
        internal_name: str | None = None,
        style: OptionalStyle = None,
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
