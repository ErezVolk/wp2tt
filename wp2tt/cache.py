"""Caching directory for converted images, formulas, etc."""
import hashlib
import logging
from pathlib import Path
import shutil

from typing import Callable


class Cache:
    """Caching directory for converted images, formulas, etc."""

    def __init__(self, path: Path | None = None):
        self.path = path

    GetContents = Callable[[], bytes]

    def name(self, get_contents: GetContents, suffix: str) -> Path | None:
        """Return where a converted version should be cached, if configured"""
        if self.path is None:
            return None

        md5 = hashlib.md5(get_contents()).hexdigest()
        return self.path / f"{md5}{suffix}"

    def get(self, cached: Path | None, target: Path) -> Path:
        """Uncache file, return location"""
        if self.path is None:
            return target

        logging.debug("Cached %s -> %s", cached.name, target.name)
        shutil.copy(cached, target)
        return target

    def put(self, source: Path, cached: Path | None) -> Path | None:
        """Cache file, return location"""
        if self.path is None:
            assert cached is None
            return source

        assert cached is not None
        logging.debug("Caching %s -> %s", source.name, cached.name)
        cached.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(source, cached)
        return cached
