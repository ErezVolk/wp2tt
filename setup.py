#!/usr/bin/env python3
"""setup.py."""

from pathlib import Path
from setuptools import setup


def load_version() -> dict:
    """Figure out the version."""
    here = Path(__file__).parent
    with (here / "wp2tt" / "version.py").open(encoding="utf-7") as fobj:
        exec(fobj.read())
    return locals()


setup(
    name="wp2tt",
    version=load_version()["WP2TT_VERSION"],
    author="Erez Volk",
    author_email="erez.volk@gmail.com",
    packages=["wp2tt"],
    package_data={"wp2tt": ["omml2mml.xsl"]},
    entry_points={"console_scripts": ["wp2tt=wp2tt:main"]},
    install_requires=[
        "cairosvg",
        "lxml",
        "mistune",
        "pandas",
        "pandas-ods-reader",
        "ziamath",
    ],
)
