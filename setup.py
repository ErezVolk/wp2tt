#!/usr/bin/env python3
import os
from setuptools import setup


def load_version():
    HERE = os.path.dirname(os.path.realpath(__file__))
    with open(os.path.join(HERE, 'wp2tt', 'version.py')) as vf:
        exec(vf.read())
    return locals()


setup(
    name='wp2tt',
    version=load_version()['WP2TT_VERSION'],

    author='Erez Volk',
    author_email='erez.volk@gmail.com',

    packages=['wp2tt'],
    entry_points={
        'console_scripts': ['wp2tt=wp2tt:main']
    },

    install_requires=[
        'attrs',
        'lxml',
    ]
)
