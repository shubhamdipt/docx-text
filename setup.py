# Release process setup see:
# https://github.com/pypa/twine
#
# Upgrade twine
#     python3 -m pip install --user --upgrade twine
#
# Run this to build the `dist/PACKAGE_NAME-xxx.tar.gz` file
#     rm -rf ./dist && python3 setup.py sdist
#
# Check dist/*
#     python3 -m twine check dist/*
#
# Run this to build & upload it to `pypi`, type your account name when prompted.
#     python3 -m twine upload dist/*
#
# In one command line:
#     rm -rf ./dist && python3 setup.py sdist bdist_wheel && python3 -m twine check dist/*
#     rm -rf ./dist && python3 setup.py sdist bdist_wheel && python3 -m twine upload dist/*
#

from setuptools import setup, find_packages

# Usage: python setup.py sdist bdist_wheel

links = []  # for repo urls (dependency_links)

DESCRIPTION = "A python module for converting docx files from SharePoint, One Drive to text"
VERSION = "0.1.2"


with open("README.md", "r") as fh:
    LONG_DESCRIPTION = fh.read()

setup(
    name="docx-text",
    version=VERSION,
    author="Shubham Dipt",
    author_email="shubham.dipt@gmail.com",
    description=DESCRIPTION,
    long_description=LONG_DESCRIPTION,
    long_description_content_type='text/markdown',
    url="https://github.com/shubhamdipt/docx-text",
    license=open('LICENSE').read(),
    packages=['doctext'],
    platforms=["any"],
    classifiers=(
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ),
    dependency_links=links,
)