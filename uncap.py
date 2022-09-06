#!/bin/python3
"""
Yes, I know there are pre-commit hooks, but this is a quick and dirty script
"""

from zipfile import ZipFile
from pathlib import Path
import xml.etree.ElementTree as ET


def uncap():
    with ZipFile("collections.ods", "r") as odsfile:
        with odsfile.open("Basic/Standard/Collections.xml") as coll:
            collections = ET.parse(coll).getroot().text
        with odsfile.open("Basic/Standard/TestCollections.xml") as coll:
            test_collections = ET.parse(coll).getroot().text

    with Path("Collections.vb").open("w", encoding="utf-8") as d:
        d.write(collections)

    with Path("TestCollections.vb").open("w", encoding="utf-8") as d:
        d.write(test_collections)


if __name__ == "__main__":
    uncap()
