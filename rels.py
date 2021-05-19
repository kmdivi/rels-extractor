#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import xml.etree.ElementTree as et
import zipfile


def rels_extract(f):
    d = zipfile.ZipFile(f)
    try:
        xml = d.read("word/_rels/settings.xml.rels")
        d.close()
        root = et.fromstring(xml)
        for i in root:
            if "Target" in i.attrib:
                return i.attrib["Target"]
            return "URL Not found"
    except:
        return "word/_rels/settings.xml.rels Not found"


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage:\n\tpython3 rels.py [file.docx]")
        exit(0)

    file = sys.argv[1]
    print(rels_extract(file))
