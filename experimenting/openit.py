#!/usr/bin/env python
"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import sys

from docx import opendocx, getdocumenttext, search, savedocx, newdocument
from lxml import etree
import re
import random


try:
    input_file = opendocx("moodys_june-1.docx")
except:
    print(
        "Please supply an input and output file. For example:\n"
        "  example-extracttext.py 'My Office 2007 document.docx' 'outp"
        "utfile.txt'"
    )
    exit()

replacements = {}

def find_replacements(line):
    subs = re.split(r'(@[^@]*@)', line)
    res = ""
    for sub in subs:
        if len(sub) > 2:
            if sub[0] == sub[-1] == '@':
                try:
                    res += replacements[sub[1:-1]].__str__()
                    # print "replacing '%s' with '%s'" % (sub, replacements[sub[1:-1]])
                except KeyError:
                    print "Key '%s' not found in replacements!" % sub
                    res += "@%s@" % sub
            else:
                res += sub
    return res

lorem_ipsum = ("Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do " + 
              "eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim " +
              "ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip " +
              "ex ea commodo consequat. Duis aute irure dolor in reprehenderit in " +
              "voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur " +
              "sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
              "mollit anim id est laborum.").split()

def generate_random_replacements(line):
    subs = re.split(r'(@[^@]*@)', line)
    for sub in subs:
        if len(sub) > 2:
            if sub[0] == sub[-1] == '@':
                txt = sub[1:-1]
                if txt not in replacements:
                    rep = " ".join(random.sample(lorem_ipsum, len(txt)))
                    # print "'%s' -> '%s'" % (txt, rep)
                    replacements[txt] = rep

print find_replacements(text)

# first loop to find all the tags and generate random lorem ipsup to replace them with
# obviously this won't be in the final product...
for elem in input_file.iter():
    if elem.text:
        generate_random_replacements(elem.text)

# second loop goes and finds all of the replacements. Of course we could do this as one loop
# but w/e
for elem in input_file.iter():
    if elem.text:
        elem.text = find_replacements(elem.text)

document = newdocument()

savedocx(input_file, "foo", "bar", "baz", "qux", "blob", "moodymod.docx")

# # Fetch all the text out of the input_file we just created
# paratextlist = getdocumenttext(document)

# # Make explicit unicode version
# newparatextlist = []
# for paratext in paratextlist:
#     newparatextlist.append(paratext.encode("utf-8"))

# # Print out text of input_file with two newlines under each paragraph
# newfile.write('\n\n'.join(newparatextlist))