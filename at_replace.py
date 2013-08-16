#!/usr/bin/env python
import sys
from docx import DocX 
import re
import random

VERBOSE = True

lorem_ipsum = ("Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do " + 
              "eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim " +
              "ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip " +
              "ex ea commodo consequat. Duis aute irure dolor in reprehenderit in " +
              "voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur " +
              "sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
              "mollit anim id est laborum.").split()

def replace_tags(line, replacements, specific_words = None):
    subs = re.split(r'(@[^@<]*@?)', line)
    res = ""
    count = 0
    for sub in subs:
        if len(sub) > 2 and sub[0] == sub[-1] == '@':
            if specific_words:
                if sub[1:-1] not in specific_words:
                    res += sub
                    continue
            try:
                res += replacements[sub[1:-1]].__str__()
                if VERBOSE: "replacing '%s' with '%s'" % (sub, replacements[sub[1:-1]])
                count += 1
            except KeyError:
                print "Key '%s' not found in replacements!" % sub
                res += sub
        else:
            res += sub
    return res, count

def generate_random(document):
    ''' Generates random lorem ipsum rules for each @-enclosed item '''
    replacements = {}
    def gen(line):
        subs = re.split(r'(@[^@]*@)', line)
        for sub in subs:
            if len(sub) > 2:
                if sub[0] == sub[-1] == '@':
                    txt = sub[1:-1]
                    if txt not in replacements:
                        rep = " ".join(random.sample(lorem_ipsum, len(txt)))
                        replacements[txt] = rep
    for elem in document.iter():
        if elem.text:
            gen(elem.text)
    return replacements

def make_replacements(document, replacements, specific_words = None):
    ''' Finds and makes all of the replacements. '''
    count = 0
    for elem in document.iter():
        if elem.text:
            elem.text, c = replace_tags(elem.text, replacements, specific_words)
            count += c
    print "Made %d replacements" % count

def process_file(filename, replacements = None, output = None, save = True):
    ''' Given a .docx filename, makes replacements and saves the document '''
    # try:
    dx = DocX(filename)
    if replacements is None:
        replacements = generate_random(dx.get_document())
    # dx.replace_images_from_dic({"awesome.png": "more_awesome.png"})
    if save: 
        dx.save(output)
    # except Exception as e:
    #     print e

if __name__ == '__main__':
    if len(sys.argv) == 3:
        process_file(sys.argv[1], \
                     output = sys.argv[2])
    else:
        process_file("experimenting/original.docx", \
                     "experimenting/modified.docx")