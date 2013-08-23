#!/usr/bin/env python
import sys
from docx import DocX 
import re
import random
import string

VERBOSE = True

lorem_ipsum = ("Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do " + 
              "eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim " +
              "ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip " +
              "ex ea commodo consequat. Duis aute irure dolor in reprehenderit in " +
              "voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur " +
              "sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt " +
              "mollit anim id est laborum.").split()

def replace_tags(line, replacements, specific_words = None):
    subs = re.split(r'(@[^@]*@)', line)
    res = ""
    count = 0
    for sub in subs:
        if len(sub) > 2 and sub[0] == sub[-1] == '@':
            # if we've given a specific word list, and this isn't in it:
            if specific_words and sub[1:-1] not in specific_words:
                res += sub # just append as-is and continue
                continue
            try:
                res += replacements[sub[1:-1]].__str__()
                if VERBOSE: "replacing '%s' with '%s'" % (sub, replacements[sub[1:-1]])
                count += 1
            except KeyError:
                #if it's not in our lookup table, append as-is
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

def make_replacements(dx, replacements, specific_words = None):
    ''' Finds and makes all of the replacements. '''
    document = dx.get_document()
    count = 0
    for elem in document.iter():
        if elem.text:
            elem.text, c = replace_tags(elem.text, replacements, specific_words)
            count += c
    print "Made %d replacements" % count

def random_str(n):
    return ''.join(random.choice(string.ascii_uppercase + string.digits) for x in range(n))


table_replacements = {
    "TABLE1": [[random.choice(lorem_ipsum) for j in range(7)] for i in range(5)],
    "TABLE2": [[random.choice(lorem_ipsum) for j in range(7)] for i in range(3)]
}

def process_file(filename, replacements = None, output = None, save = True):
    ''' Given a .docx filename, makes replacements and saves the document '''
    # try:
    dx = DocX(filename)
    # dx.fill_tables(table_replacements)
    # if replacements is None:
    #     replacements = generate_random(dx.get_document())
    # make_replacements(dx, replacements)
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