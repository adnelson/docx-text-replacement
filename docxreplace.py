#!/usr/bin/env python
from docx import DocX
import sys 
import re, random, string
import json

VERBOSE = True

def replace_tags(line, replacements, specific_words = None):
    subs = re.split(r'(@[^@]*@)', line)
    res = ""
    count = 0
    for sub in subs:
        if is_key(sub):
            # if we've given a specific word list, and this isn't in it:
            if specific_words and sub[1:-1] not in specific_words:
                res += sub # just append as-is and continue
                continue
            try:
                key = sub[1:-1]
                if VERBOSE: 
                    print "replacing '%s' with '%s'" % (key, replacements[key])
                res += replacements[key].__str__()
                count += 1
            except KeyError:
                #if it's not in our lookup table, append as-is
                print "Key '%s' not found in replacements!" % sub
                res += sub
        else:
            res += sub
    return res, count

def is_key(string):
    return len(string) > 2 and string[0] == string[-1] == '@'

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

def get_replacements(filename):
    f = open(filename)
    return json.loads(f.read())
    return text, tables, images

def process_file(input_filename, 
                 json_filename,
                 output_filename):
    ''' Given a .docx filename, makes replacements and saves the document '''
    dx = DocX(input_filename)
    dx.load_replacements(json_file = json_filename)
    print "replacing text..."
    dx.make_replacements(dx.text_reps)
    print "finished replacing text, next replacing tables"
    dx.replace_tables()
    print "finished replacing tables, next images"
    dx.replace_images()
    print "done, saving..."
    dx.save(output_filename)

if __name__ == '__main__':
    if len(sys.argv) == 4:
        input_docx = sys.argv[1]
        output_docx = sys.argv[2]
        json_filename = sys.argv[3]
        print "Input file: %s\nOutput file: %s\n JSON file: %s" %\
                 (input_docx, output_docx, json_filename)
        process_file(input_docx, json_filename, output_docx)
    else:
        print "Error, not enough arguments. Should be: <input docx> <output docx> <json file>"