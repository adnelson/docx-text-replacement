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

def process_file(filename, replacements = None, output = None, save = True):
    ''' Given a .docx filename, makes replacements and saves the document '''
    dx = DocX(filename)
    text = replacements.get("text", {})
    tables = replacements.get("tables", {})
    images = replacements.get("images", {})
    make_replacements(dx, text)
    dx.fill_tables(tables)
    dx.replace_images_from_dic(images)
    if save: dx.save(output)

if __name__ == '__main__':
    if len(sys.argv) == 4:
        input_docx = sys.argv[1]
        output_docx = sys.argv[2]
        json_data = sys.argv[3]
        print "Input file: %s\nOutput file: %s\n JSON file: %s" % (input_docx, output_docx, json_data)
        process_file(input_docx,
                     replacements = get_replacements(json_data),
                     output = output_docx)
    else:
        print "Error, not enough arguments. Should be: <input docx> <output docx> <json file>"