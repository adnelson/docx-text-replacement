#!/usr/bin/env python
from docx import DocX
import sys 
import re, random, string
import json

class DocXReplace(DocX):
    def __init__(self, input_filename, json_file = None, 
                       jsonstr = None, dic = None):
        super(DocXReplace, self).__init__(input_filename)
        if json_file is not None:
            f = open(json_file)
            self.replacements = json.loads(f.read())
            f.close()
        elif jsonstr is not None:
            self.replacements = json.loads(jsonstr)
        elif dic is not None:
            self.replacements = dic
        else:
            raise Exception("No data supplied to load_replacements")

        self.text_reps = self.replacements.get("text", {})
        self.table_reps = self.replacements.get("tables", {})
        self.image_reps = self.replacements.get("images", {})

    def replace_tags(self, line, replacements, specific_words = None):
        subs = re.split(r'(@[^@]*@)', line)
        res = ""
        count = 0
        for sub in subs:
            if self.is_key(sub):
                # if we've given a specific word list, and this isn't in it:
                if specific_words and sub[1:-1] not in specific_words:
                    res += sub # just append as-is and continue
                    continue
                try:
                    key = sub[1:-1]
                    self.log("replacing '%s' with '%s'" % (key, replacements[key]))
                    res += replacements[key].__str__()
                    count += 1
                except KeyError:
                    #if it's not in our lookup table, append as-is
                    if self.verbose:
                        print "Key '%s' not found in replacements!" % sub
                    res += sub
            else:
                res += sub
        return res, count

    def is_key(self, string):
        return len(string) > 2 and string[0] == string[-1] == '@'

    def replace_text(self, replacements = None, specific_words = None):
        ''' Finds and makes all of the replacements. '''
        if replacements is None:
            if self.text_reps is not None:
                replacements = self.text_reps
            else:
                raise Exception("No text replacements defined")
        document = self.get_document()
        count = 0
        for elem in document.iter():
            if elem.text:
                elem.text, c = self.replace_tags(elem.text, replacements, specific_words)
                count += c
        self.log("Made %d replacements" % count)


def process_file(input_filename, 
                 json_filename,
                 output_filename):
    ''' Given a .docx filename, makes replacements and saves the document '''
    dx = DocXReplace(input_filename, json_file = json_filename)
    dx.verbose_only = True
    print "replacing text..."
    dx.replace_text()
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