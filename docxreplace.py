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