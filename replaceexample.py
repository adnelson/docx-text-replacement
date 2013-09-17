#!/usr/bin/env python

from docxreplace import DocXReplace
import sys

if __name__ == '__main__':
    if len(sys.argv) == 4:
        input_docx = sys.argv[1]
        output_docx = sys.argv[2]
        json_filename = sys.argv[3]
        print "Input file: %s\nOutput file: %s\n JSON file: %s" %\
                 (input_docx, output_docx, json_filename)
        dx = DocXReplace(input_docx, json_file = json_filename)
        dx.replace_all()
        dx.save(output_docx)                
    else:
        print "Error, not enough arguments. Should be: <input docx> <output docx> <json file>"