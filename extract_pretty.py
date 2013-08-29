#!/usr/bin/env python
import sys
from docx import DocX 

def process_file(filename, replacements = None, output = None, save = True):
    ''' Given a .docx filename, makes replacements and saves the document '''

if __name__ == '__main__':
    if len(sys.argv) == 3:
        dx = DocX(sys.argv[1])
        dx.save(sys.argv[2])
    else:
        print "Please enter an input and output filename"