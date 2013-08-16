#!/bin/bash

rm -rf experimenting/fail*
./at_replace.py experimenting/orig_awesome.docx experimenting/fail.docx
cd experimenting
unzip fail.docx -d fail