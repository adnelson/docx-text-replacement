#!/bin/bash

# rm -rf experimenting/addtableflag*
./at_replace.py experimenting/addtableflag.docx experimenting/addtableflag.docx
cd experimenting
unzip addtableflag.docx -d addtableflag