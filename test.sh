#!/bin/bash

rm -r experimenting/modified*
./at_replace.py experimenting/original.docx experimenting/modified.docx
cd experimenting
unzip modified.docx -d modified