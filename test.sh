#!/bin/bash

cd example
rm -rf finished finished.docx
../docxreplace.py moodys_example.docx finished.docx replace.json
unzip finished.docx -d finished
open finished.docx