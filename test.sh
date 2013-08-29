#!/bin/bash

cd experimenting/graytv
rm -rf graytv.docx graytv
../../docxreplace.py moodys_june.docx graytv.docx replace.json
unzip graytv.docx -d graytv