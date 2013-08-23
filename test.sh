#!/bin/bash

cd experimenting
rm -rf minimal1*
../docxreplace.py mock.docx minimal1.docx mock.json
unzip minimal1.docx -d minimal1