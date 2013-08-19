#!/bin/bash

rm -rf experimenting/tester1*
./at_replace.py experimenting/blibber.docx experimenting/tester1.docx
cd experimenting
unzip tester1.docx -d tester1