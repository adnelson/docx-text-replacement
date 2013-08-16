#!/bin/bash

rm -rf experimenting/table*
./at_replace.py experimenting/table.docx experimenting/table2.docx
cd experimenting
unzip table2.docx -d table2