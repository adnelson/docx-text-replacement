#!/usr/bin/env python
from docx import DocX
from at_replace import replace_tags

doc = DocX()

title    = 'Python docx demo'
subject  = 'A practical example of making docx from Python'
creator  = 'Mike MacCana'
keywords = ['python', 'Office Open XML', 'Word']

doc.set_title(title)
doc.set_subject(subject)
doc.set_keywords(keywords)
doc.set_creator(creator)

doc.save("foo.docx")