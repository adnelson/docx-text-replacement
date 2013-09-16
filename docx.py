#!/usr/bin/env python2.6
# -*- coding: utf-8 -*-
"""
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and
'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import logging
from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image
import zipfile
import shutil
import re
import time
import os
from os.path import join
import subprocess
import random
import json

log = logging.getLogger(__name__)

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
template_dir = join(os.path.dirname(__file__), 'docx-template')  # installed
if not os.path.isdir(template_dir):
    template_dir = join(os.path.dirname(__file__), 'template')  # dev

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
nsprefixes = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # Properties (core and extended)
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'}

def make_dummy_table(nrows, ncols, multiplier = 1):
    dummy_table = []
    for i in range(nrows * ncols):
        x = (i * multiplier).__str__()
        if i % ncols == 0:
            dummy_table.append([x])
        else:
            dummy_table[-1].append(x)
    return dummy_table

class DocX():
    def __init__(self, filename = None):
        self.text_reps = None
        self.table_reps = None
        self.image_reps = None

        self.relationships = relationshiplist()
        self.trees = {}
        self.images = {}
        self.other = {}
        if filename:
            self.filename = filename
            print "Opening file '%s'" % self.filename
            try:
                doc = zipfile.ZipFile(self.filename) 
                for name in doc.namelist():
                    if name.endswith("xml") or name.endswith("rels"):
                        print "\tAdding xml file", name, " to DocX object"
                        self.trees[name] = etree.fromstring(doc.read(name))
                    elif name.endswith("jpeg") or name.endswith("png") or name.endswith("jpg"):
                        # open the image and read its contents into memory
                        print "\tAdding image: %s" % (name)
                        self.images[name] = doc.read(name)
                    else:
                        print "\tFound a file %s that we're not doing anything with" % name
                        self.other[name] = doc.read(name)
            except Exception as e:
                print e
                raise
        else:
            self.trees['word/document.xml'] = newdocument()
            self.trees['docProps/core.xml'] = None # modify this later
            self.trees['docProps/app.xml'] = appproperties()
            self.trees['[Content_Types].xml'] = contenttypes()
            self.trees['word/webSettings.xml'] = websettings()
            self.trees['word/_rels/document.xml.rels'] = wordrelationships(self.relationships)
        self.body = self.get_document().xpath("/w:document/w:body", namespaces = nsprefixes)
        self.core_props = {}

    def set_title(self, title):
        self.get_core_props()['title'] = title

    def set_subject(self, subject):
        self.get_core_props()['subject'] = subject

    def set_creator(self, creator):
        self.get_core_props()['creator'] = creator

    def set_keywords(self, keywords):
        self.get_core_props()['keywords'] = keywords

    def get_document(self):
        return self.trees['word/document.xml']

    def get_core_props(self):
        return self.core_props

    def get_relationships(self):
        return self.trees['word/_rels/document.xml.rels']

    def set_image_relation(self, rel_id, image_path):
        rels = self.trees['word/_rels/document.xml.rels']
        # read image into memory
        try:
            img = open(image_path).read()
        except Exception as e:
            print "Error opening image %s: %s" % (image_path, e)
            return
        # fix the image path (e.g. "foo/bar/baz.jpg" -> "media/baz.jpg")
        image_path = "media/" + image_path.split('/')[-1]
        self.images["word/" + image_path] = img
        for rel in rels:
            if 'Id' in rel.attrib and rel.attrib['Id'] == rel_id:
                print "%s was pointed at %s" % (rel_id, rel.attrib['Target'])
                rel.attrib['Target'] = image_path
                print 'Now %s is pointing at %s' % (rel_id, image_path)
                return
        raise Exception('Relationship ID %s was not found!' % rel_id)

    def save(self, output = None):
        '''Save a modified document'''
        assert os.path.isdir(template_dir)
        if output is None:
            output = self.filename
        docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

        # set up the core properties if not already
        if self.trees['docProps/core.xml'] is None:
            self.trees['docProps/core.xml'] = coreproperties(**self.get_core_props())

        # For some reason this version tag doesn't get appended automatically, so for the 
        # time being we're doing it manually...
        version_tag = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"

        # Serialize our trees into out zip file
        for filename in self.trees:
            log.info('Saving XML file: %s' % filename)
            treestring = version_tag + etree.tostring(self.trees[filename], pretty_print = True)
            print 'Saving %s' % (filename)
            docxfile.writestr(filename, treestring)
        for filename in self.images:
            print "Saving image: %s" % filename
            docxfile.writestr(filename, self.images[filename])
        for filename in self.other:
            print "Saving other file: %s" % filename
            docxfile.writestr(filename, self.other[filename])
        print "finished adding files. Archive now contains:"
        docxfile.printdir()
        print 'Saved to: %r' % output
        docxfile.close()

    def replace_image(self, imagename, new_image):
        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "graphic":
                rid = get_id(elem)
                picname = get_pic_name(elem)
                if picname == imagename and rid is not None:
                    self.set_image_relation(rid, new_image)
                    return

    def replace_images(self, replacements = None):
        if replacements is None:
            if self.image_reps is not None:
                replacements = self.image_reps
            else:
                raise Exception("No image replacements defined")
        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "graphic":
                picname = get_pic_name(elem)
                if picname and picname in replacements:
                    rid = get_id(elem)
                    if rid is not None:
                        print "Replacing %s with %s" % (picname, replacements[picname])
                        self.set_image_relation(rid, replacements[picname])
                    else:
                        print "Relation id for image %s not present; can't replace" % picname

    def replace_tables(self, table_replacements = None):
        if table_replacements is None:
            if self.table_reps is not None:
                table_replacements = self.table_reps
            else:
                raise Exception("no table replacements dict defined")

        for elem in self.get_document().iter():
            if elem.tag.split("}")[-1] == "tbl":
                try:
                    nrows = get_num_rows(elem)
                    ncols = get_num_columns(elem)
                    for i in range(len(elem)):
                        if elem[i].tag.split("}")[-1] == "tr":
                            rowelem = elem[i]
                            col = find_subelem_list(rowelem, ["tc", "p", "r", "t"])

                            if col is None:
                                raise Exception("Could not find column")

                            tags = re.findall(r'@@([^@]+)@@', col.text)

                            if not tags:
                                # print "No @@ tag found to describe source in row %d" % i
                                continue
                            source = tags[0]
                            print "Found table tag %s, querying dictionary" % source
                            if source not in table_replacements:
                                raise Exception("Error: couldn't find %s in replacements dict" % source)
                            settings = table_replacements[source][0]
                            font_size = settings.get("font_size", None)
                            # hack because fonts appear half-size for some reason
                            if font_size is not None:
                                font_size *= 2
                            font_face = settings.get("font_face", None)
                            # get the border settings
                            borders = settings.get("borders", [])

                            under_border = settings.get("under_border", False)
                            content = table_replacements[source][1]
                            
                            tbl_ncols = len(content[0])
                            if tbl_ncols != ncols:
                                raise Exception("Error: should have %d columns, but "
                                                "source has %d columns" % (ncols, tbl_ncols))
                            first = True
                            j = 0
                            for row in content:
                                if first:
                                    elem[i] = make_row(row, font_face=font_face, 
                                                            font_size=font_size.__str__(),
                                                            borders = borders)
                                    first = False
                                else:
                                    elem.append(make_row(row, font_face=font_face,
                                                              font_size=font_size.__str__(),
                                                              borders = borders))
                                j += 1
                            print "Inserted %d rows into table %s" % (j, source)
                            break # only do it once for each table
                except Exception as e:
                    print e
                    print "Error reading or constructing table element, no rows added"
                    return

    def load_replacements(self, json_file = None, jsonstr = None, dic = None):
        ''' Loads a file or string containing JSON into a python dictionary,
            or can be supplied a dictionary directly
        '''
        if json_file is not None:
            f = open(json_file)
            self.replacements = json.loads(f.read())
            f.close()
        elif jsonstr is not None:
            self.replacements = json.loads(jsonstr)
        elif dic is not None:
            self.replacements = dic
        else:
            raise Exception("No data supplied to load_replacements")

        self.text_reps = self.replacements.get("text", {})
        self.table_reps = self.replacements.get("tables", {})
        self.image_reps = self.replacements.get("images", {})

    def replace_tags(self, line, replacements, specific_words = None):
        subs = re.split(r'(@[^@]*@)', line)
        res = ""
        count = 0
        for sub in subs:
            if self.is_key(sub):
                # if we've given a specific word list, and this isn't in it:
                if specific_words and sub[1:-1] not in specific_words:
                    res += sub # just append as-is and continue
                    continue
                try:
                    key = sub[1:-1]
                    # if VERBOSE: 
                    print "replacing '%s' with '%s'" % (key, replacements[key])
                    res += replacements[key].__str__()
                    count += 1
                except KeyError:
                    #if it's not in our lookup table, append as-is
                    print "Key '%s' not found in replacements!" % sub
                    res += sub
            else:
                res += sub
        return res, count

    def is_key(self, string):
        return len(string) > 2 and string[0] == string[-1] == '@'

    def make_replacements(self, replacements, specific_words = None):
        ''' Finds and makes all of the replacements. '''
        document = self.get_document()
        count = 0
        for elem in document.iter():
            if elem.text:
                elem.text, c = self.replace_tags(elem.text, replacements, specific_words)
                count += c
        print "Made %d replacements" % count

    #####################                    
    # end of class DocX #
    #####################

def find_subelem(elem, name):
    ''' Given an etree graphic element, finds first subelement with given name '''
    for subelem in elem:
        if subelem.tag.split("}")[-1] == name:
            return subelem
    return None

def find_subelem_index(elem, name):
    ''' Given an etree graphic element, finds index of first subelement with given name '''
    i = 0
    for subelem in elem:
        if subelem.tag.split("}")[-1] == name:
            return i
        i += 1
    return None

def find_subelem_list(elem, namelist):
    for name in namelist:
        elem = find_subelem(elem, name)
        if elem is None:
            # print "find_subelem_list failed at element %s" % name
            return None
    return elem

path_to_description = ["graphicData", "wsp", "txbx", "txbxContent", "p", "r", "t"]
path_to_id = ["graphicData", "pic", "blipFill", "blip"]
path_to_picname = ["graphicData", "pic", "nvPicPr", "cNvPr"]

def get_description(graphicselem):
    e = find_subelem_list(graphicselem, path_to_description)
    if e is not None:
        return e.text
    else:
        return None

def get_id(graphicselem):
    elem = find_subelem_list(graphicselem, path_to_id)
    if elem is not None:
        try: 
            return elem.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
        except KeyError:
            print "Id tag found but no embed attribute"
            return None
    else:
        return None

def get_pic_name(graphicselem):
    e = find_subelem_list(graphicselem, path_to_picname)
    if e is not None:
        try:
            return e.attrib['name']
        except KeyError:
            print "Pic tag found but no name attribute"
            return None
    else:
        return None

def get_num_columns(table_elem):
    grid = find_subelem_list(table_elem, ["tblGrid"])
    if grid is not None:
        return len(grid)
    else:
        raise Exception("tblGrid element could not be found in table")

def get_num_rows(table_elem):
    count = 0
    for subelem in table_elem:
        if subelem.tag.split("}")[-1] == "tr":
            count += 1
    return count

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)
    return document


def newdocument():
    document = makeelement('document')
    document.append(makeelement('body'))
    return document


def makeelement(tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
    '''Create an element & return it'''
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
        # FIXME: rest of code below expects a single prefix
        nsprefix = nsprefix[0]
    if nsprefix:
        namespace = '{'+nsprefixes[nsprefix]+'}'
    else:
        # For when namespace = None
        namespace = ''
    newelement = etree.Element(namespace+tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
            # Quick hack: it seems every element that has a 'w' nsprefix for its tag uses the same prefix for its attributes
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
            attributenamespace = '{'+nsprefixes[attrnsprefix]+'}'

        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext
    return newelement


def pagebreak(type='page', orient='portrait'):
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if type not in validtypes:
        tmpl = 'Page break style "%s" not implemented. Valid styles: %s.'
        raise ValueError(tmpl % (type, validtypes))
    pagebreak = makeelement('p')
    if type == 'page':
        run = makeelement('r')
        br = makeelement('br', attributes={'type': type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = makeelement('pPr')
        sectPr = makeelement('sectPr')
        if orient == 'portrait':
            pgSz = makeelement('pgSz', attributes={'w': '12240', 'h': '15840'})
        elif orient == 'landscape':
            pgSz = makeelement('pgSz', attributes={'h': '12240', 'w': '15840',
                                                   'orient': 'landscape'})
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak


def paragraph(paratext, style='BodyText', breakbefore=False, jc='left', 
              font_face=None, font_size=None):
    '''Make a new paragraph element, containing a run, and some text.
    Return the paragraph element.

    @param string jc: Paragraph alignment, possible values:
                      left, center, right, both (justified), ...
                      see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                      for a full list

    If paratext is a list, spawn multiple run/text elements.
    Support text styles (paratext must then be a list of lists in the form
    <text> / <style>. Stile is a string containing a combination od 'bui' chars

    example
    paratext =\
        [ ('some bold text', 'b')
        , ('some normal text', '')
        , ('some italic underlined text', 'iu')
        ]

    '''
    # Make our elements
    paragraph = makeelement('p')

    if isinstance(paratext, list):
        text = []
        for pt in paratext:
            if isinstance(pt, (list, tuple)):
                text.append([makeelement('t', tagtext=pt[0]), pt[1]])
            else:
                text.append([makeelement('t', tagtext=pt), ''])
    else:
        text = [[makeelement('t', tagtext=paratext), ''], ]
    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={'val': style})
    pJc = makeelement('jc', attributes={'val': jc})
    pPr.append(pStyle)
    pPr.append(pJc)
    # if we've specified a font size/face, add them here
    if font_size is not None or font_face is not None:
        rPr = makeelement('rPr')
        if font_size is not None:
            sz = makeelement('sz', attributes={'val': font_size})
            szCs = makeelement('szCs', attributes={'val': font_size})
            rPr.append(sz)
            rPr.append(szCs)
        if font_size is not None:
            font = makeelement('rFonts', attributes={'ascii':font_face, 
                                                     'hAnsi': font_face})
            rPr.append(font)
        pPr.append(rPr)

    # Add the text the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = makeelement('r')
        rPr = makeelement('rPr')
        if font_size is not None or font_face is not None:
            if font_size is not None:
                sz = makeelement('sz', attributes={'val': font_size})
                szCs = makeelement('szCs', attributes={'val': font_size})
                rPr.append(sz)
                rPr.append(szCs)
            if font_size is not None:
                font = makeelement('rFonts', attributes={'ascii':font_face, 
                                                         'hAnsi': font_face})
                rPr.append(font)
        # Apply styles
        if t[1].find('b') > -1:
            b = makeelement('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = makeelement('u', attributes={'val': 'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = makeelement('i')
            rPr.append(i)
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
            lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph


def contenttypes():
    types = etree.fromstring(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/conten'
        't-types"></Types>')
    parts = {
        '/word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocu'
                                  'ment.theme+xml',
        '/word/fontTable.xml':    'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.fontTable+xml',
        '/docProps/core.xml':     'application/vnd.openxmlformats-package.co'
                                  're-properties+xml',
        '/docProps/app.xml':      'application/vnd.openxmlformats-officedocu'
                                  'ment.extended-properties+xml',
        '/word/document.xml':     'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.document.main+xml',
        '/word/settings.xml':     'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.settings+xml',
        '/word/numbering.xml':    'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.numbering+xml',
        '/word/styles.xml':       'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.styles+xml',
        '/word/webSettings.xml':  'application/vnd.openxmlformats-officedocu'
                                  'ment.wordprocessingml.webSettings+xml'}
    for part in parts:
        types.append(makeelement('Override', nsprefix=None,
                                 attributes={'PartName': part,
                                             'ContentType': parts[part]}))
    # Add support for filetypes
    filetypes = {'gif':  'image/gif',
                 'jpeg': 'image/jpeg',
                 'jpg':  'image/jpeg',
                 'png':  'image/png',
                 'rels': 'application/vnd.openxmlformats-package.relationships+xml',
                 'xml':  'application/xml'}
    for extension in filetypes:
        types.append(makeelement('Default', nsprefix=None,
                                 attributes={'Extension': extension,
                                             'ContentType': filetypes[extension]}))
    return types


def heading(headingtext, headinglevel, lang='en'):
    '''Make a new heading, return the heading element'''
    lmap = {'en': 'Heading', 'it': 'Titolo'}
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={'val': lmap[lang]+str(headinglevel)})
    run = makeelement('r')
    text = makeelement('t', tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)
    paragraph.append(run)
    # Return the combined paragraph
    return paragraph


def table(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
    """
    Return a table element based on specified parameters

    @param list contents: A list of lists describing contents. Every item in
                          the list can be a string or a valid XML element
                          itself. It can also be a list. In that case all the
                          listed elements will be merged into the cell.
    @param bool heading:  Tells whether first line should be treated as
                          heading or not
    @param list colw:     list of integer column widths specified in wunitS.
    @param str  cwunit:   Unit used for column width:
                            'pct'  : fiftieths of a percent
                            'dxa'  : twentieths of a point
                            'nil'  : no width
                            'auto' : automagically determined
    @param int  tblw:     Table width
    @param int  twunit:   Unit used for table width. Same possible values as
                          cwunit.
    @param dict borders:  Dictionary defining table border. Supported keys
                          are: 'top', 'left', 'bottom', 'right',
                          'insideH', 'insideV', 'all'.
                          When specified, the 'all' key has precedence over
                          others. Each key must define a dict of border
                          attributes:
                            color : The color of the border, in hex or
                                    'auto'
                            space : The space, measured in points
                            sz    : The size of the border, in eighths of
                                    a point
                            val   : The style of the border, see
                http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
    @param list celstyle: Specify the style for each colum, list of dicts.
                          supported keys:
                          'align' : specify the alignment, see paragraph
                                    documentation.
    @return lxml.etree:   Generated XML etree element
    """
    table = makeelement('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle', attributes={'val': ''})
    tableprops.append(tablestyle)
    tablewidth = makeelement('tblW', attributes={'w': str(tblw), 'type': str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = makeelement('tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
                attrs = {}
                for a in borders[k].keys():
                    attrs[a] = unicode(borders[k][a])
                borderelem = makeelement(b, attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = makeelement('tblLook', attributes={'val': '0400'})
    tableprops.append(tablelook)
    table.append(tableprops)
    # Table Grid
    tablegrid = makeelement('tblGrid')
    for i in range(columns):
        tablegrid.append(makeelement('gridCol', attributes={'w': str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)
    # Heading Row
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle', attributes={'val': '000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    if heading:
        i = 0
        for heading in contents[0]:
            cell = makeelement('tc')
            # Cell properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellstyle = makeelement('shd', attributes={'val': 'clear',
                                                       'color': 'auto',
                                                       'fill': 'FFFFFF',
                                                       'themeFill': 'text2',
                                                       'themeFillTint': '99'})
            cellprops.append(cellwidth)
            cellprops.append(cellstyle)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(heading, (list, tuple)):
                heading = [heading]
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
                    cell.append(paragraph(h, jc='center'))
            row.append(cell)
            i += 1
        table.append(row)
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = makeelement('tr')
        i = 0
        for content in contentrow:
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(content, (list, tuple)):
                content = [content]
            for c in content:
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    if celstyle and 'align' in celstyle[i].keys():
                        align = celstyle[i]['align']
                    else:
                        align = 'left'
                    cell.append(paragraph(c, jc=align))
            row.append(cell)
            i += 1
        table.append(row)
    return table

def make_row(contentrow, colw = None, cwunit = "dxa", celstyle = None,
             font_face = None, font_size = None, borders=[]):
    row = makeelement('tr')
    i = 0
    for content in contentrow:
        cell = makeelement('tc')
        # Properties
        cellprops = makeelement('tcPr')
        if colw:
            wattr = {'w': str(colw[i]), 'type': cwunit}
        else:
            wattr = {'w': '0', 'type': 'auto'}
        do_bottom = True if "bottom" in borders else False
        do_top = True if "top" in borders else False
        if do_bottom:
            bottom = makeelement('bottom', attributes={'val':'single', 'sz':'4',
                                                        'space':'0', 'color':'auto'})
            cellprops.append(bottom)
        if do_top:
            top = makeelement('top', attributes={'val':'single', 'sz':'4',
                                                        'space':'0', 'color':'auto'})
            cellprops.append(top)

        cellwidth = makeelement('tcW', attributes=wattr)
        cellprops.append(cellwidth)
        cell.append(cellprops)
        # Paragraph (Content)
        if not isinstance(content, (list, tuple)):
            content = [content]
        for c in content:
            if isinstance(c, etree._Element):
                cell.append(c)
            else:
                if celstyle and 'align' in celstyle[i].keys():
                    align = celstyle[i]['align']
                else:
                    align = 'left'
                cell.append(paragraph(c, jc=align, font_face=font_face, 
                                                    font_size=font_size))
        row.append(cell)
        i += 1
    return row

def picture(relationshiplist, picname, picdescription, pixelwidth=None, pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
    '''Take a relationshiplist, picture file name, and return a paragraph containing the image
    and an updated relationshiplist'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''
    # Copy the file into the media dir
    media_dir = join(template_dir, 'word', 'media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    shutil.copyfile(picname, join(media_dir, picname))

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth, pixelheight = Image.open(picname).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs
    emuperpixel = 12700
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)

    # Set relationship ID to the first available
    picid = '2'
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])

    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill', nsprefix='pic')
    blipfill.append(makeelement('blip', nsprefix='a', attrnsprefix='r',
                    attributes={'embed': picrelid}))
    stretch = makeelement('stretch', nsprefix='a')
    stretch.append(makeelement('fillRect', nsprefix='a'))
    blipfill.append(makeelement('srcRect', nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = makeelement('nvPicPr', nsprefix='pic')
    cnvpr = makeelement('cNvPr', nsprefix='pic',
                        attributes={'id': '0', 'name': 'Picture 1', 'descr': picname})
    nvpicpr.append(cnvpr)
    cnvpicpr = makeelement('cNvPicPr', nsprefix='pic')
    cnvpicpr.append(makeelement('picLocks', nsprefix='a',
                    attributes={'noChangeAspect': str(int(nochangeaspect)),
                                'noChangeArrowheads': str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)

    # 3. The Shape properties
    sppr = makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
    pic = makeelement('pic', nsprefix='pic')
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = makeelement('graphicData', nsprefix='a',
                              attributes={'uri': 'http://schemas.openxmlforma'
                                                 'ts.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic', nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks', nsprefix='a',
                             attributes={'noChangeAspect': '1'})
    framepr = makeelement('cNvGraphicFramePr', nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr', nsprefix='wp',
                        attributes={'id': picid, 'name': 'Picture 1',
                                    'descr': picdescription})
    effectextent = makeelement('effectExtent', nsprefix='wp',
                               attributes={'l': '25400', 't': '0', 'r': '0',
                                           'b': '0'})
    extent = makeelement('extent', nsprefix='wp',
                         attributes={'cx': width, 'cy': height})
    inline = makeelement('inline', attributes={'distT': "0", 'distB': "0",
                                               'distL': "0", 'distR': "0"},
                         nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return relationshiplist, paragraph


def search(document, search):
    '''Search a document for a regex, return success / fail result'''
    result = False
    searchre = re.compile(search)
    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    result = True
    return result


def replace(document, search, replace):
    ''' Replace all occurences of string with a different string, return updated document'''
    newdocument = document
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search, replace, element.text)
    return newdocument


def clean(document):
    """ Perform misc cleaning operations on documents.
        Returns cleaned document.
    """

    newdocument = document

    # Clean empty text and r tags
    for t in ('t', 'r'):
        rmlist = []
        for element in newdocument.iter():
            if element.tag == '{%s}%s' % (nsprefixes['w'], t):
                if not element.text and not len(element):
                    rmlist.append(element)
        for element in rmlist:
            element.getparent().remove(element)

    return newdocument


def findTypeParent(element, tag):
    """ Finds fist parent of element of the given type

    @param object element: etree element
    @param string the tag parent to search for

    @return object element: the found parent or None when not found
    """

    p = element
    while True:
        p = p.getparent()
        if p.tag == tag:
            return p

    # Not found
    return None


def AdvSearch(document, search, bs=3):
    '''Return set of all regex matches

    This is an advanced version of python-docx.search() that takes into
    account blocks of <bs> elements at a time.

    What it does:
    It searches the entire document body for text blocks.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.

    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search : 'Hello,'
    output blocks : [ 'Hello,' ]

    original text blocks : [ 'Hel', 'lo', ' __', 'name', '__!' ]
    search : '(__[a-z]+__)'
    output blocks : [ '__name__' ]

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
                          append, or a list of etree elements
    @param int       bs: See above

    @return set      All occurences of search string

    '''

    # Compile the search regexp
    searchre = re.compile(search)

    matches = []

    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []

    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)

                # Search all combinations, of searchels, starting from
                # smaller up to bigger ones
                # l = search lenght
                # s = search start
                # e = element IDs to merge
                found = False
                for l in range(1, len(searchels)+1):
                    if found:
                        break
                    for s in range(len(searchels)):
                        if found:
                            break
                        if s+l <= len(searchels):
                            e = range(s, s+l)
                            txtsearch = ''
                            for k in e:
                                txtsearch += searchels[k].text

                            # Searcs for the text in the whole txtsearch
                            match = searchre.search(txtsearch)
                            if match:
                                matches.append(match.group())
                                found = True
    return set(matches)


def advReplace(document, search, replace, bs=3):
    """
    Replace all occurences of string with a different string, return updated
    document

    This is a modified version of python-docx.replace() that takes into
    account blocks of <bs> elements at a time. The replace element can also
    be a string or an xml etree element.

    What it does:
    It searches the entire document body for text blocks.
    Then scan thos text blocks for replace.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.

    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello,' / 'Hi!'
    output blocks : [ 'Hi!', '', ' world!' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello, world' / 'Hi!'
    output blocks : [ 'Hi!!', '', '' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hel' / 'Hal'
    output blocks : [ 'Hal', 'lo,', ' world!' ]

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
    @param mixed     replace: The replacement text or lxml.etree element to
                         append, or a list of etree elements
    @param int       bs: See above

    @return instance The document with replacement applied

    """
    # Enables debug output
    DEBUG = False

    newdocument = document

    # Compile the search regexp
    searchre = re.compile(search)

    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []

    for element in newdocument.iter():
        if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)

                # Search all combinations, of searchels, starting from
                # smaller up to bigger ones
                # l = search lenght
                # s = search start
                # e = element IDs to merge
                found = False
                for l in range(1, len(searchels)+1):
                    if found:
                        break
                    #print "slen:", l
                    for s in range(len(searchels)):
                        if found:
                            break
                        if s+l <= len(searchels):
                            e = range(s, s+l)
                            #print "elems:", e
                            txtsearch = ''
                            for k in e:
                                txtsearch += searchels[k].text

                            # Searcs for the text in the whole txtsearch
                            match = searchre.search(txtsearch)
                            if match:
                                found = True

                                # I've found something :)
                                if DEBUG:
                                    log.debug("Found element!")
                                    log.debug("Search regexp: %s", searchre.pattern)
                                    log.debug("Requested replacement: %s", replace)
                                    log.debug("Matched text: %s", txtsearch)
                                    log.debug("Matched text (splitted): %s", map(lambda i: i.text, searchels))
                                    log.debug("Matched at position: %s", match.start())
                                    log.debug("matched in elements: %s", e)
                                    if isinstance(replace, etree._Element):
                                        log.debug("Will replace with XML CODE")
                                    elif isinstance(replace(list, tuple)):
                                        log.debug("Will replace with LIST OF ELEMENTS")
                                    else:
                                        log.debug("Will replace with:", re.sub(search, replace, txtsearch))

                                curlen = 0
                                replaced = False
                                for i in e:
                                    curlen += len(searchels[i].text)
                                    if curlen > match.start() and not replaced:
                                        # The match occurred in THIS element. Puth in the
                                        # whole replaced text
                                        if isinstance(replace, etree._Element):
                                            # Convert to a list and process it later
                                            replace = [replace]
                                        if isinstance(replace, (list, tuple)):
                                            # I'm replacing with a list of etree elements
                                            # clear the text in the tag and append the element after the
                                            # parent paragraph
                                            # (because t elements cannot have childs)
                                            p = findTypeParent(searchels[i], '{%s}p' % nsprefixes['w'])
                                            searchels[i].text = re.sub(search, '', txtsearch)
                                            insindex = p.getparent().index(p) + 1
                                            for r in replace:
                                                p.getparent().insert(insindex, r)
                                                insindex += 1
                                        else:
                                            # Replacing with pure text
                                            searchels[i].text = re.sub(search, replace, txtsearch)
                                        replaced = True
                                        log.debug("Replacing in element #: %s", i)
                                    else:
                                        # Clears the other text elements
                                        searchels[i].text = ''
    return newdocument


def getdocumenttext(document):
    '''Return the raw text of a document, as a list of paragraphs.'''
    paratextlist = []
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in document.iter():
        # Find p (paragraph) elements
        if element.tag == '{'+nsprefixes['w']+'}p':
            paralist.append(element)
    # Since a single sentence might be spread over multiple text elements, iterate through each
    # paragraph, appending all text (t) children to that paragraphs text.
    for para in paralist:
        paratext = u''
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == '{'+nsprefixes['w']+'}t':
                if element.text:
                    paratext = paratext+element.text
            elif element.tag == '{'+nsprefixes['w']+'}tab':
                paratext = paratext + '\t'
        # Add our completed paragraph text to the list of paragraph text
        if not len(paratext) == 0:
            paratextlist.append(paratext)
    return paratextlist


def coreproperties(**kwargs):
    '''Create core properties (common document properties referred to in the 'Dublin Core' specification).
    See appproperties() for other stuff.'''
    coreprops = makeelement('coreProperties', nsprefix='cp')
    for s in ['title', 'subject', 'creator']:
        if s in kwargs:
            coreprops.append(makeelement(s, tagtext=kwargs[s], nsprefix='dc'))
    if 'keywords' in kwargs:
        coreprops.append(makeelement('keywords', tagtext=','.join(kwargs['keywords']), nsprefix='cp'))
    if 'lastmodifiedby' in kwargs:
        coreprops.append(makeelement('lastModifiedBy', tagtext=kwargs['lastmodifiedby'], nsprefix='cp'))
    elif 'creator' in kwargs:
        coreprops.append(makeelement('lastModifiedBy', tagtext=kwargs['creator'], nsprefix='cp'))
    coreprops.append(makeelement('revision', tagtext='1', nsprefix='cp'))
    coreprops.append(makeelement('category', tagtext='Examples', nsprefix='cp'))
    coreprops.append(makeelement('description', tagtext='Examples', nsprefix='dc'))
    currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
    # Document creation and modify times
    # Prob here: we have an attribute who name uses one namespace, and that
    # attribute's value uses another namespace.
    # We're creating the element from a string as a workaround...
    for doctime in ['created', 'modified']:
        coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
        pass
    return coreprops


def appproperties():
    """
    Create app-specific properties. See docproperties() for more common
    document properties.

    """
    appprops = makeelement('Properties', nsprefix='ep')
    appprops = etree.fromstring(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties x'
        'mlns="http://schemas.openxmlformats.org/officeDocument/2006/extended'
        '-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocum'
        'ent/2006/docPropsVTypes"></Properties>')
    props =\
        {'Template':             'Normal.dotm',
         'TotalTime':            '6',
         'Pages':                '1',
         'Words':                '83',
         'Characters':           '475',
         'Application':          'Microsoft Word 12.0.0',
         'DocSecurity':          '0',
         'Lines':                '12',
         'Paragraphs':           '8',
         'ScaleCrop':            'false',
         'LinksUpToDate':        'false',
         'CharactersWithSpaces': '583',
         'SharedDoc':            'false',
         'HyperlinksChanged':    'false',
         'AppVersion':           '12.0000'}
    for prop in props:
        appprops.append(makeelement(prop, tagtext=props[prop], nsprefix=None))
    return appprops


def websettings():
    '''Generate websettings'''
    web = makeelement('webSettings')
    web.append(makeelement('allowPNG'))
    web.append(makeelement('doNotSaveAsSingleFile'))
    return web


def relationshiplist():
    relationshiplist =\
        [['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/numbering', 'numbering.xml'],
         ['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/styles', 'styles.xml'],
         ['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/settings', 'settings.xml'],
         ['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/webSettings', 'webSettings.xml'],
         ['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/fontTable', 'fontTable.xml'],
         ['http://schemas.openxmlformats.org/officeDocument/2006/'
          'relationships/theme', 'theme/theme1.xml']]
    return relationshiplist


def wordrelationships(relationshiplist):
    '''Generate a Word relationships file'''
    # Default list of relationships
    # FIXME: using string hack instead of making element
    #relationships = makeelement('Relationships', nsprefix='pr')
    relationships = etree.fromstring(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006'
        '/relationships"></Relationships>')
    count = 0
    for relationship in relationshiplist:
        # Relationship IDs (rId) start at 1.
        rel_elm = makeelement('Relationship', nsprefix=None,
                              attributes={'Id':     'rId'+str(count+1),
                                          'Type':   relationship[0],
                                          'Target': relationship[1]}
                              )
        relationships.append(rel_elm)
        count += 1
    return relationships


def savedocx(document, coreprops, appprops, contenttypes, websettings, wordrelationships, output):
    '''Save a modified document'''
    assert os.path.isdir(template_dir)
    docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

    # Move to the template data path
    prev_dir = os.path.abspath('.')  # save previous working dir
    os.chdir(template_dir)

    # Serialize our trees into out zip file
    treesandfiles = {document:     'word/document.xml',
                     coreprops:    'docProps/core.xml',
                     appprops:     'docProps/app.xml',
                     contenttypes: '[Content_Types].xml',
                     websettings:  'word/webSettings.xml',
                     wordrelationships: 'word/_rels/document.xml.rels'}
    for tree in treesandfiles:
        log.info('Saving: %s' % treesandfiles[tree])
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr(treesandfiles[tree], treestring)

    # Add & compress support files
    files_to_ignore = ['.DS_Store']  # nuisance from some os's
    for dirpath, dirnames, filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath, filename)
            archivename = templatefile[2:]
            log.info('Saving: %s', archivename)
            docxfile.write(templatefile, archivename)
    log.info('Saved new file to: %r', output)
    docxfile.close()
    os.chdir(prev_dir)  # restore previous working dir
