#!/usr/bin/env python2
# -*- coding: utf-8 -*-
'''
convert Microsoft Word 2007 docx files to wiki markup files
'''

from lxml import etree
import zipfile
import string
import sys

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
nsprefixes = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    'dc':"http://purl.org/dc/elements/1.1/",
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships'
    }

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)
    return document

def ns_w(tag):
    return '{' + nsprefixes['w'] + '}' + tag

def ns_pic(tag):
    return '{' + nsprefixes['pic'] + '}' + tag

def parsenode(node):
    operaters = {
            ns_w('p') : lambda x : parsep(x),
            ns_w('t') : lambda x : parset(x),
            ns_w('tbl') : lambda x : parsetbl(x),
            ns_w('tr') : lambda x : parsetr(x),
            ns_w('tc') : lambda x : parsetc(x)
            }
    operater = operaters.get(node.tag)
    if not operater:
        operater = extracttext

    return operater(node)

def parsep(p):
    # heading
    # list
    # normal
    content = ''
    for element in p.iter():
        if element.tag == ns_w('pStyle'):
            heading = element.attrib.get(ns_w('val'))
            if heading and heading in [str(i) for i in [1,2,3,4,5,6]]:
                content += 'h' + heading + '. '
                break
        if element.tag == ns_w('ilvl'):
            level = element.attrib.get(ns_w('val'))
            content += '*' * (1 + string.atoi(level)) + ' '
    content += extracttext(p)
    return content + '\n'

def parset(t):
    content = ''
    content += t.text
    return content

def parsetbl(tbl):
    content = ''
    for element in tbl:
        content += parsenode(element)
    return content

def parsetr(tr):
    content = '|'
    for element in tr:
        content += parsenode(element)
    content += '\n'
    return content

def parsetc(tc):
    content = ''
    escape = False
    for element in tc.iter():
        if element.tag == ns_w('p'):
            if escape:
                content += '\n'
            content += extracttext(element)
            escape = True
    if escape:
        content = content.replace('\n', '\\\\\n')
    if content == '':
        content = ' '
    return content + '|'

def parsebody(body):
    content = ''
    for element in body:
        content += parsenode(element)
    return content

def extracttext(node):
    content = ''
    for element in node.iter():
        if element.tag == ns_w('t'):
            content += element.text
        elif element.tag == ns_pic('cNvPr'):
            content += '!' + element.attrib.get('name') + '!'
    return content

def docx2wiki(docx):
    for element in docx:
        if element.tag == ns_w('body'):
            return parsebody(element)

if __name__ == '__main__':
    docx = opendocx(sys.argv[1])
    newfile = open(sys.argv[2],'w')        
    wiki = docx2wiki(docx)
    newfile.write(wiki.encode('utf-8'))

