#!/usr/bin/env python3

# Convert a Word bibliography into a set of BibTeX entries.
# Only vaguely tested, so use at your own risk.
# Author: Alex Jordan <alex.jordan@raytheon.com>

import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

wordns = { 'b': 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography' }

def tex_escape(text):
    """
        :param text: a plain text message
        :return: the message escaped to appear correctly in LaTeX
        Copied from https://stackoverflow.com/a/25875504/3699
    """
    conv = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '^': r'\^{}',
        '\\': r'\textbackslash{}',
        '<': r'\textless ',
        '>': r'\textgreater ',
    }
    regex = re.compile('|'.join(re.escape(key) for key in sorted(conv.keys(), key = lambda item: - len(item))))
    return regex.sub(lambda match: conv[match.group()], text)

def handle_author(author):
    # yeah - for some reason, b:Author has a sub-node called b:Author
    author_list = []
    real_author = author.find('b:Author', wordns)
    for entry in real_author:
        if entry.tag.endswith('Corporate'):
            author_list.append(entry.text)
        elif entry.tag.endswith('NameList'):
            for person in entry.findall('b:Person', wordns):
                firstname_node = person.find('b:First', wordns)
                lastname_node = person.find('b:Last', wordns)
                firstname = firstname_node.text if firstname_node is not None else ''
                lastname = lastname_node.text if lastname_node is not None else ''
                author_list.append(firstname + " " + lastname)
    return " and ".join(author_list)

def handle_site(source):
    template = \
'''
@misc{%s,
    author = {%s},
    title = {%s},
    howpublished = {Available at \\url{%s} (%s)}   
}'''
    tag = source.find('b:Tag', wordns).text
    author = handle_author(source.find('b:Author', wordns))
    title = source.find('b:Title', wordns).text
    url = tex_escape(source.find('b:URL', wordns).text)
    year_accessed = source.find('b:YearAccessed', wordns).text
    month_accessed = source.find('b:MonthAccessed', wordns).text
    day_accessed = source.find('b:DayAccessed', wordns).text
    date_accessed = year_accessed + "/" + month_accessed + "/" + day_accessed

    return template % (tag, author, title, url, date_accessed)

def handle_journalarticle(source):
    return None

def handle_report(source):
    return None

def handle_conferenceproceedings(source):
    return None

def process_file(infile, outfile):
    with zipfile.ZipFile(infile, 'r') as wordfile:
        refs_file = wordfile.open('customXml/item1.xml')
        xmldata = refs_file.read()
        tree = ET.fromstring(xmldata)
        sources = tree.findall('b:Source', wordns)
        for source in sources:
            entry = None
            source_type = source.find('b:SourceType', wordns).text
            if source_type == 'JournalArticle':
                entry = handle_journalarticle(source)
            elif source_type == 'InternetSite':
                entry = handle_site(source)
            elif source_type == 'Report':
                entry = handle_report(source)
            elif source_type == 'ConferenceProceedings':
                entry = handle_conferenceproceedings(source)
            else:
                print("Don't understand how to process source type {}".format(source_type))

            if entry is not None:
                outfile.write(entry)
            else:
                sys.stderr.write('Could not process entry with title "{}" ({})\n'.format(source.find('b:Title', wordns).text, source_type))

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Convert a bibliography embedded in a MS Word document into a BibTeX file.')
    parser.add_argument('infile', metavar='INFILE', help='Word file to extract from')
    parser.add_argument('-o', dest='outfile', metavar='OUTFILE', default='<STDOUT>', help='Output file. Default is stdout.')
    args = parser.parse_args()

    sys.stderr.write('Extracting "{}" to {}'.format(args.infile, args.outfile))

    outfile = sys.stdout
    if args.outfile != '<STDOUT>':
        outfile = open(args.outfile, 'w')

    process_file(args.infile, outfile)