#!/usr/bin/env python3

# Convert a Word bibliography into a set of BibTeX entries.
# Only vaguely tested, so use at your own risk.
# Author: Alex Jordan <alex.jordan@raytheon.com>

import argparse
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

wordns = { 'b': 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography' }

def handle_site(source):
    return None

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
                print(entry)
            else:
                sys.stderr.write('Could not process entry with title "{}"\n'.format(source.find('b:Title', wordns).text))

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