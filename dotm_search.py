#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = 'Eric Hanson'

import os
import zipfile
import argparse


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', help='search directory')
    parser.add_argument('text', help='string to search for')
    return parser


def main():
    parser = create_parser()
    entry = parser.parse_args()
    text = entry.text
    path = entry.dir
    print("Searching directory {} for files with text '{}' ").format(path, text)
    file_list = os.listdir(path)
    match, scanned = 0, 0
    for file in file_list:
        scanned += 1
        full_path = os.path.join(path, file)
        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as zip:
                name = zip.namelist()
                if 'word/document.xml' in name:
                    with zip.open('word/document.xml', 'r') as doc:
                        for line in doc:
                            line = line.decode('utf-8')
                            location = line.find(text)
                            if location >= 0:
                                print('Match found in file ' + path)
                                print('...%s...' % line[location - 40:location+41])
                                match += 1
    print('Total dotm files searched: %s' % scanned)
    print('Total dotm files matched: %s' % match)

if __name__ == '__main__':
    main()
