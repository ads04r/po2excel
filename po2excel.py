#!/usr/bin/python3
# PO to Excel
# By Ash Smith, MarEA project, University of Southampton
# Based heavily on code by Thomas Huet, EAMENA project, University of Oxford

import os, re, csv, polib, sys, argparse
from deep_translator import GoogleTranslator
from deep_translator.exceptions import NotValidPayload
from progress.bar import IncrementalBar as progressbar
from openpyxl import Workbook

argp = argparse.ArgumentParser()
argp.add_argument('POFile', metavar='po_file', type=str, help='The PO file to convert')
argp.add_argument('CSVFile', metavar='csv_file', type=str, help='The CSV file to write')
argp.add_argument('-lang', '--language', action='store', type=str, help='The 2-letter code for the language to use', default='')
argp.add_argument('-f', '--format', action='store', type=str, help='The file format to export, xlsx or csv (default is csv)', default='csv')
args = argp.parse_args()

target_language = args.language
target_format = args.format
path_fold = os.getcwd()
input_file = args.POFile
output_file = args.CSVFile

po = polib.pofile(input_file)

target_format = target_format.lower()
if target_format == 'xls':
	target_format = 'xlsx'
if target_format != 'xlsx':
	target_format = 'csv'
if target_language == '':
	if 'Language' in po.metadata:
		target_language = po.metadata['Language']

if target_language == '':
	print("Cannot determine language from PO file or input arguments. Try calling the command again using the --language switch.")
	sys.exit(1)

print('Working in language: ' + target_language)
print('File contains ' + str(len(po)) + ' strings')
print('of which ' + str(po.percent_translated()) + '% have already been translated')

ret = []

ret.append(['string', 'translation', 'automatic_translation'])

with progressbar('Translating', max=len(po)) as bar:
	for item in po:

		bar.next()

		if item.msgstr != '':
			ret.append([item.msgid, item.msgstr, ''])
			continue
		try:
			item.msgstr = GoogleTranslator(source='en', target=target_language).translate(item.msgid)
		except NotValidPayload:
			item.msgstr = ''

		ret.append([item.msgid, '', item.msgstr])

if target_format == 'csv':

	with open(output_file, 'w') as csv_file:
		csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
		for item in ret:
			csv_writer.writerow(item)

if target_format == 'xlsx':

	wb = Workbook()
	ws = wb.active
	for item in ret:
		ws.append(item)

	wb.save(output_file)
