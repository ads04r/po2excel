#!/usr/bin/python3
# Excel to PO
# By Ash Smith, MarEA project, University of Southampton

import os, re, csv, polib, sys, argparse
from openpyxl.cell import Cell, MergedCell
from openpyxl import load_workbook

argp = argparse.ArgumentParser()
argp.add_argument('CSVFile', metavar='csv_file', type=str, help='The CSV or XLSX file to convert')
argp.add_argument('POFile', metavar='po_file', type=str, help='The PO file to write')
argp.add_argument('-lang', '--language', action='store', type=str, help='The 2-letter code for the language to use', default='')
argp.add_argument('-b', '--base', action='store', type=str, help='The PO file to use as a base', default='')
args = argp.parse_args()

target_language = args.language
base_file = args.base
input_file = args.CSVFile
output_file = args.POFile
mapping = {}

if os.path.exists(base_file):
	po = polib.pofile(base_file)
else:
	po = polib.POFile()
if target_language != '':
	po.metadata['Language'] = target_language
if 'Language' in po.metadata:
	target_language = po.metadata['Language']

if target_language == '':
	print("Cannot determine language from PO file or input arguments. Try calling the command again using the --language switch.")
	sys.exit(1)

print('Working in language: ' + target_language)

data = []

if input_file.lower().endswith('.xlsx'):
	wb = load_workbook(input_file)
	ws = wb.active
	for row in ws.rows:
		item = []
		for cell in row:
			if cell.value is None:
				item.append('')
			else:
				item.append(str(cell.value))
		data.append(item)

if input_file.lower().endswith('.csv'):
	with open(input_file) as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='"')
		for row in csvreader:
			data.append(row)

headers = []
data_formatted = []
for row in data:
	if len(headers) == 0:
		headers = row
		continue
	if len(row) > len(headers):
		continue
	item = {}
	itemlen = 0
	for i in range(0, len(headers)):
		k = str(headers[i])
		item[k] = row[i]
		itemlen = itemlen + len(row[i])
	if itemlen == 0:
		continue
	data_formatted.append(item)
data = data_formatted

for i in range(0, len(po)):
	msgid = str(po[i].msgid)
	mapping[msgid] = i

for item in data:

	if not('string' in item):
		continue
	k = str(item['string'])
	if k == '':
		continue
	if not(k in mapping):
		entry = polib.POEntry(msgid=k, msgstr=item['translation'])
		po.append(entry)
		continue
	eid = mapping[k]
	if not('translation' in item):
		continue
	if item['translation'] == '':
		continue
	po[eid].msgstr = item['translation']

po.save(output_file)
