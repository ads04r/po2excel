PO <---> Excel
==============

A tool to assist Arches translators.

One thing cannot be denied: love them, or hate them, people know how to use spreadsheets.
Transifex is a nice tool for editing strings one at a time, and I like how it lets you
download the PO file in order to bulk-edit multiple strings in one session. However I
figured it'll be better for more users if there were a way of bulk-editing a spreadsheet
of translated strings, rather than having to modify a PO file which is possibly a bit
technical for many non-programmers.

Meanwhile, [@zoometh](https://github.com/zoometh) developed a way of automatically
filling PO files with automated translations from Google Translate using the Python
deep-translator library. So I borrowed some of that code in order to include an extra
step in my conversion process.

Example workflow
----------------

1. Download PO file from Transifex
2. Run PO file through po2excel

    ```python3 po2excel.py input_file.po output_file.xlsx --format xlsx```

3. Load the resulting file into Excel or equivalent
4. Browse the columns; the first column will contain the msgid, the second will
   contain any previous (manually) translation extracted from the PO file, and
   the third will contain an automatic translation of the msgid according to
   Google Translate.
5. For all the rows where the automatic translation is sufficient, copy the
   cell in the third column to the second column
6. If you like, any rows that do not have a sufficient automatic translation,
   you can enter this into the second column instead
7. Save the spreadsheet
8. Run the spreadsheet through excel2po, using the unmodified PO file downloaded
   from Transifex as a base file. Doing this will sort out any clashes.

    ```python exceltopo.py spreadsheet.xlsx output_file.po --base input_file.po```

9. Upload the resulting po file back to Transifex.

