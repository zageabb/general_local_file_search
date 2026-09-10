# Generic Local File Search GUI

A generic Tkinter search application for large local folders and locally synchronised OneDrive folders.

## Main change from the supplier-only version

The search logic is no longer hard-coded for supplier dashboards.

Enter any natural-language request in **What are you looking for?**, for example:

- `approved suppliers by MDF with supplier ratings and approved status`
- `old spreadsheet containing the Europe PGGI Hub dashboard`
- `documents discussing transformer delivery delays`
- `PowerPoint about SCM digitalisation`
- `files mentioning SAP VIM invoice exceptions`
- `quotation for 132kV disconnectors`
- `meeting notes about Project Procure data quality`

The original supplier-dashboard task is therefore just one saved/example query rather than the purpose of the program.

## Why it indexes first

If your OneDrive contains tens of thousands of files, re-opening every workbook for every search is far too slow.

The application builds a local SQLite index at:

`~/.generic_file_search/file_index.sqlite3`

After the first run, **Build / Update Index** only re-reads files whose size or modification time changed.

## Supported content

Direct readers:

- XLSX / XLSM
- XLS
- XLSB
- CSV / TSV
- DOCX
- PPTX
- PDF
- TXT / Markdown
- JSON / XML / YAML
- source-code and common configuration files

Optional LibreOffice fallback:

- old `.doc`
- old `.ppt`
- `.ods`
- troublesome `.xls`

The source file is never modified.

## Local Ollama

Defaults:

- `http://192.168.1.249:11434`
- `llama3.2:latest`

The local deterministic index produces the first ranking. Ollama is then used only on the best candidates, so a search over 30,000 files does **not** send 30,000 files through the model.

Only extracted text is sent to Ollama.

## Install

```bash
python3 -m pip install -r generic_file_search_requirements.txt
```

Tkinter ships with many Python installations. On macOS python.org builds generally include it.

LibreOffice is optional but recommended if old Office files need to be searched.

## Run

```bash
python3 generic_file_search_gui.py
```

## Recommended first use

1. Select your locally synchronised OneDrive folder.
2. Click **Build / Update Index**.
3. Let the initial index complete.
4. Type any search request.
5. Click **Search Indexed Files**.
6. Double-click a result to open it.
7. Right-click a result to reveal it in Finder/Explorer or copy its full path.

## Python 3.13 / XLSM fix

This version explicitly keeps `keep_vba=False` when using openpyxl for read-only searching.

That avoids the repeated:

`ZipFile.__del__: ValueError: I/O operation on closed file`

warning that appeared when scanning large numbers of macro-enabled workbooks. VBA does not need to be copied because the application never saves the workbook.

The index stores up to 120,000 extracted characters per file to keep very large collections manageable.

## File type filtering

The **File type** dropdown filters searches without rebuilding the index.

Available presets:

- All supported files
- Spreadsheets
- Excel workbooks
- CSV / TSV
- Documents
- Word documents
- Presentations
- PowerPoint
- PDF
- Text / Markdown
- Code / config

The full index is retained, so changing the file type filter does not require re-indexing.

## v2.2 file-type indexing behaviour

The **File type (index + search)** selector now applies to both operations:

- **Build / Update Index** only enumerates and opens files in the selected category.
- **Search Indexed Files** only searches that selected category.
- Updating one category does not remove previously indexed files from other categories.
- Choose **All supported files** when you deliberately want a complete index refresh.
