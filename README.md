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

## v2.3 structured LLM search instructions

The natural-language prompt is now interpreted once as **search instructions** before candidate files are LLM-ranked. The model separates the request into:

- filename hints
- folder/path hints
- file-type hints
- content hints
- exclusions
- hard/must conditions
- soft preferences

Each candidate file is then evaluated against separate evidence sections rather than treating everything as document text.

### Metadata sent separately to the LLM

- file name
- file name without extension
- extension
- friendly file type such as `excel spreadsheet workbook`
- parent folder
- full folder path
- full file path
- size
- modified date

The extracted file content is supplied as its own independent section.

For example:

`find an Excel quotation in the Finance folder mentioning transformers, but exclude PDFs`

is interpreted as instructions across metadata and content. `Excel` applies to file type, `Finance` applies to the folder/path, `quotation` may be supported by filename or content, `transformers` applies to content/context, and the PDF exclusion is treated as a file-type rule rather than requiring those literal words to appear in a document.

When the GUI file-type selector is **All supported files**, LLM file-type hints can also narrow the candidate set automatically.

The preserved v2.2.0 package remains in the repo and `v2_3_search_logic.py` is applied by the launcher, making the new search behaviour isolated and easy to review or roll back.

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

The **File type (index + search)** selector applies to both operations:

- **Build / Update Index** only enumerates and opens files in the selected category.
- **Search Indexed Files** only searches that selected category.
- Updating one category does not remove previously indexed files from other categories.
- Choose **All supported files** when you deliberately want a complete index refresh.
