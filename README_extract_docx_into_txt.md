# extract_docx_into_txt.py

Extracts the full text content from a `.docx` file — including all embedded files (Excel, PDF, Word, email, CSV) — and outputs everything as a single plain text file.

Embedded files are extracted **in place**, meaning their text content appears at the exact position where the embedded object exists in the original document.

## Installation

pip install -r requirements_extract_docx_into_txt.txt

## Usage

python3 extract_docx_into_txt.py <input.docx> > output.txt

Output is printed to stdout. Redirect to a file with `>`.

## Supported Embedded File Types

- `.docx` (Word documents)
- `.xlsx` / `.xlsm` (Excel spreadsheets)
- `.pdf` (PDF documents)
- `.eml` (email files)
- `.msg` (Outlook messages)
- `.csv` / `.txt` (plain text files)

## Output Format

The output is a plain text file where:

- Normal document text appears as-is (paragraphs, table rows)
- Tables are rendered as pipe-separated rows: `Cell 1 | Cell 2 | Cell 3`
- Each embedded file is shown between clear markers