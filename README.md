# Receipt to Spreadsheet

Extracts receipt data from images and PDFs using Claude (Anthropic) and outputs to Excel.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Create `.env`:
```
ANTHROPIC_API_KEY=your_key
RECEIPT_MODEL=claude-sonnet-4-5-20250514
```

## Usage

```bash
python receipt_to_sheet.py        # Process all files
python receipt_to_sheet.py -n 10  # Process 10 newest files
```

Files are read from the folder(s) configured in `config.py` (newest first) and appended to `receipts.xlsx`.

## How It Works

- **Images** (`.jpg`, `.png`, `.webp`, `.gif`) are sent to Claude Vision as base64
- **PDFs** are sent natively to Claude — no conversion or text extraction needed

Claude returns structured JSON via constrained decoding, which is written directly into the spreadsheet.

## Output Fields (default config)

- Date
- Description
- House Number
- Amount
