# Receipt to Spreadsheet

Extracts receipt data from images using OpenAI Vision and outputs to Excel.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Create `.env`:
```
OPENAI_API_KEY=your_key
RECEIPT_MODEL=gpt-4o-mini
```

## Usage

```bash
python receipt_to_sheet.py        # Process all images
python receipt_to_sheet.py -n 10  # Process 10 newest images
```

Images are read from `receipt_images/` (newest first) and appended to `receipts.xlsx`.

## Output Fields

- Date
- Payee
- Description (Meal/Treat)
- Entertainment (blank)
- GST
- Total

