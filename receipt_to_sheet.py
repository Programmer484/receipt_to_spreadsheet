import os
import json
import base64
import mimetypes
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, Optional

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

load_dotenv()

HEADERS = ["Date", "Payee", "Description", "Entertainment", "GST", "Total"]
SHEET_NAME = "Receipts"

# Excel-friendly formats
DATE_FMT = "d-mmm-yyyy"
CURRENCY_FMT = '"$"#,##0.00'  # change if you prefer a different currency format


def image_path_to_data_url(image_path: str) -> str:
    """
    Encodes a local image as a Base64 data URL for OpenAI vision input.
    OpenAI supports passing images as Base64-encoded data URLs.  (docs)
    """
    mime, _ = mimetypes.guess_type(image_path)
    if not mime:
        # reasonable default for receipts if extension is missing/odd
        mime = "image/jpeg"

    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")

    return f"data:{mime};base64,{b64}"


def safe_decimal(x: Any) -> Optional[Decimal]:
    if x is None:
        return None
    try:
        return Decimal(str(x))
    except (InvalidOperation, ValueError):
        return None


def ensure_sheet_and_headers(wb_path: str) -> tuple[Any, Worksheet]:
    """
    Creates/opens the workbook and ensures the target sheet + headers exist.
    """
    if os.path.exists(wb_path):
        wb = load_workbook(wb_path)
    else:
        wb = Workbook()

    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb.active
        ws.title = SHEET_NAME

    # Write headers if missing / mismatched
    existing = [ws.cell(row=1, column=i + 1).value for i in range(len(HEADERS))]
    if existing != HEADERS:
        for i, h in enumerate(HEADERS, start=1):
            ws.cell(row=1, column=i).value = h

    return wb, ws


def append_receipt_row(
    ws: Worksheet,
    receipt_date: Optional[date],
    payee: Optional[str],
    description: Optional[str],
    gst: Optional[Decimal],
    total: Optional[Decimal],
) -> None:
    """
    Appends a single row in the required column order.
    Entertainment is intentionally left blank (for your spreadsheet formula).
    """
    next_row = ws.max_row + 1

    # Date
    c_date = ws.cell(row=next_row, column=1, value=receipt_date)
    c_date.number_format = DATE_FMT

    # Payee, Description
    ws.cell(row=next_row, column=2, value=(payee or "").strip() or None)
    ws.cell(row=next_row, column=3, value=(description or "").strip() or None)

    # Entertainment (blank)
    ws.cell(row=next_row, column=4, value=None)

    # GST, Total
    c_gst = ws.cell(row=next_row, column=5, value=float(gst) if gst is not None else None)
    c_gst.number_format = CURRENCY_FMT

    c_total = ws.cell(row=next_row, column=6, value=float(total) if total is not None else None)
    c_total.number_format = CURRENCY_FMT


def extract_receipt_fields(client: OpenAI, image_path: str, model: str) -> Dict[str, Any]:
    """
    Calls the OpenAI Responses API with:
      - image input (base64 data URL)
      - Structured Outputs (JSON schema) to force consistent fields
    """
    data_url = image_path_to_data_url(image_path)

    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "date_iso": {
                "type": ["string", "null"],
                "description": "Receipt date in YYYY-MM-DD format. Null if not visible."
            },
            "payee": {
                "type": ["string", "null"],
                "description": "Store/vendor name from receipt."
            },
            "description": {
                "type": "string",
                "enum": ["Meal", "Treat"],
                "description": "Meal: food for immediate consumption (dine-in/takeout). Treat: all other purchases."
            },
            "gst": {
                "type": ["number", "null"],
                "description": "GST amount from receipt (do not calculate). Null if not shown."
            },
            "total": {
                "type": ["number", "null"],
                "description": "Total amount paid, including tip if present."
            },
        },
        "required": ["date_iso", "payee", "description", "gst", "total"],
    }

    resp = client.responses.create(
        model=model,
        input=[
            {
                "role": "system",
                "content": "Extract receipt data accurately. Use exact values from receipt."
            },
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": "Extract all fields from this receipt."},
                    {"type": "input_image", "image_url": data_url},
                ],
            },
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "receipt_row",
                "schema": schema,
                "strict": True,
            }
        },
    )

    # When using Structured Outputs, the JSON is returned as text you can parse.
    raw = (resp.output_text or "").strip()
    if not raw:
        return {
            "date_iso": None,
            "payee": None,
            "description": None,
            "gst": None,
            "total": None,
        }

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {
            "date_iso": None,
            "payee": None,
            "description": None,
            "gst": None,
            "total": None,
        }


def get_images(folder: str = "receipt_images", limit: Optional[int] = None) -> list[str]:
    """Get image files from the specified folder."""
    if not os.path.exists(folder):
        raise FileNotFoundError(f"Folder '{folder}' not found.")
    
    image_exts = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
    images = []
    for filename in sorted(os.listdir(folder), reverse=True):
        if os.path.splitext(filename.lower())[1] in image_exts:
            images.append(os.path.join(folder, filename))
            if limit and len(images) >= limit:
                break
    
    if not images:
        raise FileNotFoundError(f"No image files found in '{folder}'.")
    
    return images


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Extract receipt fields from images and append to an .xlsx sheet.")
    parser.add_argument("-n", "--num", type=int, default=0, help="Number of images to process from receipt_images/ (default: 1, use 0 for all).")
    parser.add_argument("--xlsx", default="receipts.xlsx", help="Output workbook path (default: receipts.xlsx).")
    parser.add_argument("--model", default=os.getenv("RECEIPT_MODEL", "gpt-4o-mini"))
    args = parser.parse_args()

    # Get images from receipt_images/
    limit = None if args.num == 0 else args.num
    images = get_images(limit=limit)
    
    print(f"Processing {len(images)} image(s)...")

    client = OpenAI()
    wb, ws = ensure_sheet_and_headers(args.xlsx)

    for i, image_path in enumerate(images, 1):
        print(f"[{i}/{len(images)}] Processing: {image_path}")
        
        extracted = extract_receipt_fields(client, image_path, args.model)

        # Parse/normalize for spreadsheet
        receipt_date = None
        if extracted.get("date_iso"):
            try:
                receipt_date = datetime.strptime(extracted["date_iso"], "%Y-%m-%d").date()
            except ValueError:
                pass  # Leave as None if date parsing fails

        gst = safe_decimal(extracted.get("gst"))
        total = safe_decimal(extracted.get("total"))

        append_receipt_row(
            ws=ws,
            receipt_date=receipt_date,
            payee=extracted.get("payee"),
            description=extracted.get("description"),
            gst=gst,
            total=total,
        )

    wb.save(args.xlsx)
    print(f"\nAppended {len(images)} row(s) to {args.xlsx} ({SHEET_NAME}).")


if __name__ == "__main__":
    main()
