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

HEADERS = ["Date", "Payee", "Description", "Entertainment", "GST", "Total", "Notes"]
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
    notes: Optional[str],
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

    # Notes
    ws.cell(row=next_row, column=7, value=(notes or "").strip() or None)


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
            # ISO keeps parsing simple; we format for Excel later
            "date_iso": {
                "type": ["string", "null"],
                "description": "Receipt date in ISO 8601 format YYYY-MM-DD (from the receipt). Null if not visible."
            },
            "payee": {
                "type": ["string", "null"],
                "description": "Store/vendor name as shown on the receipt."
            },
            "description": {
                "type": ["string", "null"],
                "description": "Short expense description/category (e.g., Meal, Groceries, Taxi)."
            },
            "gst": {
                "type": ["number", "null"],
                "description": "GST amount exactly as shown on the receipt (do not compute). Use 0 only if explicitly shown as 0."
            },
            "total": {
                "type": ["number", "null"],
                "description": "Total paid INCLUDING tip (if tip is present). Use the final charged/paid amount."
            },
            "notes": {
                "type": ["string", "null"],
                "description": "If any field is uncertain/ambiguous/missing, explain briefly here."
            },
        },
        "required": ["date_iso", "payee", "description", "gst", "total", "notes"],
    }

    # Responses API supports image + text inputs. (docs)
    resp = client.responses.create(
        model=model,
        input=[
            {
                "role": "system",
                "content": (
                    "You are a meticulous bookkeeping assistant. "
                    "Extract receipt fields exactly from the receipt. "
                    "Do not guess missing values. If uncertain, put details in notes."
                ),
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": (
                            "Extract fields for my expense spreadsheet.\n"
                            "- Total MUST include tip (if any).\n"
                            "- GST must come straight from the receipt (do not compute).\n"
                            "- Payee can include the store name.\n"
                            "Return only the structured JSON."
                        ),
                    },
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
            "notes": "Model returned empty output_text (possible refusal or unexpected response).",
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
            "notes": f"Failed to parse JSON. Raw output_text: {raw[:500]}",
        }


def get_first_image(folder: str = "receipt_images") -> str:
    """Get the first image file from the specified folder."""
    if not os.path.exists(folder):
        raise FileNotFoundError(f"Folder '{folder}' not found.")
    
    image_exts = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
    for filename in sorted(os.listdir(folder)):
        if os.path.splitext(filename.lower())[1] in image_exts:
            return os.path.join(folder, filename)
    
    raise FileNotFoundError(f"No image files found in '{folder}'.")


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Extract receipt fields from one image and append to an .xlsx sheet.")
    parser.add_argument("image_path", nargs="?", help="Path to receipt image (jpg/png/webp). If omitted, uses first image from receipt_images/.")
    parser.add_argument("--xlsx", default="receipts.xlsx", help="Output workbook path (default: receipts.xlsx).")
    parser.add_argument("--model", default=os.getenv("RECEIPT_MODEL", "gpt-4.1"),
                        help="OpenAI model (default: gpt-4.1 or env RECEIPT_MODEL).")
    args = parser.parse_args()

    # Use provided path or get first image from receipt_images/
    image_path = args.image_path or get_first_image()
    print(f"Processing: {image_path}")

    client = OpenAI()

    extracted = extract_receipt_fields(client, image_path, args.model)

    # Parse/normalize for spreadsheet
    receipt_date = None
    if extracted.get("date_iso"):
        try:
            receipt_date = datetime.strptime(extracted["date_iso"], "%Y-%m-%d").date()
        except ValueError:
            extracted["notes"] = (extracted.get("notes") or "") + f" | Bad date_iso: {extracted['date_iso']}"

    gst = safe_decimal(extracted.get("gst"))
    total = safe_decimal(extracted.get("total"))

    # Quick sanity check: GST > Total is almost certainly wrong
    notes = extracted.get("notes")
    if gst is not None and total is not None and gst > total:
        notes = (notes or "")
        notes += " | Sanity check: GST > Total; please verify receipt."

    wb, ws = ensure_sheet_and_headers(args.xlsx)
    append_receipt_row(
        ws=ws,
        receipt_date=receipt_date,
        payee=extracted.get("payee"),
        description=extracted.get("description"),
        gst=gst,
        total=total,
        notes=notes,
    )
    wb.save(args.xlsx)

    print(f"Appended 1 row to {args.xlsx} ({SHEET_NAME}).")


if __name__ == "__main__":
    main()
