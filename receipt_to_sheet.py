import os
import json
import base64
import mimetypes
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import tempfile
from typing import Any, Dict, Optional

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from pdf2image import convert_from_path
from natsort import natsorted
from PyPDF2 import PdfReader

from config import (
    HEADERS, DATE_FMT, CURRENCY_FMT,
    FIELD_MAPPING, SCHEMA, SYSTEM_PROMPT, USER_PROMPT_TEMPLATE, SHEET_CONFIGS
)

load_dotenv()


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


def ensure_sheet_and_headers(wb_path: str, headers: list[str], sheet_name: str) -> tuple[Any, Worksheet]:
    """
    Creates/opens the workbook and ensures the target sheet + headers exist.
    """
    if os.path.exists(wb_path):
        wb = load_workbook(wb_path)
    else:
        wb = Workbook()

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active
        ws.title = sheet_name

    # Write headers if missing / mismatched
    existing = [ws.cell(row=1, column=i + 1).value for i in range(len(headers))]
    if existing != headers:
        for i, h in enumerate(headers, start=1):
            ws.cell(row=1, column=i).value = h

    return wb, ws


def append_row(
    ws: Worksheet,
    extracted: Dict[str, Any],
    field_mapping: list[tuple],
    date_fmt: str,
    currency_fmt: str,
) -> None:
    """
    Appends a row based on field mapping configuration.
    """
    next_row = ws.max_row + 1

    for header, json_key, col_idx, format_type in field_mapping:
        cell = ws.cell(row=next_row, column=col_idx)
        
        if json_key is None:
            # Blank cell (e.g., Entertainment)
            cell.value = None
        elif format_type == "date":
            # Parse date
            date_str = extracted.get(json_key)
            if date_str:
                try:
                    cell.value = datetime.strptime(date_str, "%Y-%m-%d").date()
                    cell.number_format = date_fmt
                except ValueError:
                    cell.value = None
            else:
                cell.value = None
        elif format_type == "currency":
            # Currency with formatting
            val = extracted.get(json_key)
            decimal_val = safe_decimal(val)
            cell.value = float(decimal_val) if decimal_val is not None else None
            cell.number_format = currency_fmt
        else:
            # Text
            val = extracted.get(json_key)
            cell.value = (val or "").strip() or None if isinstance(val, str) else val


def extract_receipt_fields(
    client: OpenAI,
    model: str,
    schema: Dict[str, Any],
    system_prompt: str,
    user_prompt: str,
    *,
    image_path: Optional[str] = None,
    text_content: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Calls the OpenAI Responses API with:
      - image input (base64 data URL)
      - Structured Outputs (JSON schema) to force consistent fields
    """
    content = [{"type": "input_text", "text": user_prompt}]
    if image_path:
        data_url = image_path_to_data_url(image_path)
        content.append({"type": "input_image", "image_url": data_url})
    elif text_content is not None:
        content.append({"type": "input_text", "text": f"Receipt text:\n{text_content}"})

    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": content,
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

    raw = (resp.output_text or "").strip()
    if not raw:
        return {key: None for key in schema.get("required", [])}

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {key: None for key in schema.get("required", [])}


def get_images(folder: str = "receipt_images", limit: Optional[int] = None) -> list[str]:
    """Get image and PDF files from the specified folder."""
    if not os.path.exists(folder):
        raise FileNotFoundError(f"Folder '{folder}' not found.")
    
    image_exts = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".pdf"}
    images: list[str] = []
    for filename in natsorted(os.listdir(folder), reverse=True):
        if os.path.splitext(filename.lower())[1] in image_exts:
            images.append(os.path.join(folder, filename))
            if limit and len(images) >= limit:
                break
    
    if not images:
        raise FileNotFoundError(f"No image files found in '{folder}'.")
    
    return images


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from a PDF file."""
    reader = PdfReader(pdf_path)
    chunks = []
    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():
            chunks.append(text)
    return "\n".join(chunks).strip()


def pdf_to_images(pdf_path: str, output_dir: str, max_pages: int = 1) -> list[str]:
    """Render PDF pages to images (PNG) and return their paths."""
    images = convert_from_path(pdf_path, first_page=1, last_page=max_pages)
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    paths = []
    for idx, img in enumerate(images, 1):
        out_path = os.path.join(output_dir, f"{stem}_page{idx}.png")
        img.save(out_path, "PNG")
        paths.append(out_path)
    return paths


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Extract receipt fields from images and append to an .xlsx sheet.")
    parser.add_argument("-n", "--num", type=int, default=0, help="Number of images to process per folder (default: 0 for all).")
    parser.add_argument("--xlsx", default="receipts.xlsx", help="Output workbook path (default: receipts.xlsx).")
    parser.add_argument("--model", default=os.getenv("RECEIPT_MODEL", "gpt-4o-mini"))
    parser.add_argument("--folder", help="Process a single folder (provide folder path). If omitted, processes all folders from config.")
    parser.add_argument("--sheet", help="Sheet name (required if --folder is specified).")
    args = parser.parse_args()

    client = OpenAI()
    limit = None if args.num == 0 else args.num

    # Single folder mode or batch mode
    if args.folder:
        if not args.sheet:
            print("Error: --sheet is required when using --folder")
            return
        configs = [{"folder": args.folder, "sheet_name": args.sheet}]
    else:
        configs = SHEET_CONFIGS

    # Open/create workbook once
    if os.path.exists(args.xlsx):
        wb = load_workbook(args.xlsx)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    total_processed = 0

    for config in configs:
        folder = config["folder"]
        sheet_name = config["sheet_name"]

        try:
            images = get_images(folder=folder, limit=limit)
        except FileNotFoundError as e:
            print(f"Skipping {folder}: {e}")
            continue

        print(f"\nProcessing {len(images)} image(s) from '{folder}' -> sheet '{sheet_name}'...")

        # Ensure sheet and headers
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
        
        # Write headers
        existing = [ws.cell(row=1, column=i + 1).value for i in range(len(HEADERS))]
        if existing != HEADERS:
            for i, h in enumerate(HEADERS, start=1):
                ws.cell(row=1, column=i).value = h

        for i, file_path in enumerate(images, 1):
            filename = os.path.basename(file_path)
            print(f"  [{i}/{len(images)}] Processing: {filename}")

            user_prompt = USER_PROMPT_TEMPLATE.format(title=filename)
            ext = os.path.splitext(filename.lower())[1]
            if ext == ".pdf":
                text_content = extract_text_from_pdf(file_path)
                if text_content:
                    extracted = extract_receipt_fields(
                        client,
                        args.model,
                        SCHEMA,
                        SYSTEM_PROMPT,
                        user_prompt,
                        text_content=text_content,
                    )
                else:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        image_paths = pdf_to_images(file_path, tmpdir, max_pages=1)
                        if not image_paths:
                            print(f"    Skipping {filename}: PDF has no text and no renderable pages.")
                            continue
                        extracted = extract_receipt_fields(
                            client,
                            args.model,
                            SCHEMA,
                            SYSTEM_PROMPT,
                            user_prompt,
                            image_path=image_paths[0],
                        )
            else:
                extracted = extract_receipt_fields(
                    client,
                    args.model,
                    SCHEMA,
                    SYSTEM_PROMPT,
                    user_prompt,
                    image_path=file_path,
                )

            append_row(ws, extracted, FIELD_MAPPING, DATE_FMT, CURRENCY_FMT)
            total_processed += 1

    wb.save(args.xlsx)
    print(f"\n✓ Total: {total_processed} receipt(s) processed across {len(configs)} sheet(s) in {args.xlsx}")


if __name__ == "__main__":
    main()
