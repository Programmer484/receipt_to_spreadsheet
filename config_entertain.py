"""Configuration for receipt extraction schema and prompts."""

# Spreadsheet configuration
HEADERS = ["Date", "Payee", "Description", "Entertainment", "GST", "Total"]
SHEET_NAME = "Receipts"
DATE_FMT = "d-mmm-yyyy"
CURRENCY_FMT = '"$"#,##0.00'

# Field mapping: column_name -> (json_key, excel_column_index, format_type)
# format_type: "date", "currency", "text", or None for blank
FIELD_MAPPING = [
    ("Date", "date_iso", 1, "date"),
    ("Payee", "payee", 2, "text"),
    ("Description", "description", 3, "text"),
    ("Entertainment", None, 4, None),  # Always blank
    ("GST", "gst", 5, "currency"),
    ("Total", "total", 6, "currency"),
]

# JSON Schema for OpenAI structured outputs
SCHEMA = {
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

# Prompts
SYSTEM_PROMPT = "Extract receipt data accurately. Use exact values from receipt."
USER_PROMPT = "Extract all fields from this receipt."
