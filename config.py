"""Configuration for receipt extraction schema and prompts."""

# Spreadsheet configuration
HEADERS = ["Date", "Description", "House Number", "Amount"]
DATE_FMT = "d-mmm-yy"
CURRENCY_FMT = '"$"#,##0.00'

# Field mapping: column_name -> (json_key, excel_column_index, format_type)
# format_type: "date", "currency", "text", or None for blank
FIELD_MAPPING = [
    ("Date", "date_iso", 1, "date"),
    ("Description", "description", 2, "text"),
    ("House Number", "house_number", 3, "text"),
    ("Amount", "amount", 4, "currency"),
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
        "description": {
            "type": ["string", "null"],
            "description": "Top 3 most expensive items from receipt, comma-separated. Add 'etc' if more than 3 items. Null if not visible."
        },
        "house_number": {
            "type": ["string", "null"],
            "description": "House number if present in the receipt title/filename. Null if not visible."
        },
        "amount": {
            "type": ["number", "null"],
            "description": "Total amount paid. Null if not visible."
        },
    },
    "required": ["date_iso", "description", "house_number", "amount"],
}

# Prompts
SYSTEM_PROMPT = (
    "Extract receipt data accurately. Use exact values from the receipt content. "
    "Use the title/filename for date, description, and house number when available."
)
USER_PROMPT_TEMPLATE = (
    "Receipt title/filename: {title}\n\n"
    "Rules:\n"
    "- Extract date, description (top 3 items + 'etc' if more), and house number from the title/filename when possible.\n"
    "- Extract only the total amount from the receipt content.\n"
    "- If any field is not visible, return null.\n"
    "Return only the structured JSON."
)

# Sheet configurations: folder_path -> sheet_name
SHEET_CONFIGS = [
    {"folder": "receipts", "sheet_name": "e-Receipts"},
]
