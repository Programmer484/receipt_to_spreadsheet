"""Configuration for receipt extraction schema and prompts."""

# Spreadsheet configuration
HEADERS = ["Date", "Description", "Amount"]
DATE_FMT = "d-mmm-yy"
CURRENCY_FMT = '"$"#,##0.00'

# Field mapping: column_name -> (json_key, excel_column_index, format_type)
# format_type: "date", "currency", "text", or None for blank
FIELD_MAPPING = [
    ("Date", "date_from_filename", 1, "date"),
    ("Description", "description", 2, "text"),
    ("Amount", "amount", 3, "currency"),
]

# JSON Schema for OpenAI structured outputs
SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "description": {
            "type": ["string", "null"],
            "description": "Top 3 most expensive items from receipt, comma-separated. Add 'etc' if more than 3 items. Null if not visible."
        },
        "amount": {
            "type": ["number", "null"],
            "description": "Total amount paid. Null if not visible."
        },
    },
    "required": ["description", "amount"],
}

# Prompts
SYSTEM_PROMPT = "Extract receipt data accurately. Use exact values from receipt."
USER_PROMPT = "Extract the top 3 most expensive items and total amount from this receipt."

# Sheet configurations: folder_path -> sheet_name
SHEET_CONFIGS = [
    {"folder": "819", "sheet_name": "819"},
    {"folder": "1705-07", "sheet_name": "1705-07"},
    {"folder": "1712", "sheet_name": "1712"},
]
