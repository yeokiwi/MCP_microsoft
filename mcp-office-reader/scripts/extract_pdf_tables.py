#!/usr/bin/env python3
"""
PDF table extraction script using pdfplumber.
Usage: python3 extract_pdf_tables.py <pdf_file_path>
Output: JSON array of tables with page number and rows
"""

import sys
import json


def extract_tables(pdf_path: str) -> list:
    try:
        import pdfplumber
    except ImportError:
        print(
            json.dumps(
                {
                    "error": "pdfplumber is not installed. Run: pip install pdfplumber"
                }
            )
        )
        sys.exit(1)

    output = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                try:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            # Clean up None values and normalize rows
                            cleaned_rows = []
                            for row in table:
                                cleaned_row = [
                                    cell if cell is not None else ""
                                    for cell in row
                                ]
                                cleaned_rows.append(cleaned_row)
                            output.append(
                                {"page": i + 1, "rows": cleaned_rows}
                            )
                except Exception as page_err:
                    # Skip pages that fail, continue with others
                    sys.stderr.write(
                        f"Warning: Could not extract tables from page {i + 1}: {page_err}\n"
                    )
    except Exception as e:
        print(json.dumps({"error": f"Failed to open PDF: {str(e)}"}))
        sys.exit(1)

    print(json.dumps(output))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: extract_pdf_tables.py <pdf_path>"}))
        sys.exit(1)

    extract_tables(sys.argv[1])
