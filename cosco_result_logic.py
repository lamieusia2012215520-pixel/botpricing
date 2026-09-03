"""Side-effect-free COSCO workbook output helpers."""


COSCO_QUOTE_OUTPUT_COLUMNS = (6, 7, 8, 9, 10, 11, 15, 16)


def clear_cosco_quote_fields(sheet, row: int) -> None:
    """Remove stale price/schedule/vessel values after a definitive no-price result."""
    for column in COSCO_QUOTE_OUTPUT_COLUMNS:
        sheet.cell(row=row, column=column).value = None
