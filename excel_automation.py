import re


def remove_excel_file_references(cell_value):
    if not isinstance(cell_value, str) or not cell_value.startswith("="):
        return cell_value

    def format_sheet_reference(sheet_name):
        """
        Format a sheet name according to Excel's syntax.
        """
        # Excel requires quotes for sheet names containing spaces
        # or other characters that make an unquoted reference invalid.
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet_name):
            return f"{sheet_name}!"

        # Apostrophes inside a quoted Excel sheet name are doubled.
        sheet_name = sheet_name.replace("'", "''")
        return f"'{sheet_name}'!"

    # ------------------------------------------------------------
    # Case 1:
    # 'C:\folder\[testA.xlsx]Sheet 2'!$C$2
    #
    # The entire workbook path + workbook name + sheet name
    # is enclosed in single quotes.
    # ------------------------------------------------------------
    quoted_pattern = re.compile(
        r"""
        '
        [^']*                # path / other content before [workbook]
        \[
        [^\]]+
        \]
        ([^']+)               # sheet name
        '
        !
        """,
        re.VERBOSE,
    )

    def replace_quoted(match):
        sheet_name = match.group(1)
        return format_sheet_reference(sheet_name)

    formula = quoted_pattern.sub(replace_quoted, cell_value)

    # ------------------------------------------------------------
    # Case 2:
    # [testA.xlsx]Sheet2!$D$2
    #
    # Unquoted external reference.
    # ------------------------------------------------------------
    unquoted_pattern = re.compile(
        r"""
        \[
        [^\]]+
        \]
        ([A-Za-z_][A-Za-z0-9_.]*)
        !
        """,
        re.VERBOSE,
    )

    def replace_unquoted(match):
        sheet_name = match.group(1)
        return format_sheet_reference(sheet_name)

    formula = unquoted_pattern.sub(replace_unquoted, formula)

    return formula