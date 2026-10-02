import re


def remove_excel_file_references(cell_value):
    if not isinstance(cell_value, str):
        return cell_value

    if not cell_value.startswith("="):
        return cell_value

    # Remove external file references such as:
    # 'C:\folder\[testA.xlsx]Sheet2'!$C$2
    # [testA.xlsx]Sheet2!$D$2
    #
    # Keep the sheet name and the !.
    pattern = r"(?:'[^']*\[[^\]]+\]([^']+)'|\[[^\]]+\]([A-Za-z0-9_ ]+))!"

    def replace_reference(match):
        sheet_name = match.group(1) or match.group(2)
        return f"{sheet_name}!"

    return re.sub(pattern, replace_reference, cell_value)