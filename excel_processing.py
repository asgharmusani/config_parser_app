# -*- coding: utf-8 -*-
"""
Standalone Built-in Parser for processing an original source Excel file.

This module takes an original Excel workbook, identifies common routing entities
(VQs, Skills, VAGs, Skill Expressions) based on predefined internal constants,
discards any struck-through items, performs cleaning, and outputs a new
workbook object with standardized sheets for each entity type.

This script is designed to be potentially user-editable for its parsing logic
if the default behavior needs adjustment for specific source Excel formats.
It aims to be self-contained for its parsing duties.
"""

import logging
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import cell as openpyxl_cell_utils
import re
from typing import Dict, Any, Optional, Tuple, Set, List

# --- Logger for this module ---
# If run standalone, this will configure a basic logger.
# If imported by Flask app, it will use the app's logger config.
logger = logging.getLogger(__name__)
if not logger.hasHandlers(): # Avoid adding handlers if already configured by Flask
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s')

# --- Default Configuration Constants for the Built-in Parser ---
# Users can modify these directly if their source Excel has different conventions.
DEFAULT_IDEAL_AGENT_HEADER_TEXT = "Ideal Agent"
# List of cell addresses to check for the Ideal Agent header text, in order of preference.
DEFAULT_IDEAL_AGENT_CELL_ADDRESSES = ["C1", "D1", "C2"]
DEFAULT_VAG_EXTRACTION_SHEET_NAME = "Default Targeting- Group"
# Sheets to always skip when this parser processes an original Excel file
DEFAULT_SHEETS_TO_SKIP_PARSING = {
    "Instructions", "Summary", "Metadata",
    # Avoid re-processing sheets that might be named like its own output
    "VQs", "Skills", "Skill_Expressions", "VAGs"
}


# --- Internalized Utility Functions ---
def _extract_skills_from_expression(expression: str) -> list[str]:
    """
    Extracts potential skill names (alphanumeric + underscore) from a skill
    expression string. Looks for patterns like 'SkillName>5'.
    Internal to this module now.
    """
    if not isinstance(expression, str):
        return []
    # Regex finds words (alphanumeric or underscore) followed immediately by '>' and one or more digits
    skills = re.findall(r'\b([a-zA-Z0-9_]+)(?=>\d+)', expression)
    logger.debug(f"Extracted skills {skills} from expression '{expression}'")
    return skills

def _identify_ideal_agent_column(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    ideal_agent_header_text: str,
    ideal_agent_cell_addresses: List[str] # Now expects a list of cell addresses
) -> Optional[int]:
    """
    Identifies the column index for the 'Ideal Agent' column.
    Iterates through the provided list of cell addresses to find the
    specified header text. The column of the first matching cell is returned.

    Args:
        sheet: The openpyxl worksheet object.
        ideal_agent_header_text: The header text to search for.
        ideal_agent_cell_addresses: A list of cell coordinates (e.g., "C1", "D2")
                                     to check in order.

    Returns:
        The 1-based column index if the header is found in one of the specified cells,
        otherwise None.
    """
    logger.debug(f"Identifying '{ideal_agent_header_text}' column in sheet: {sheet.title} by checking cell addresses: {ideal_agent_cell_addresses}")

    # --- MODIFICATION START: Iterate through cell addresses and parse directly ---
    for cell_address in ideal_agent_cell_addresses:
        try:
            # Parse the cell address to get column string and row number
            col_str, row_idx_from_address = openpyxl_cell_utils.coordinate_to_tuple(cell_address)
            # Convert column string (e.g., "C") to column index (e.g., 3)
            col_idx_to_check = openpyxl_cell_utils.column_index_from_string(col_str)

            # Check if the parsed cell address is within the sheet's bounds
            if row_idx_from_address <= sheet.max_row and col_idx_to_check <= sheet.max_column:
                cell_value = sheet.cell(row=row_idx_from_address, column=col_idx_to_check).value
                if cell_value and ideal_agent_header_text in str(cell_value):
                    logger.debug(f"Found '{ideal_agent_header_text}' at cell '{cell_address}' (Column {col_idx_to_check}). Using this column for 'Ideal Agent' data.")
                    return col_idx_to_check # Return the column index where the header was found
            else:
                logger.debug(f"Cell address '{cell_address}' is out of bounds for sheet '{sheet.title}'.")
        except openpyxl_cell_utils.IllegalCharacterError:
            logger.warning(f"Invalid cell address format in ideal_agent_cell_addresses: '{cell_address}'. Skipping this address.")
        except Exception as e:
             logger.warning(f"Could not parse or check ideal agent location '{cell_address}': {e}")
    # --- MODIFICATION END ---

    logger.debug(f"'{ideal_agent_header_text}' column not found in sheet: {sheet.title} using configured cell addresses.")
    return None


# --- Main Parser Function ---
def parse_source_excel_to_standardized_workbook(
    source_workbook: openpyxl.workbook.Workbook
) -> openpyxl.workbook.Workbook:
    """
    Parses the original source Excel workbook to extract entities based on built-in logic
    defined by internal constants.
    It discards struck-through items, cleans data, and creates a new
    workbook object with standardized output sheets for each entity type.

    Args:
        source_workbook: The openpyxl.Workbook object of the original uploaded Excel.
                         Expected to be loaded with style information to detect strikethrough.
    Returns:
        A new openpyxl.Workbook object containing the parsed and standardized entity sheets.
    """
    logger.info("Starting built-in parsing of source workbook to create standardized entity sheets...")

    # Use internal constants directly
    ideal_agent_header = DEFAULT_IDEAL_AGENT_HEADER_TEXT
    ideal_agent_cell_addrs = DEFAULT_IDEAL_AGENT_CELL_ADDRESSES # Use the list
    vag_sheet_name_hint = DEFAULT_VAG_EXTRACTION_SHEET_NAME
    sheets_to_skip = DEFAULT_SHEETS_TO_SKIP_PARSING.copy()

    # Intermediate storage for unique, non-struck entities
    parsed_data = {
        "VQs": set(),
        "Skills": set(),
        "Skill_Expressions": [], # List of dicts for structure
        "VAGs": set()
    }

    # --- Iterate through sheets and cells of the source workbook ---
    for sheet in source_workbook.worksheets:
        if sheet.title in sheets_to_skip:
            logger.debug(f"Skipping sheet during parsing: {sheet.title} (in predefined skip list)")
            continue

        logger.info(f"Built-in parser processing sheet: {sheet.title}")
        ideal_agent_col_idx = _identify_ideal_agent_column(
            sheet,
            ideal_agent_header,
            ideal_agent_cell_addrs # Pass the list of addresses
        )

        for row_idx in range(1, sheet.max_row + 1):
            for col_idx in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                if cell.value is None:
                    continue
                
                value_str_raw = str(cell.value)
                value_str_stripped = value_str_raw.strip()

                if not value_str_stripped:
                    continue

                if cell.font and cell.font.strike:
                    logger.debug(f"Skipping struck-through cell {cell.coordinate}: '{value_str_raw}'")
                    continue

                if (value_str_stripped.lower().startswith("vq_") or "vq" in value_str_stripped.lower()) and ">" not in value_str_stripped:
                    cleaned_vq = value_str_stripped.replace(" ", "").replace('\u00A0', '')
                    if cleaned_vq:
                        parsed_data["VQs"].add(cleaned_vq)
                        logger.debug(f"Parser found VQ: {cleaned_vq} from {cell.coordinate}")

                elif ">" in value_str_stripped:
                    raw_expression = value_str_stripped
                    ideal_expression_str = ""
                    if ideal_agent_col_idx and ideal_agent_col_idx <= sheet.max_column:
                        ideal_cell = sheet.cell(row=row_idx, column=ideal_agent_col_idx)
                        if ideal_cell.value is not None:
                            if not (ideal_cell.font and ideal_cell.font.strike):
                                ideal_expression_str = str(ideal_cell.value).strip()
                            else:
                                logger.debug(f"Ideal Agent cell {ideal_cell.coordinate} is struck-through for expression '{raw_expression}'. Ignoring ideal part.")

                    cleaned_expression = raw_expression.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                    cleaned_ideal = ideal_expression_str.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                    concatenated_key = cleaned_expression
                    if cleaned_ideal: concatenated_key = f"{cleaned_expression} {cleaned_ideal}".strip()
                    extracted_skills_list_for_this_expr = []
                    if cleaned_expression:
                        skills_from_expr = _extract_skills_from_expression(cleaned_expression)
                        for skill in skills_from_expr:
                            cleaned_skill = skill.replace(" ", "").replace('\u00A0', '')
                            if cleaned_skill:
                                parsed_data["Skills"].add(cleaned_skill)
                                extracted_skills_list_for_this_expr.append(cleaned_skill)
                    if cleaned_expression:
                        parsed_data["Skill_Expressions"].append({
                            "Original Expression": cleaned_expression, "Ideal Expression": cleaned_ideal,
                            "Concatenated Key": concatenated_key,
                            "Extracted_Skills_List_String": ", ".join(extracted_skills_list_for_this_expr)
                        })
                        logger.debug(f"Parser found Skill Expression: {concatenated_key} from {cell.coordinate}")

                elif value_str_stripped.startswith("VAG_") and sheet.title == vag_sheet_name_hint:
                    cleaned_vag = value_str_stripped.replace(" ", "").replace('\u00A0', '')
                    if cleaned_vag:
                        parsed_data["VAGs"].add(cleaned_vag)
                        logger.debug(f"Parser found VAG: {cleaned_vag} from {cell.coordinate}")
                
    logger.info("Built-in parser finished initial data extraction from source workbook.")

    output_workbook = openpyxl.Workbook()
    if "Sheet" in output_workbook.sheetnames and len(output_workbook.sheetnames) == 1:
        try: output_workbook.remove(output_workbook.active)
        except Exception as e_rm_sheet: logger.warning(f"Could not remove default sheet: {e_rm_sheet}")

    bold_font = Font(bold=True)
    if parsed_data["VQs"]:
        vq_sheet = output_workbook.create_sheet("VQs")
        vq_sheet.cell(row=1, column=1, value="VQ Name").font = bold_font
        for i, vq_name in enumerate(sorted(list(parsed_data["VQs"])), start=2): vq_sheet.cell(row=i, column=1, value=vq_name)
        logger.info(f"Created 'VQs' output sheet with {len(parsed_data['VQs'])} items.")
    if parsed_data["Skills"]:
        skill_sheet = output_workbook.create_sheet("Skills")
        skill_sheet.cell(row=1, column=1, value="Skill Name").font = bold_font
        for i, skill_name in enumerate(sorted(list(parsed_data["Skills"])), start=2): skill_sheet.cell(row=i, column=1, value=skill_name)
        logger.info(f"Created 'Skills' output sheet with {len(parsed_data['Skills'])} items.")
    if parsed_data["VAGs"]:
        vag_sheet = output_workbook.create_sheet("VAGs_Output")
        vag_sheet.cell(row=1, column=1, value="VAG Name").font = bold_font
        for i, vag_name in enumerate(sorted(list(parsed_data["VAGs"])), start=2): vag_sheet.cell(row=i, column=1, value=vag_name)
        logger.info(f"Created 'VAGs' sheet with {len(parsed_data['VAGs'])} items.")
    if parsed_data["Skill_Expressions"]:
        se_sheet = output_workbook.create_sheet("Skill_Expressions_Output")
        se_headers = ["Original Expression", "Ideal Expression", "Concatenated Key", "Extracted_Skills_List_String"]
        for col_idx, header in enumerate(se_headers, start=1): se_sheet.cell(row=1, column=col_idx, value=header).font = bold_font
        sorted_skill_expressions = sorted(parsed_data["Skill_Expressions"], key=lambda x: x.get("Concatenated Key", ""))
        for row_idx, se_data in enumerate(sorted_skill_expressions, start=2):
            se_sheet.cell(row=row_idx, column=1, value=se_data.get("Original Expression"))
            se_sheet.cell(row=row_idx, column=2, value=se_data.get("Ideal Expression"))
            se_sheet.cell(row=row_idx, column=3, value=se_data.get("Concatenated Key"))
            se_sheet.cell(row=row_idx, column=4, value=se_data.get("Extracted_Skills_List_String"))
        logger.info(f"Created 'Skill_Expressions' sheet with {len(parsed_data['Skill_Expressions'])} items.")

    logger.info("Built-in parser finished creating standardized output workbook object.")
    return output_workbook


# Example usage if run as a standalone script (for testing the parser)
if __name__ == '__main__': # pragma: no cover
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s')
    logger.info("Testing excel_processing.py standalone parser.")
    test_wb_path = "test_source_excel.xlsx"
    wb = openpyxl.Workbook()
    sheet1 = wb.active; sheet1.title = "Sheet1_VQs_Skills"
    sheet1['A1'] = "VQ_SALES_EN"; sheet1['A2'] = "VQ_SUPPORT_ES"; sheet1['A2'].font = Font(strike=True)
    sheet1['B1'] = "SkillA>5 & SkillB>3"; sheet1['C1'] = "Ideal Agent Text Here" # Put header text in C1
    sheet1['B2'] = "SkillC>2"; sheet1['A3'] = "VQ_Billing"; sheet1['A4'] = "  VQ_ espa\u00A0ce  "
    sheet2 = wb.create_sheet("Default Targeting- Group")
    sheet2['A1'] = "VAG_Tier1_Support"; sheet2['A2'] = "VAG_Sales_VIP"; sheet2['A2'].font = Font(strike=True)
    wb.save(test_wb_path)
    logger.info(f"Created dummy test workbook: {test_wb_path}")
    loaded_test_wb = openpyxl.load_workbook(test_wb_path, data_only=False, read_only=False) # Need styles for strike
    
    # Test with the new parse_source_excel_to_standardized_workbook which uses internal constants
    processed_wb = parse_source_excel_to_standardized_workbook(loaded_test_wb) # No config_hints needed
    loaded_test_wb.close()
    processed_output_path = "test_source_excel_PARSED.xlsx"
    processed_wb.save(processed_output_path)
    logger.info(f"Saved parsed output to: {processed_output_path}")
    verify_wb = openpyxl.load_workbook(processed_output_path)
    print("\nGenerated Sheets:");
    for sheetname in verify_wb.sheetnames:
        print(f"- {sheetname}"); ws = verify_wb[sheetname]
        for row in ws.iter_rows(values_only=True): print(f"  {row}")
    verify_wb.close()
    print(f"\nTest complete. Check '{processed_output_path}'.")

