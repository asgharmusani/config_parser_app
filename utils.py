# -*- coding: utf-8 -*-
"""
Common utility functions used across the Excel Comparator application.
Includes Excel styling, placeholder replacement, ID generation,
reading processed comparison data, and identifier matching.
"""

import logging
import re
import os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import cell as openpyxl_cell_utils
from openpyxl.utils.exceptions import InvalidFileException
from typing import Optional, Any, Dict, Tuple, Set, List
from flask import current_app

logger = logging.getLogger(__name__)

# --- Constants ---
COMPARISON_SUFFIX = " Comparison"
METADATA_SHEET_NAME = "Metadata"
MAX_DN_ID_VALUE_CELL = "B1"
MAX_AG_ID_VALUE_CELL = "B2"

# --- Excel Utilities ---
def copy_cell_style(source_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell):
    """Copies font, fill, alignment, and number format style."""
    if source_cell.has_style:
        target_cell.font = Font(name=source_cell.font.name, size=source_cell.font.size, bold=source_cell.font.bold, italic=source_cell.font.italic, vertAlign=source_cell.font.vertAlign, underline=source_cell.font.underline, strike=source_cell.font.strike, color=source_cell.font.color)
        target_cell.fill = PatternFill(fill_type=source_cell.fill.fill_type, start_color=source_cell.fill.start_color, end_color=source_cell.fill.end_color)
        if source_cell.alignment:
            target_cell.alignment = openpyxl.styles.Alignment(horizontal=source_cell.alignment.horizontal, vertical=source_cell.alignment.vertical, text_rotation=source_cell.alignment.text_rotation, wrap_text=source_cell.alignment.wrap_text, shrink_to_fit=source_cell.alignment.shrink_to_fit, indent=source_cell.alignment.indent)
        target_cell.number_format = source_cell.number_format
    else:
        target_cell.font = Font(); target_cell.fill = PatternFill(); target_cell.alignment = openpyxl.styles.Alignment(); target_cell.number_format = 'General'

def extract_skills(expression: str) -> list[str]:
    """Extracts potential skill names from a skill expression string."""
    skills = re.findall(r'\b([a-zA-Z0-9_]+)(?=>\d+)', expression)
    return skills

# --- MODIFICATION START: Moved and adapted _match_identifier ---
def match_identifier_logic(value_to_check_str: str, identifier_rule: Dict[str, Any]) -> bool:
    """
    Checks if a string value matches a given identifier rule.
    This is a utility version of ExcelRuleEngine._match_identifier.
    Assumes identifier_rule contains pre-processed keys if coming from ExcelRuleEngine,
    or raw keys if called directly with a rule snippet.

    Args:
        value_to_check_str: The string value to check.
        identifier_rule: The identifier part of an entity rule.
                         May contain pre-processed keys like '_type_processed',
                         '_value_to_compare_processed', '_compiled_regex_processed',
                         or raw 'type', 'value', 'caseSensitive'.
    Returns:
        True if the value matches the rule, False otherwise.
    """
    id_type = identifier_rule.get('_type_processed', str(identifier_rule.get("type", "")).lower())
    case_sensitive = identifier_rule.get('_case_sensitive_processed', identifier_rule.get("caseSensitive", False))

    # Get the value to compare against from the rule
    # If pre-processed, it's under _value_to_compare_processed or _value_original (for regex)
    # Otherwise, get it from "value"
    if id_type == "regex":
        value_from_rule = identifier_rule.get('_value_original', identifier_rule.get("value", ""))
        compiled_regex = identifier_rule.get('_compiled_regex_processed')
        if not compiled_regex and value_from_rule: # Compile if not already
            try:
                compiled_regex = re.compile(value_from_rule)
            except re.error as e:
                logger.warning(f"Invalid regex '{value_from_rule}' in identifier rule: {e}")
                return False
    else:
        value_from_rule = identifier_rule.get('_value_to_compare_processed', identifier_rule.get("value", ""))
        if not case_sensitive: # If not regex and not case sensitive, rule value should be lower
            value_from_rule = value_from_rule.lower()


    if not value_from_rule and id_type != "regex": # Regex can be valid without _value_to_compare_processed if compiled
         if not compiled_regex and id_type == "regex":
            logger.debug(f"Identifier rule missing value/compiled regex: {identifier_rule}")
            return False
    elif not value_from_rule and id_type != "regex":
        logger.debug(f"Identifier rule missing value: {identifier_rule}")
        return False


    # Prepare cell value for comparison based on case sensitivity for non-regex types
    val_to_check_prepared = value_to_check_str if (id_type == "regex" or case_sensitive) else value_to_check_str.lower()

    if id_type == "startswith":
        return val_to_check_prepared.startswith(value_from_rule)
    elif id_type == "contains":
        return value_from_rule in val_to_check_prepared
    elif id_type == "exactmatch":
        return val_to_check_prepared == value_from_rule
    elif id_type == "regex":
        if compiled_regex:
            return bool(compiled_regex.search(value_to_check_str)) # Regex uses original case string
        return False # Should have been caught by compiled_regex check
    else:
        logger.warning(f"Unknown identifier type: '{id_type}' in rule: {identifier_rule}")
        return False
# --- MODIFICATION END ---


# --- ID Generation Helper ---
# ... (Keep existing IdGenerator class) ...
class IdGenerator:
    def __init__(self, max_dn_id: int = 0, max_ag_id: int = 0):
        self._next_dn_id = max_dn_id + 1
        self._next_ag_id = max_ag_id + 1
        logger.info(f"ID Generator initialized. Next DN ID: {self._next_dn_id} (max: {max_dn_id}), Next AG ID: {self._next_ag_id} (max: {max_ag_id})")
    def get_next_dn_id(self) -> int: next_id = self._next_dn_id; self._next_dn_id += 1; logger.debug(f"Gen next DN ID: {next_id}"); return next_id
    def get_next_ag_id(self) -> int: next_id = self._next_ag_id; self._next_ag_id += 1; logger.debug(f"Gen next AG ID: {next_id}"); return next_id

# --- Template Placeholder Replacement ---
# ... (Keep existing replace_placeholders function) ...
def replace_placeholders(template_data: Any, row_data: dict, current_row_next_id: Optional[int] = None) -> Any:
    placeholder_pattern = re.compile(r'{(\w+)\.([^}]+)}')
    def perform_replace(text: str) -> str:
        if not isinstance(text, str): return text
        def replace_match(match):
            placeholder_type = match.group(1).lower(); placeholder_name = match.group(2).strip()
            if placeholder_type == 'row':
                found_key = None
                for key in row_data.keys():
                    if key.lower() == placeholder_name.lower(): found_key = key; break
                if found_key: replacement = row_data.get(found_key, "")
                else: replacement = ""; logger.warning(f"Placeholder {{row.{placeholder_name}}} not found in row data keys: {list(row_data.keys())}")
                return str(replacement)
            elif placeholder_type == 'func':
                if placeholder_name == 'next_id':
                    if current_row_next_id is not None: return str(current_row_next_id)
                    else: logger.warning(f"Placeholder {{func.next_id}} used but no ID provided."); return "{ERROR:next_id_missing}"
                else: logger.warning(f"Unknown function placeholder: {match.group(0)}"); return match.group(0)
            else: logger.warning(f"Unknown placeholder type: {match.group(0)}"); return match.group(0)
        return placeholder_pattern.sub(replace_match, text)
    if isinstance(template_data, str): return perform_replace(template_data)
    elif isinstance(template_data, dict): return { key: replace_placeholders(value, row_data, current_row_next_id) for key, value in template_data.items() }
    elif isinstance(template_data, list): return [ replace_placeholders(item, row_data, current_row_next_id) for item in template_data ]
    else: return template_data

# --- Function to Read Processed Excel Data ---
# ... (Keep existing read_comparison_data function) ...
def read_comparison_data(filename: str) -> bool:
    """ Reads data from '* Comparison' sheets and 'Metadata' sheet into app config cache. """
    comparison_data = {}; workbook = None; comparison_sheet_names = []; sheet_headers_cache = {}
    max_dn_id_from_metadata = 0; max_ag_id_from_metadata = 0
    try:
        logger.info(f"Reading comparison data from: {filename}")
        if not os.path.exists(filename): raise FileNotFoundError(f"Processed Excel file not found at {filename}")
        workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)
        logger.info(f"Workbook loaded successfully. Sheets: {workbook.sheetnames}")
        if METADATA_SHEET_NAME in workbook.sheetnames:
            try:
                metadata_sheet = workbook[METADATA_SHEET_NAME]
                dn_id_val = metadata_sheet[MAX_DN_ID_VALUE_CELL].value
                if dn_id_val is not None and str(dn_id_val).isdigit(): max_dn_id_from_metadata = int(dn_id_val); logger.info(f"Read Max DN ID from '{METADATA_SHEET_NAME}' ({MAX_DN_ID_VALUE_CELL}): {max_dn_id_from_metadata}")
                else: logger.warning(f"Value in '{METADATA_SHEET_NAME}' cell {MAX_DN_ID_VALUE_CELL} not valid number: '{dn_id_val}'. Using 0.")
                ag_id_val = metadata_sheet[MAX_AG_ID_VALUE_CELL].value
                if ag_id_val is not None and str(ag_id_val).isdigit(): max_ag_id_from_metadata = int(ag_id_val); logger.info(f"Read Max AG ID from '{METADATA_SHEET_NAME}' ({MAX_AG_ID_VALUE_CELL}): {max_ag_id_from_metadata}")
                else: logger.warning(f"Value in '{METADATA_SHEET_NAME}' cell {MAX_AG_ID_VALUE_CELL} not valid number: '{ag_id_val}'. Using 0.")
            except Exception as meta_e: logger.error(f"Error reading Max IDs from '{METADATA_SHEET_NAME}': {meta_e}. Using 0 for both.")
        else: logger.warning(f"'{METADATA_SHEET_NAME}' sheet not found. Max IDs will be 0.")
        current_app.config['MAX_DN_ID'] = max_dn_id_from_metadata
        current_app.config['MAX_AG_ID'] = max_ag_id_from_metadata
        comparison_sheet_names = sorted([s for s in workbook.sheetnames if s.endswith(COMPARISON_SUFFIX)])
        logger.info(f"Found comparison sheets: {comparison_sheet_names}")
        if not comparison_sheet_names:
            logger.warning(f"No sheets ending with '{COMPARISON_SUFFIX}' found in {filename}.")
            current_app.config['EXCEL_DATA'] = {}; current_app.config['COMPARISON_SHEETS'] = []; current_app.config['EXCEL_FILENAME'] = filename; current_app.config['SHEET_HEADERS'] = {}; return True
        for sheet_name in comparison_sheet_names:
            sheet = workbook[sheet_name]; data: List[Dict[str, Any]] = []
            try:
                headers = [str(h).strip() for h in sheet[1] if h.value is not None] # Read headers from row 1
                if not headers: raise IndexError("No valid headers found in row 1.")
                sheet_headers_cache[sheet_name] = headers
            except IndexError: logger.warning(f"Sheet '{sheet_name}' empty/no header. Skipping."); continue
            max_cols = len(headers)
            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, max_col=max_cols), start=2):
                row_values = [cell.value for cell in row]
                if row_values and row_values[0] is not None and str(row_values[0]).strip() != "":
                    row_data = {headers[i]: row_values[i] if i < len(row_values) else None for i in range(max_cols)}; row_data['Header'] = headers[0]; data.append(row_data)
            comparison_data[sheet_name] = data; logger.info(f"Read {len(data)} valid rows from sheet '{sheet_name}'. Headers: {headers}")
        current_app.config['EXCEL_DATA'] = comparison_data; current_app.config['COMPARISON_SHEETS'] = comparison_sheet_names; current_app.config['EXCEL_FILENAME'] = filename; current_app.config['SHEET_HEADERS'] = sheet_headers_cache; return True
    except FileNotFoundError: logger.error(f"Excel file not found: {filename}"); current_app.config.update({'EXCEL_DATA': {}, 'COMPARISON_SHEETS': [], 'EXCEL_FILENAME': None, 'MAX_DN_ID': 0, 'MAX_AG_ID': 0, 'SHEET_HEADERS': {}}); return False
    except InvalidFileException: logger.error(f"Invalid Excel file format: {filename}"); current_app.config.update({'EXCEL_DATA': {}, 'COMPARISON_SHEETS': [], 'EXCEL_FILENAME': None, 'MAX_DN_ID': 0, 'MAX_AG_ID': 0, 'SHEET_HEADERS': {}}); return False
    except Exception as e: logger.error(f"Error reading Excel file '{filename}': {e}", exc_info=True); current_app.config.update({'EXCEL_DATA': {}, 'COMPARISON_SHEETS': [], 'EXCEL_FILENAME': None, 'MAX_DN_ID': 0, 'MAX_AG_ID': 0, 'SHEET_HEADERS': {}}); return False
    finally:
        if workbook:
            try: workbook.close(); logger.debug("Workbook closed after reading.")
            except Exception as close_e: logger.warning(f"Error closing workbook: {close_e}")

