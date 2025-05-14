# -*- coding: utf-8 -*-
"""
Common utility functions used across the Excel Comparator application.
Includes Excel styling, placeholder replacement, ID generation,
and reading processed comparison data from Excel.
"""

import logging
import re
import os # Added for path manipulation
import openpyxl
from openpyxl.styles import Font, PatternFill # Ensure Font/PatternFill are imported if used
from openpyxl.utils import cell as openpyxl_cell_utils
from openpyxl.utils.exceptions import InvalidFileException # Added for specific exception
from typing import Optional, Any, Dict, Tuple, Set, List # Added List, Set
from flask import current_app # Added to access config within read_comparison_data

logger = logging.getLogger(__name__) # Use module-specific logger

# --- Constants (Copied from ui_viewer.py for use in read_comparison_data) ---
COMPARISON_SUFFIX = " Comparison"
METADATA_SHEET_NAME = "Metadata"
MAX_DN_ID_VALUE_CELL = "B1" # Cell containing Max DN ID value
MAX_AG_ID_VALUE_CELL = "B2" # Cell containing Max AG ID value


# --- Excel Utilities ---
def copy_cell_style(source_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell):
    """
    Copies font, fill, alignment, and number format style from source_cell to target_cell.

    Args:
        source_cell: The cell to copy style from.
        target_cell: The cell to copy style to.
    """
    if source_cell.has_style:
        # Copy Font properties
        target_cell.font = Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            vertAlign=source_cell.font.vertAlign,
            underline=source_cell.font.underline,
            strike=source_cell.font.strike,
            color=source_cell.font.color
        )
        # Copy Fill properties
        target_cell.fill = PatternFill(
            fill_type=source_cell.fill.fill_type,
            start_color=source_cell.fill.start_color,
            end_color=source_cell.fill.end_color
        )
        # Copy Alignment properties
        if source_cell.alignment:
            target_cell.alignment = openpyxl.styles.Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                text_rotation=source_cell.alignment.text_rotation,
                wrap_text=source_cell.alignment.wrap_text,
                shrink_to_fit=source_cell.alignment.shrink_to_fit,
                indent=source_cell.alignment.indent
            )
        # Copy Number format
        target_cell.number_format = source_cell.number_format
    else:
        # Apply default styles if source has no specific style applied
        target_cell.font = Font()
        target_cell.fill = PatternFill()
        target_cell.alignment = openpyxl.styles.Alignment()
        target_cell.number_format = 'General'


def extract_skills(expression: str) -> list[str]:
    """
    Extracts potential skill names (alphanumeric + underscore) from a skill
    expression string. Looks for patterns like 'SkillName>5'.

    Args:
        expression: The skill expression string.

    Returns:
        A list of extracted skill names.
    """
    # Regex finds words (alphanumeric or underscore) followed immediately by '>' and one or more digits
    skills = re.findall(r'\b([a-zA-Z0-9_]+)(?=>\d+)', expression)
    # Logging moved to caller if specific context is needed
    # logger.debug(f"Extracted skills {skills} from expression '{expression}'")
    return skills


# --- ID Generation Helper ---
class IdGenerator:
    """
    Generates sequential IDs within a single request context, maintaining
    separate sequences for DN (VQ) and Agent Group entities based on
    Max IDs provided during initialization.
    """
    def __init__(self, max_dn_id: int = 0, max_ag_id: int = 0):
        """
        Initializes the generator with the starting max IDs.

        Args:
            max_dn_id: The highest existing ID found for DN entities.
            max_ag_id: The highest existing ID found for Agent Group entities.
        """
        # The next ID should be one greater than the maximum found for each type
        self._next_dn_id = max_dn_id + 1
        self._next_ag_id = max_ag_id + 1
        logger.info(f"ID Generator initialized. Next DN ID: {self._next_dn_id} (based on max: {max_dn_id}), Next AG ID: {self._next_ag_id} (based on max: {max_ag_id})")

    def get_next_dn_id(self) -> int:
        """Returns the next sequential DN ID and increments the DN counter."""
        next_id = self._next_dn_id
        self._next_dn_id += 1
        logger.debug(f"Generated next DN ID: {next_id}")
        return next_id

    def get_next_ag_id(self) -> int:
        """Returns the next sequential Agent Group ID and increments the AG counter."""
        next_id = self._next_ag_id
        self._next_ag_id += 1
        logger.debug(f"Generated next AG ID: {next_id}")
        return next_id

# --- Template Placeholder Replacement ---
def replace_placeholders(template_data: Any, row_data: dict, current_row_next_id: Optional[int] = None) -> Any:
    """
    Recursively traverses a template structure (dict, list, or string)
    and replaces placeholders with values from row_data or the pre-generated ID.

    Supported Placeholders:
    - {row.ColumnName}: Replaced with the value from the corresponding column in row_data.
                        Lookup is case-insensitive for the ColumnName part.
    - {func.next_id}: Replaced with the current_row_next_id value.

    Also handles simple string concatenation like "prefix:{row.ColumnName}".

    Args:
        template_data: The template structure (can be dict, list, string, etc.).
        row_data: The dictionary containing data for the current row (keys are actual headers).
        current_row_next_id: The pre-generated sequential ID for the current row.

    Returns:
        The template structure with placeholders replaced.
    """
    # Regex to find placeholders like {row.ColumnName} or {func.FunctionName}
    placeholder_pattern = re.compile(r'{(\w+)\.([^}]+)}') # Captures type (row/func) and name

    # --- Inner replacement function ---
    def perform_replace(text: str) -> str:
        """Performs replacements on a single string."""
        if not isinstance(text, str):
            return text # Return non-strings as is

        # Function to handle each match found by the regex
        def replace_match(match):
            placeholder_type = match.group(1).lower() # 'row' or 'func'
            placeholder_name = match.group(2).strip() # 'ColumnName' or 'next_id'

            if placeholder_type == 'row':
                # --- Case-insensitive lookup ---
                found_key = None
                for key in row_data.keys():
                    if key.lower() == placeholder_name.lower():
                        found_key = key
                        break

                if found_key:
                    replacement = row_data.get(found_key, "") # Use the actual key found
                else:
                    replacement = "" # Default to empty if no matching key found
                    logger.warning(f"Placeholder {{row.{placeholder_name}}} not found in row data keys: {list(row_data.keys())}")
                # --- End Case-insensitive lookup ---
                return str(replacement) # Ensure replacement is a string

            elif placeholder_type == 'func':
                # Handle the {func.next_id} placeholder
                if placeholder_name == 'next_id':
                    if current_row_next_id is not None:
                        # Use the ID pre-generated for this specific row
                        return str(current_row_next_id)
                    else:
                        # Log a warning if the placeholder is used but no ID was provided
                        logger.warning(f"Placeholder {{func.next_id}} used but no ID provided for this row.")
                        return "{ERROR:next_id_missing}" # Indicate error in output
                else:
                    # Handle unknown function placeholders
                    logger.warning(f"Unknown function placeholder: {match.group(0)}")
                    return match.group(0) # Return the unknown placeholder itself
            else:
                 # Handle unknown placeholder types (neither row nor func)
                 logger.warning(f"Unknown placeholder type in template: {match.group(0)}")
                 return match.group(0) # Return the placeholder itself

        # Use re.sub with the handler function to replace all occurrences in the string
        return placeholder_pattern.sub(replace_match, text)
    # --- End of inner replacement function ---

    # --- Main logic for traversing template data ---
    # Process strings using the inner function
    if isinstance(template_data, str):
        return perform_replace(template_data)
    # Recursively process dictionaries
    elif isinstance(template_data, dict):
        return {
            key: replace_placeholders(value, row_data, current_row_next_id)
            for key, value in template_data.items()
        }
    # Recursively process lists
    elif isinstance(template_data, list):
        return [
            replace_placeholders(item, row_data, current_row_next_id)
            for item in template_data
        ]
    # Return numbers, booleans, None, etc., directly without modification
    else:
        return template_data


# --- Function to Read Processed Excel Data ---
def read_comparison_data(filename: str) -> bool:
    """
    Reads data from '* Comparison' sheets and 'Metadata' sheet
    of a processed Excel file into the Flask app's config cache.
    Uses headers from sheet as keys for row data dictionaries.
    Reads the maximum numeric IDs (DN and AG) from the 'Metadata' sheet.

    Args:
        filename: Path to the processed Excel file (*_processed.xlsx).

    Returns:
        True if data loading was successful (even if no comparison sheets found),
        False if a critical error occurred (e.g., file not found, invalid format).
    """
    # Note: This function modifies current_app.config directly.
    # Ensure it's called within a Flask application context.
    comparison_data = {}
    workbook = None
    comparison_sheet_names = []
    sheet_headers_cache = {} # Store headers read from each sheet
    max_dn_id_from_metadata = 0 # Initialize max IDs read from file
    max_ag_id_from_metadata = 0

    try:
        logger.info(f"Reading comparison data from: {filename}")
        if not os.path.exists(filename):
            raise FileNotFoundError(f"Processed Excel file not found at {filename}")

        workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)
        logger.info(f"Workbook loaded successfully. Sheets: {workbook.sheetnames}")

        # --- Read Max IDs from Metadata sheet ---
        if METADATA_SHEET_NAME in workbook.sheetnames:
            try:
                metadata_sheet = workbook[METADATA_SHEET_NAME]
                # Read DN Max ID value
                dn_id_val = metadata_sheet[MAX_DN_ID_VALUE_CELL].value
                if dn_id_val is not None and str(dn_id_val).isdigit():
                    max_dn_id_from_metadata = int(dn_id_val)
                    logger.info(f"Read Max DN ID from '{METADATA_SHEET_NAME}' ({MAX_DN_ID_VALUE_CELL}): {max_dn_id_from_metadata}")
                else:
                    logger.warning(f"Value in '{METADATA_SHEET_NAME}' cell {MAX_DN_ID_VALUE_CELL} is not a valid number: '{dn_id_val}'. Using 0.")

                # Read Agent Group Max ID value
                ag_id_val = metadata_sheet[MAX_AG_ID_VALUE_CELL].value
                if ag_id_val is not None and str(ag_id_val).isdigit():
                    max_ag_id_from_metadata = int(ag_id_val)
                    logger.info(f"Read Max AG ID from '{METADATA_SHEET_NAME}' ({MAX_AG_ID_VALUE_CELL}): {max_ag_id_from_metadata}")
                else:
                     logging.warning(f"Value in '{METADATA_SHEET_NAME}' cell {MAX_AG_ID_VALUE_CELL} is not a valid number: '{ag_id_val}'. Using 0.")

            except Exception as meta_e:
                logger.error(f"Error reading Max IDs from '{METADATA_SHEET_NAME}' sheet: {meta_e}. Using 0 for both.")
        else:
            logger.warning(f"'{METADATA_SHEET_NAME}' sheet not found in workbook. Max IDs will be 0.")

        # Store the read (or default 0) max IDs in app config
        current_app.config['MAX_DN_ID'] = max_dn_id_from_metadata
        current_app.config['MAX_AG_ID'] = max_ag_id_from_metadata
        # --- End Read Max IDs ---


        # Find all sheets ending with the comparison suffix
        comparison_sheet_names = sorted([s for s in workbook.sheetnames if s.endswith(COMPARISON_SUFFIX)])
        logging.info(f"Found comparison sheets: {comparison_sheet_names}")

        # If no comparison sheets found, still return True but with empty data
        if not comparison_sheet_names:
            logging.warning(f"No sheets ending with '{COMPARISON_SUFFIX}' found in {filename}.")
            current_app.config['EXCEL_DATA'] = {}
            current_app.config['COMPARISON_SHEETS'] = []
            current_app.config['EXCEL_FILENAME'] = filename
            current_app.config['SHEET_HEADERS'] = {}
            # MAX_IDs are already set from Metadata sheet check above
            return True

        # Process each comparison sheet
        for sheet_name in comparison_sheet_names:
            sheet = workbook[sheet_name]
            data: List[Dict[str, Any]] = []
            try:
                # Read the header row (expected to be row 1)
                headers = [cell.value for cell in sheet[1]]
                # Filter out None headers, ensure they are strings and stripped
                headers = [str(h).strip() for h in headers if h is not None]
                if not headers:
                    raise IndexError("No valid headers found in row 1.")
                sheet_headers_cache[sheet_name] = headers # Cache headers for this sheet
            except IndexError:
                 # Handle case where sheet might be completely empty or has no header
                 logging.warning(f"Sheet '{sheet_name}' seems empty or has no header row. Skipping.")
                 continue # Skip this sheet

            # Read data rows (starting from row 2)
            # Use the length of actual headers read to determine max columns to read
            max_cols = len(headers)
            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, max_col=max_cols), start=2):
                row_values = [cell.value for cell in row]
                # Only add row if the first cell (Key/Item) has a value
                if row_values and row_values[0] is not None and str(row_values[0]).strip() != "":
                    # Create dict using the actual headers read as keys
                    row_data = {headers[i]: row_values[i] if i < len(row_values) else None for i in range(max_cols)}
                    # Add the 'Header' key for display purposes in the template (using the first actual header)
                    row_data['Header'] = headers[0]
                    data.append(row_data)
                    # Note: Max ID calculation is now done from Metadata sheet, not here.

            comparison_data[sheet_name] = data # Store data for this sheet
            logging.info(f"Read {len(data)} valid rows from sheet '{sheet_name}'. Headers used as keys: {headers}")

        # --- Store results in app config ---
        current_app.config['EXCEL_DATA'] = comparison_data
        current_app.config['COMPARISON_SHEETS'] = comparison_sheet_names
        current_app.config['EXCEL_FILENAME'] = filename
        current_app.config['SHEET_HEADERS'] = sheet_headers_cache # Store the read headers
        # MAX_IDs were already stored earlier from Metadata sheet
        # --- End Store results ---

        return True # Indicate success

    except FileNotFoundError:
        logging.error(f"Excel file not found: {filename}")
        # Reset cache on error
        current_app.config['EXCEL_DATA'] = {}
        current_app.config['COMPARISON_SHEETS'] = []
        current_app.config['EXCEL_FILENAME'] = None
        current_app.config['MAX_DN_ID'] = 0
        current_app.config['MAX_AG_ID'] = 0
        current_app.config['SHEET_HEADERS'] = {}
        return False # Indicate failure
    except InvalidFileException:
        logging.error(f"Invalid Excel file format or corrupted file: {filename}")
        current_app.config['EXCEL_DATA'] = {}
        current_app.config['COMPARISON_SHEETS'] = []
        current_app.config['EXCEL_FILENAME'] = None
        current_app.config['MAX_DN_ID'] = 0
        current_app.config['MAX_AG_ID'] = 0
        current_app.config['SHEET_HEADERS'] = {}
        return False # Indicate failure
    except Exception as e:
        # Catch-all for other errors during file processing
        logging.error(f"Error reading Excel file '{filename}': {e}", exc_info=True)
        current_app.config['EXCEL_DATA'] = {}
        current_app.config['COMPARISON_SHEETS'] = []
        current_app.config['EXCEL_FILENAME'] = None
        current_app.config['MAX_DN_ID'] = 0
        current_app.config['MAX_AG_ID'] = 0
        current_app.config['SHEET_HEADERS'] = {}
        return False # Indicate failure
    finally:
        # Ensure workbook is closed to release resources
        if workbook:
            try:
                workbook.close()
                logging.debug("Workbook closed after reading.")
            except Exception as close_e:
                logging.warning(f"Error closing workbook: {close_e}")

