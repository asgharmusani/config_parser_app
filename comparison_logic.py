# -*- coding: utf-8 -*-
"""
Handles the comparison logic between processed Excel data and fetched API data,
and writes the results to dedicated comparison sheets in the workbook.
"""

import logging
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import cell as openpyxl_cell_utils
from typing import Dict, Any, Set

logger = logging.getLogger(__name__) # Use module-specific logger

# --- Comparison and Reporting ---
def write_comparison_sheets(
    workbook: openpyxl.workbook.Workbook,
    sheet_data_for_comparison: Dict[str, Set[str]],
    api_data: Dict[str, Dict[str, Any]], # API data structure varies by key
    intermediate_data: Dict[str, Dict[str, Dict[str, Any]]] # Full sheet data details
):
    """
    Compares sheet data (non-struck only) with API data and writes results
    to dedicated comparison sheets in the workbook. Adds Expression/Ideal
    columns to the Skill_exprs Comparison sheet.

    Args:
        workbook: The openpyxl Workbook object to write results into.
        sheet_data_for_comparison: Dict containing sets of non-struck item keys found in sheets.
                                   Format: {"vqs": {set_of_names}, "skills": {set}, ...}
        api_data: Dict containing items fetched from the API.
                  Format: {"vqs": {name: id}, "skills": {name: id}, ...,
                           "skill_exprs": {concat_key: {"id":.., "expr":.., "ideal":..}}}
        intermediate_data: Dict containing detailed items found in sheets, including style info
                           and separate expr/ideal for skill expressions. Used to get details
                           for items marked as 'New in Sheet'.
                           Format: {"vqs": {name: details}, "skills": {name: details}, ...}
    """
    logging.info("Starting comparison and writing results to comparison sheets.")

    # Basic checks for empty data
    if not api_data:
        logging.warning("API data is empty or None, skipping comparison writing.")
        # Optionally create empty comparison sheets? For now, just return.
        return
    if not sheet_data_for_comparison:
        logging.warning("Sheet data for comparison is empty or None, skipping comparison writing.")
        return

    # Map internal keys to sheet title prefixes used in comparison sheet names
    # These prefixes are used for sheet naming and the first column header in standard sheets.
    comparison_keys_map = {
        "vqs": "Vqs",
        "skills": "Skills",
        "skill_exprs": "Skill_exprs",
        "vags": "Vags"
    }

    # Iterate through each entity type (vqs, skills, etc.)
    for key, sheet_title_prefix in comparison_keys_map.items():
        logging.info(f"Generating comparison sheet for: {key}")
        comparison_sheet_title = f"{sheet_title_prefix} Comparison"

        # Ensure sheet doesn't already exist (should have been removed earlier, but double-check)
        if comparison_sheet_title in workbook.sheetnames:
            try:
                del workbook[comparison_sheet_title]
                logging.debug(f"Removed pre-existing sheet: {comparison_sheet_title}")
            except Exception as e:
                 logging.warning(f"Could not remove existing sheet '{comparison_sheet_title}': {e}")
                 # Continue anyway, openpyxl might handle overwriting implicitly or error out later

        # Create the new comparison sheet
        sheet = workbook.create_sheet(title=comparison_sheet_title)

        # --- Prepare data specific to this entity type ---
        sheet_items_non_struck = sheet_data_for_comparison.get(key, set()) # Keys of non-struck items from sheet
        api_items_dict = api_data.get(key, {}) # Dict of API items {key: details_or_id}
        api_items_keys = set(api_items_dict.keys()) # Keys of items found in API

        # Calculate differences based on the KEYS
        # Items present (non-struck) in sheet but not in API
        new_in_sheet = sheet_items_non_struck - api_items_keys
        # Items present in API but not present (non-struck) in sheet
        missing_from_sheet_non_struck = api_items_keys - sheet_items_non_struck

        row_num = 2 # Start writing data rows from row 2

        # --- Set Headers and Column Widths based on entity type ---
        if key == "skill_exprs":
            # Define headers for the 5-column Skill Exprs comparison sheet
            headers = ["Concatenated Key", "Expression", "Ideal Expression", "ID (from API)", "Status"]
            # Define approximate column widths for better viewing
            col_widths = [45, 45, 35, 20, 35]
        else:
            # Define headers for the standard 3-column comparison sheets
            # Use the sheet_title_prefix (e.g., "Vqs", "Skills") as the first column header
            headers = [sheet_title_prefix, "ID (from API)", "Status"]
            col_widths = [45, 20, 35]

        # Write headers to the sheet and apply formatting
        for col_idx, header in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True) # Make headers bold
            # Set column width for better readability
            try:
                column_letter = openpyxl_cell_utils.get_column_letter(col_idx)
                sheet.column_dimensions[column_letter].width = col_widths[col_idx-1]
            except IndexError: # Safety check for col_widths definition
                 pass # Ignore error if width definition is wrong


        # --- Write Data Rows ---

        # Write items that are "New in Sheet"
        if new_in_sheet:
            logging.debug(f"'{key}' - Found {len(new_in_sheet)} items New in Sheet (Non-Struck).")
            # Sort items alphabetically by key for consistent report order
            for item_key in sorted(list(new_in_sheet)):
                if key == "skill_exprs":
                    # Lookup details from intermediate_data (which originates from sheet processing)
                    item_details = intermediate_data.get('skill_exprs', {}).get(item_key, {})
                    sheet.cell(row=row_num, column=1, value=item_key) # Concatenated Key
                    sheet.cell(row=row_num, column=2, value=item_details.get('expr', '')) # Expression from sheet
                    sheet.cell(row=row_num, column=3, value=item_details.get('ideal', '')) # Ideal Expression from sheet
                    sheet.cell(row=row_num, column=4, value="N/A") # ID (Not applicable as it's not from API)
                    sheet.cell(row=row_num, column=5, value="New in Sheet (Non-Struck)") # Status
                else:
                    # Standard 3-column layout for VQ, Skill, VAG
                    sheet.cell(row=row_num, column=1, value=item_key) # Item Name
                    sheet.cell(row=row_num, column=2, value="N/A") # ID
                    sheet.cell(row=row_num, column=3, value="New in Sheet (Non-Struck)") # Status
                row_num += 1
        else:
             # Log if no items were found only in the sheet
             logging.debug(f"'{key}' - No items found only in the sheet (non-struck).")

        # Write items that are "Missing from Sheet" (or only struck out)
        if missing_from_sheet_non_struck:
             logging.debug(f"'{key}' - Found {len(missing_from_sheet_non_struck)} items Missing from Sheet (or only Struck Out).")
             # Sort items alphabetically by key for consistent report order
             for item_key in sorted(list(missing_from_sheet_non_struck)):
                if key == "skill_exprs":
                    # Lookup details from api_data (which originates from API)
                    # api_items_dict is api_data['skill_exprs'] here
                    api_details = api_items_dict.get(item_key, {}) # api_details is {'id': ..., 'expr': ..., 'ideal': ...}
                    sheet.cell(row=row_num, column=1, value=item_key) # Concatenated Key
                    sheet.cell(row=row_num, column=2, value=api_details.get('expr', '')) # Expression from API
                    sheet.cell(row=row_num, column=3, value=api_details.get('ideal', '')) # Ideal Expression from API
                    sheet.cell(row=row_num, column=4, value=api_details.get('id', 'ID Not Found')) # ID from API
                    sheet.cell(row=row_num, column=5, value="Missing in Sheet (or only Struck Out)") # Status
                else:
                    # Standard 3-column layout for VQ, Skill, VAG
                    # api_items_dict is api_data['vqs'], etc. here, value is just the ID string
                    api_id = api_items_dict.get(item_key, "ID Not Found") # api_id is just the string ID here
                    sheet.cell(row=row_num, column=1, value=item_key) # Item Name
                    sheet.cell(row=row_num, column=2, value=api_id) # ID from API
                    sheet.cell(row=row_num, column=3, value="Missing in Sheet (or only Struck Out)") # Status
                row_num += 1
        else:
            # Log if no items were found only in the API data
            logging.debug(f"'{key}' - No items found only in the API (when compared to non-struck sheet items).")

        logging.info(f"Finished comparison sheet for: {key}")

