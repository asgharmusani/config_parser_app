# -*- coding: utf-8 -*-
"""
Handles reading and processing the Excel workbook.

Core Functions:
- identify_ideal_agent_column: Finds the 'Ideal Agent' column based on config.
- process_sheet: Extracts entities (VQ, Skill, VAG, Skill Expr) and metadata
                 (strike status, style) from a single sheet.
- collect_routing_entities: Orchestrates processing all relevant sheets,
                            resolves strike-through status, populates output sheets
                            in the workbook, and prepares data for comparison.
"""

import logging
import openpyxl
from openpyxl.styles import Font # Needed for type hinting if used directly
from openpyxl.utils import cell as openpyxl_cell_utils
from typing import Dict, Any, Optional, Tuple, Set

# Import utilities from utils.py
try:
    from utils import copy_cell_style, extract_skills
except ImportError:
    logging.error("Failed to import required functions from utils.py")
    # Define dummy functions or raise error if utils are critical
    def copy_cell_style(s, t): pass
    def extract_skills(e): return []
    # Consider raising an error here if utils are essential:
    # raise ImportError("Could not import utility functions. Ensure utils.py is present.")

logger = logging.getLogger(__name__) # Use module-specific logger

# --- Constants from Config (Passed as Arguments) ---
# These were previously global, now passed via config dict or arguments

# --- Helper Functions ---

def identify_ideal_agent_column(sheet: openpyxl.worksheet.worksheet.Worksheet, config: Dict[str, Any]) -> Optional[int]:
    """
    Identifies the column index for the 'Ideal Agent' column based on config.
    Searches header row 1 (specifically Columns C & D) first, then checks a fallback cell.

    Args:
        sheet: The openpyxl worksheet object.
        config: The application configuration dictionary containing keys like
                'ideal_agent_header_text' and 'ideal_agent_fallback_cell'.

    Returns:
        The column index (1-based) if found, otherwise None.
    """
    header_text = config.get('ideal_agent_header_text', 'Ideal Agent') # Use default if key missing
    fallback_cell_coord = config.get('ideal_agent_fallback_cell', 'C2') # Use default if key missing
    logger.debug(f"Identifying '{header_text}' column in sheet: {sheet.title}")

    # 1. Check headers in row 1, specifically columns C and D
    for col_idx in [3, 4]:  # Column C=3, Column D=4
        # Check if sheet has enough columns before accessing cell
        if col_idx <= sheet.max_column:
            cell_value = sheet.cell(row=1, column=col_idx).value
            # Check if cell has a value and contains the header text
            if cell_value and header_text in str(cell_value):
                logger.debug(f"Found '{header_text}' in header row 1 at column {col_idx}")
                return col_idx
        else:
            logger.debug(f"Sheet '{sheet.title}' has only {sheet.max_column} columns, cannot check column {col_idx} in header.")

    # 2. Check specific fallback cell from config if not found in header cols C/D
    try:
        # Convert cell coordinate like "C2" to row/col index
        col_str, row_str = openpyxl_cell_utils.coordinate_to_tuple(fallback_cell_coord)
        fallback_col_idx = openpyxl_cell_utils.column_index_from_string(col_str)
        fallback_row_idx = int(row_str)

        # Check if fallback cell is within sheet bounds
        if fallback_row_idx <= sheet.max_row and fallback_col_idx <= sheet.max_column:
            cell_value_fallback = sheet.cell(row=fallback_row_idx, column=fallback_col_idx).value
            if cell_value_fallback and header_text in str(cell_value_fallback):
                logger.debug(f"Found '{header_text}' at fallback cell {fallback_cell_coord} (Col {fallback_col_idx})")
                return fallback_col_idx
        else:
            logger.warning(f"Fallback cell '{fallback_cell_coord}' is outside the bounds of sheet '{sheet.title}'.")

    except Exception as e:
         # Catch potential errors during coordinate conversion or cell access
         logger.warning(f"Could not parse or check fallback cell '{fallback_cell_coord}': {e}")

    # 3. If not found by either method
    logger.warning(f"'{header_text}' column not found using header search (Cols C/D) or fallback cell '{fallback_cell_coord}' in sheet: {sheet.title}.")
    return None


def process_sheet(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    intermediate_data: Dict[str, Dict[str, Dict[str, Any]]],
    config: Dict[str, Any],
    excluded_sheets: Set[str]
):
    """
    Processes a single sheet to extract routing entities and their strike status,
    updating the intermediate_data dictionary according to the strike-preference rule.
    Stores separate expr/ideal for skill_exprs.

    Args:
        sheet: The openpyxl worksheet object to process.
        intermediate_data: The dictionary holding collected data across sheets.
                           Structure: {"vqs": {name: details}, "skills": {name: details}, ...}
                           Details: {"strike": bool, "style_cell": Cell, "expr": str?, "ideal": str?}
        config: The loaded application configuration dictionary.
        excluded_sheets: A set of sheet names to skip during processing.
    """
    if sheet.title in excluded_sheets:
        logger.debug(f"Skipping sheet: {sheet.title} (excluded name)")
        return

    logging.info(f"Processing sheet: {sheet.title} (Max Row: {sheet.max_row}, Max Col: {sheet.max_column})")
    ideal_agent_col_idx = identify_ideal_agent_column(sheet, config)
    vag_sheet_name = config.get('vag_extraction_sheet', 'Default Targeting- Group') # Get from config or use default

    # Helper for VQ, VAG, Skills (non-skill-expressions)
    def update_intermediate_generic(data_dict: Dict, key: str, current_strike: bool, cell_obj: openpyxl.cell.Cell):
        """Updates the dict for generic items, preferring non-struck entries."""
        if not key: # Skip empty keys
             return
        # Check if item exists
        if key not in data_dict:
            # Add new item
            data_dict[key] = {"strike": current_strike, "style_cell": cell_obj}
            logging.debug(f"Sheet '{sheet.title}' Row {cell_obj.row}: Added new item '{key}' with strike={current_strike}")
        # If item exists, update only if changing strike from True to False
        elif data_dict[key]["strike"] and not current_strike:
            data_dict[key]["strike"] = False
            data_dict[key]["style_cell"] = cell_obj # Use style from non-struck cell
            logging.debug(f"Sheet '{sheet.title}' Row {cell_obj.row}: Updated item '{key}' to strike=False")
        # else: if existing is False, or both are True, no change needed

    processed_cells = 0
    # Iterate through all cells within the sheet's used range
    for row_idx in range(1, sheet.max_row + 1):
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)

            # Skip empty cells
            if cell.value is None or str(cell.value).strip() == "":
                continue

            value_str = str(cell.value).strip()
            # Determine strike status from cell font
            strike = bool(cell.font and cell.font.strike)
            processed_cells += 1

            # --- Skill Expression Processing (Specific Handling) ---
            if ">" in value_str:
                # Normalize expression part
                formatted_expr = value_str.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                ideal_agent_value = ""
                # Get corresponding Ideal Agent value if column was found
                if ideal_agent_col_idx:
                    # Check bounds before accessing cell
                    if ideal_agent_col_idx <= sheet.max_column:
                        ideal_cell = sheet.cell(row=row_idx, column=ideal_agent_col_idx)
                        ideal_agent_value = str(ideal_cell.value).strip() if ideal_cell.value else ""
                        # Normalize ideal part as well
                        ideal_agent_value = ideal_agent_value.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                    else:
                        # Log if ideal agent column index is out of bounds for this sheet
                        logging.warning(f"Ideal agent column {ideal_agent_col_idx} out of bounds for sheet '{sheet.title}' (max col: {sheet.max_column}) at row {row_idx}")

                # Create the concatenated key used for comparison and storage
                concatenated_value = f"{formatted_expr} {ideal_agent_value}".strip()
                if not concatenated_value:
                    # Skip if the key becomes empty after processing (e.g., only operators left)
                    continue

                # Update intermediate_data for skill_exprs directly (handles strike rule)
                skill_expr_dict = intermediate_data["skill_exprs"]
                if concatenated_value not in skill_expr_dict:
                    # Add new skill expression entry with all details
                    skill_expr_dict[concatenated_value] = {
                        "strike": strike,
                        "style_cell": cell,
                        "expr": formatted_expr,  # Store separate expression
                        "ideal": ideal_agent_value # Store separate ideal expression
                    }
                    logging.debug(f"Sheet '{sheet.title}' Row {cell.row}: Added new skill_expr '{concatenated_value}' with strike={strike}")
                elif skill_expr_dict[concatenated_value]["strike"] and not strike:
                    # Update existing entry if changing strike from True to False
                    skill_expr_dict[concatenated_value]["strike"] = False
                    skill_expr_dict[concatenated_value]["style_cell"] = cell # Use non-struck cell style
                    # Keep original expr/ideal values associated with the key
                    logging.debug(f"Sheet '{sheet.title}' Row {cell.row}: Updated skill_expr '{concatenated_value}' to strike=False")

                # Also extract individual skills from the expression part for the "Skills" sheet
                individual_skills = extract_skills(formatted_expr)
                for skill in individual_skills:
                     update_intermediate_generic(intermediate_data["skills"], skill, strike, cell)

            # --- VQ Processing ---
            # Check if it looks like a VQ name and NOT a skill expression
            elif (value_str.startswith("VQ_") or value_str.startswith("vq_") or ("VQ" in value_str)) and ">" not in value_str:
                 normalized_vq = value_str.replace(" ", "").replace('\u00A0', '')
                 update_intermediate_generic(intermediate_data["vqs"], normalized_vq, strike, cell)

            # --- VAG Processing (Only from specific sheet defined in config) ---
            elif "VAG_" in value_str and sheet.title == vag_sheet_name:
                 normalized_vag = value_str.replace(" ", "").replace('\u00A0', '')
                 update_intermediate_generic(intermediate_data["vags"], normalized_vag, strike, cell)

            # --- Simple Skill Check (Optional - Add if simple skills appear outside expressions) ---
            # Example: Check if it looks like a skill name and not other patterns
            # elif re.match(r'^[a-zA-Z0-9_]+$', value_str) and not any(x in value_str for x in ['>', 'VQ_', 'VAG_']):
            #      update_intermediate_generic(intermediate_data["skills"], value_str, strike, cell)

    logging.debug(f"Processed {processed_cells} non-empty cells in sheet '{sheet.title}'.")


def collect_routing_entities(
    workbook: openpyxl.workbook.Workbook,
    config: Dict[str, Any],
    metadata_sheet_name: str # Pass metadata sheet name constant
) -> Tuple[Dict[str, Set[str]], Dict[str, Dict[str, Dict[str, Any]]]]:
    """
    Processes workbook sheets using intermediate data structure, populates final
    output sheets based on resolved strike status (incl. separate expr/ideal for Skill Expr),
    and returns sheet data (non-struck only) for comparison AND the full intermediate data.
    Modifies the workbook object in place by deleting/adding sheets.

    Args:
        workbook: The openpyxl Workbook object (loaded from the working copy).
        config: The loaded application configuration dictionary.
        metadata_sheet_name: The name defined for the metadata sheet.

    Returns:
        Tuple containing:
        - sheet_data_for_comparison: Dict[str, Set[str]] - Sets of non-struck keys for comparison.
        - intermediate_data: Dict[str, Dict[str, Dict[str, Any]]] - Full collected data with details.
    """
    logging.info("Starting collection and processing of routing entities from workbook sheets.")

    # --- Intermediate Data Structure ---
    # Stores details for each unique item found across all sheets
    intermediate_data = {
        "vqs": {}, "skills": {}, "skill_exprs": {}, "vags": {}
    } # Format: {"Type": {"Name/Key": {"strike": bool, "style_cell": Cell, "expr"?: str, "ideal"?: str}}}

    # Define structure and headers for the output sheets
    output_sheet_specs = {
        "Skill Expr": {"key": "skill_exprs", "headers": ["Expression", "Ideal Expression", "Concatenated Key", "HasStrikeThrough"]},
        "VQ": {"key": "vqs", "headers": ["VQ Name", "HasStrikeThrough"]},
        "VAG": {"key": "vags", "headers": ["VAG Name", "HasStrikeThrough"]},
        "Skills": {"key": "skills", "headers": ["Skill", "HasStrikeThrough"]}
    }
    # Define names for comparison sheets
    comparison_prefixes = {"Skill Expr": "Skill_exprs", "VQ": "Vqs", "VAG": "Vags", "Skills": "Skills"}
    # List all sheets to remove (output, comparison, metadata)
    sheets_to_remove = list(output_sheet_specs.keys()) + \
                       [f"{comparison_prefixes[t]} Comparison" for t in output_sheet_specs.keys()] + \
                       [metadata_sheet_name]

    # --- Phase 0: Remove old generated sheets ---
    logging.info("Removing previously generated output/comparison/metadata sheets...")
    for sheet_name in sheets_to_remove:
         if sheet_name in workbook.sheetnames:
             try:
                 del workbook[sheet_name]
                 logging.debug(f"Removed existing sheet: {sheet_name}")
             except Exception as e:
                 logging.warning(f"Could not remove sheet '{sheet_name}': {e}")


    # --- Phase 1: Process all sheets and populate intermediate_data ---
    # Define the set of sheets to exclude from processing here
    excluded_sheets_for_processing = set(sheets_to_remove) # Start with sheets we'll create/remove
    # Add any other specific sheets to exclude if necessary
    # excluded_sheets_for_processing.add("Instructions")

    for sheet in workbook.worksheets:
        process_sheet(sheet, intermediate_data, config, excluded_sheets_for_processing)
    logging.info("Finished processing all sheets, resolved strikethrough status.")


    # --- Phase 2: Create/Populate Output Sheets from intermediate_data ---
    logging.info("Populating dedicated output sheets (VQ, Skills, Skill Expr, VAG)...")
    for title, spec in output_sheet_specs.items():
        sheet = workbook.create_sheet(title=title)
        # Write Headers and make bold
        for col_idx, header in enumerate(spec["headers"], start=1):
             sheet.cell(row=1, column=col_idx, value=header).font = Font(bold=True)

        row_num = 2 # Start writing data from row 2
        data_items = intermediate_data.get(spec["key"], {}) # Get data for this type

        # Sort items alphabetically by key for consistent output order
        for item_key, data in sorted(data_items.items()):
             if title == "Skill Expr":
                 # Populate the 4 columns for Skill Expr sheet
                 sheet.cell(row=row_num, column=1, value=data.get("expr", "")) # Expression
                 sheet.cell(row=row_num, column=2, value=data.get("ideal", "")) # Ideal Expression
                 cell_key = sheet.cell(row=row_num, column=3, value=item_key) # Concatenated Key
                 sheet.cell(row=row_num, column=4, value=str(data["strike"])) # HasStrikeThrough
                 # Copy style from the representative cell found during processing to the Key cell
                 copy_cell_style(data["style_cell"], cell_key)
             else:
                 # Standard 2-column population for VQ, VAG, Skills
                 cell_a = sheet.cell(row=row_num, column=1, value=item_key) # Name
                 sheet.cell(row=row_num, column=2, value=str(data["strike"])) # HasStrikeThrough
                 # Copy style from the representative cell
                 copy_cell_style(data["style_cell"], cell_a)
             row_num += 1
        logging.debug(f"Populated '{title}' sheet with {row_num - 2} items.")
    logging.info("Finished populating output sheets.")

    # --- Phase 3: Prepare data for comparison (only non-struck items' keys) ---
    # This data structure is used by write_comparison_sheet
    sheet_data_for_comparison = {
        key: {name for name, data in items.items() if not data["strike"]} # Set of keys where strike is False
        for key, items in intermediate_data.items()
    }

    logging.info("Prepared final sheet data for comparison (non-struck items only).")
    logging.debug(f"Comparison Data Summary: VQs={len(sheet_data_for_comparison['vqs'])}, Skills={len(sheet_data_for_comparison['skills'])}, SkillExprs={len(sheet_data_for_comparison['skill_exprs'])}, VAGs={len(sheet_data_for_comparison['vags'])}")

    # Return both the comparison keys and the full intermediate data with details
    return sheet_data_for_comparison, intermediate_data

