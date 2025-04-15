# -*- coding: utf-8 -*-
"""
Processes an Excel workbook (.xlsx) to extract routing entities (VQs, Skills,
VAGs, Skill Expressions), compares them against data fetched from Genesys
configuration APIs, and reports differences in new sheets within the workbook.

Handles strikethrough formatting: if an item appears both with and without
strikethrough in the source sheet, it's treated as effectively 'present'
(not struck out) for comparison purposes.

The 'Skill Expr' output sheet includes separate 'Expression' and
'Ideal Expression' columns for reference, while comparison uses the
'Concatenated Key'.

Configuration (file paths, API URLs, sheet layout details) is loaded from
a 'config.ini' file expected in the same directory.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import cell as openpyxl_cell_utils
import re
import requests
import logging
import shutil
import os
import configparser
from typing import Dict, Any, Optional, Tuple, Set, List

# --- Constants ---
CONFIG_FILE = 'config.ini'
LOG_FILE = 'log.txt'

# --- Logging Setup ---
logging.basicConfig(
    level=logging.DEBUG, # Set root logger level
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w'), # Overwrite log file each run
        logging.StreamHandler() # Log to console (stderr by default)
    ]
)
# Adjust console handler level if needed (e.g., only show INFO+)
for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.StreamHandler):
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(levelname)s: %(message)s') # Simpler format for console
        handler.setFormatter(formatter)


# --- Configuration Loading ---
def load_configuration(config_path: str) -> Dict[str, Any]:
    """
    Loads settings from the specified INI configuration file.

    Args:
        config_path: Path to the configuration file (e.g., 'config.ini').

    Returns:
        A dictionary containing the loaded configuration settings.

    Raises:
        FileNotFoundError: If the configuration file does not exist.
        ValueError: If the configuration file has missing sections/options
                    or other parsing errors.
    """
    logging.info(f"Loading configuration from: {config_path}")
    if not os.path.exists(config_path):
        logging.error(f"Configuration file '{config_path}' not found.")
        raise FileNotFoundError(f"Configuration file '{config_path}' not found.")

    config = configparser.ConfigParser(interpolation=None) # Disable interpolation
    try:
        config.read(config_path, encoding='utf-8') # Specify encoding

        settings = {
            'source_file': config.get('Files', 'source_file'),
            'dn_api_url': config.get('API', 'dn_url'),
            'agent_group_url': config.get('API', 'agent_group_url'),
            'ideal_agent_header_text': config.get('SheetLayout', 'ideal_agent_header_text'),
            'ideal_agent_fallback_cell': config.get('SheetLayout', 'ideal_agent_fallback_cell'),
            'vag_extraction_sheet': config.get('SheetLayout', 'vag_extraction_sheet'),
            'api_timeout': config.getint('API', 'timeout', fallback=15), # Optional timeout
        }
        logging.info("Configuration loaded successfully.")

        # Basic validation
        if not settings['source_file'].lower().endswith('.xlsx'):
             msg = f"Source file '{settings['source_file']}' in config must be an .xlsx file."
             logging.error(msg)
             raise ValueError(msg)

        # Validate fallback cell format
        try:
            openpyxl_cell_utils.coordinate_to_tuple(settings['ideal_agent_fallback_cell'])
        except openpyxl_cell_utils.IllegalCharacterError:
            msg = f"Invalid format for 'ideal_agent_fallback_cell' in config: {settings['ideal_agent_fallback_cell']}"
            logging.error(msg)
            raise ValueError(msg)

        return settings

    except (configparser.NoSectionError, configparser.NoOptionError) as e:
        logging.error(f"Configuration error in '{config_path}': {e}")
        raise ValueError(f"Configuration error in '{config_path}': {e}")
    except Exception as e:
        logging.error(f"Error reading configuration '{config_path}': {e}", exc_info=True)
        raise ValueError(f"Error reading configuration '{config_path}': {e}")


# --- Excel Utilities ---
def copy_cell_style(source_cell: openpyxl.cell.Cell, target_cell: openpyxl.cell.Cell):
    """Copies font, fill, and alignment style from source_cell to target_cell."""
    if source_cell.has_style:
        # Copy Font properties
        target_cell.font = Font(name=source_cell.font.name,
                                size=source_cell.font.size,
                                bold=source_cell.font.bold,
                                italic=source_cell.font.italic,
                                vertAlign=source_cell.font.vertAlign,
                                underline=source_cell.font.underline,
                                strike=source_cell.font.strike,
                                color=source_cell.font.color)
        # Copy Fill properties
        target_cell.fill = PatternFill(fill_type=source_cell.fill.fill_type,
                                       start_color=source_cell.fill.start_color,
                                       end_color=source_cell.fill.end_color)
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
        # Apply default styles if source has no style
        target_cell.font = Font()
        target_cell.fill = PatternFill()
        target_cell.alignment = openpyxl.styles.Alignment()
        target_cell.number_format = 'General'


def identify_ideal_agent_column(sheet: openpyxl.worksheet.worksheet.Worksheet, config: Dict[str, Any]) -> Optional[int]:
    """
    Identifies the column index for the 'Ideal Agent' column based on config.
    Searches header row 1 (Cols C, D) first, then checks a fallback cell.

    Args:
        sheet: The openpyxl worksheet object.
        config: The loaded configuration dictionary.

    Returns:
        The column index (1-based) if found, otherwise None.
    """
    header_text = config['ideal_agent_header_text']
    fallback_cell = config['ideal_agent_fallback_cell']
    logging.debug(f"Identifying '{header_text}' column in sheet: {sheet.title}")

    # Check headers in row 1, columns C and D
    for col_idx in [3, 4]:
        cell_value = sheet.cell(row=1, column=col_idx).value
        if cell_value and header_text in str(cell_value):
            logging.debug(f"Found '{header_text}' in header at column {col_idx}")
            return col_idx

    # Check specific fallback cell from config
    try:
        col_str, row_str = openpyxl_cell_utils.coordinate_to_tuple(fallback_cell)
        fallback_col_idx = openpyxl_cell_utils.column_index_from_string(col_str)
        fallback_row_idx = int(row_str)

        cell_value_fallback = sheet.cell(row=fallback_row_idx, column=fallback_col_idx).value
        if cell_value_fallback and header_text in str(cell_value_fallback):
            logging.debug(f"Found '{header_text}' at fallback cell {fallback_cell} (Col {fallback_col_idx})")
            return fallback_col_idx
    except Exception as e:
         logging.warning(f"Could not parse or check fallback cell '{fallback_cell}': {e}")

    logging.warning(f"'{header_text}' column not found using header search or fallback cell in sheet: {sheet.title}.")
    return None

def extract_skills(expression: str) -> list[str]:
    """Extracts potential skill names (alphanumeric + underscore) from a skill expression string."""
    skills = re.findall(r'\b([a-zA-Z0-9_]+)(?=>\d+)', expression)
    logging.debug(f"Extracted skills {skills} from expression '{expression}'")
    return skills


# --- API Interaction ---
def fetch_api_data(config: Dict[str, Any]) -> Dict[str, Dict[str, Any]]: # Return type changed slightly
    """
    Fetches routing entity data (VQs, Skills, VAGs, Skill Expressions)
    from APIs specified in the configuration.
    MODIFIED: Stores more detail for skill_exprs from API.

    Args:
        config: The loaded configuration dictionary.

    Returns:
        A dictionary containing the fetched API data. Structure:
        {
            "vqs": {norm_name: id_str},
            "skills": {norm_name: id_str},
            "vags": {norm_name: id_str},
            "skill_exprs": {concat_key: {"id": id_str, "expr": expr_str, "ideal": ideal_str}} # <-- Changed structure
        }
    """
    logging.info("Fetching API data...")
    dn_url = config['dn_api_url']
    agent_group_url = config['agent_group_url']
    timeout = config['api_timeout']

    api_data = {"vqs": {}, "skills": {}, "skill_exprs": {}, "vags": {}}

    try:
        # Fetch DN (VQ) data
        logging.debug(f"Fetching DN data from {dn_url} with timeout={timeout}s")
        dn_response = requests.get(dn_url, timeout=timeout)
        dn_response.raise_for_status()
        dn_json = dn_response.json()
        logging.info(f"Successfully fetched DN response ({len(dn_json)} items).")

        # Fetch Agent Group data
        logging.debug(f"Fetching Agent Group data from {agent_group_url} with timeout={timeout}s")
        ag_response = requests.get(agent_group_url, timeout=timeout)
        ag_response.raise_for_status()
        ag_json = ag_response.json()
        logging.info(f"Successfully fetched Agent Group response ({len(ag_json)} items).")

    except requests.exceptions.Timeout:
         logging.error(f"API request timed out after {timeout} seconds.")
         print(f"ERROR: API request timed out. Check API URLs and network connectivity.")
         return api_data
    except requests.exceptions.RequestException as e:
        logging.error(f"API fetch failed: {e}")
        print(f"ERROR: Failed to fetch data from APIs. Check URLs and network. Details in {LOG_FILE}.")
        return api_data

    # --- Process DN (VQ) data ---
    vq_count = 0
    for item in dn_json:
        data = item.get('data', {})
        vq_name = data.get('name')
        vq_id = data.get('id')
        if vq_name and vq_id is not None:
            normalized_vq = vq_name.replace(" ", "").replace('\u00A0', '')
            api_data["vqs"][normalized_vq] = str(vq_id)
            logging.debug(f"Processed VQ: Name='{normalized_vq}', ID='{vq_id}'")
            vq_count += 1
        else:
            logging.warning(f"Skipping DN item due to missing name or id: {item}")
    logging.info(f"Processed {vq_count} VQs from API.")


    # --- Process Agent Group (Skill, Skill Expr, VAG) data ---
    skill_count, expr_count, vag_count = 0, 0, 0
    skipped_ag_count = 0
    for item in ag_json:
        data = item.get('data', {})
        ag_id = data.get('id')
        expression = data.get('expression', '') or ''
        ideal_expression = data.get('IdealExpression', '') or ''

        if ag_id is None:
            logging.warning(f"Skipping AG item due to missing id: {item}")
            skipped_ag_count += 1
            continue
        ag_id_str = str(ag_id)

        norm_expr = expression.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
        norm_ideal = ideal_expression.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")

        combined_expr = norm_expr
        is_skill_expr = ">" in norm_expr
        if is_skill_expr and norm_ideal:
             combined_expr = f"{norm_expr} {norm_ideal}".strip()

        if is_skill_expr: # Skill Expression
            # Store dict with details from API
            api_data["skill_exprs"][combined_expr] = {
                'id': ag_id_str,
                'expr': norm_expr,   # Store the base expression from API
                'ideal': norm_ideal # Store the ideal part from API
            }
            logging.debug(f"Processed Skill Expr: Key='{combined_expr}', ID='{ag_id_str}'")
            expr_count += 1
        elif "VAG_" in norm_expr: # VAG
            api_data["vags"][norm_expr] = ag_id_str # Store only ID for VAGs
            logging.debug(f"Processed VAG: Name='{norm_expr}', ID='{ag_id_str}'")
            vag_count += 1
        elif norm_expr: # Potentially a Simple Skill
             if re.search(r'[a-zA-Z0-9]', norm_expr):
                 api_data["skills"][norm_expr] = ag_id_str # Store only ID for Skills
                 logging.debug(f"Processed Skill: Name='{norm_expr}', ID='{ag_id_str}'")
                 skill_count += 1
             else:
                 logging.warning(f"Skipping AG item - skill name seems empty or invalid after normalization: {item}")
                 skipped_ag_count += 1
        else:
             logging.warning(f"Skipping AG item with empty expression: {item}")
             skipped_ag_count += 1

    logging.info(f"Processed Agent Groups from API: Skills={skill_count}, SkillExprs={expr_count}, VAGs={vag_count}. Skipped={skipped_ag_count}.")
    logging.info("Finished parsing API data.")
    return api_data


# --- Core Processing Logic ---
def process_sheet(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    intermediate_data: Dict[str, Dict[str, Dict[str, Any]]],
    config: Dict[str, Any]
):
    """
    Processes a single sheet to extract routing entities and their strike status,
    updating the intermediate_data dictionary according to the strike-preference rule.
    Stores separate expr/ideal for skill_exprs.

    Args:
        sheet: The openpyxl worksheet object to process.
        intermediate_data: The dictionary holding collected data across sheets.
        config: The loaded configuration dictionary.
    """
    excluded_sheets = {"Skill Expr", "VQ", "VAG", "Skills",
                       "Vqs Comparison", "Skills Comparison",
                       "Skill_exprs Comparison", "Vags Comparison"}
    if sheet.title in excluded_sheets:
        logging.debug(f"Skipping sheet: {sheet.title}")
        return

    logging.info(f"Processing sheet: {sheet.title} ({sheet.max_row} rows)")
    ideal_agent_col = identify_ideal_agent_column(sheet, config)
    vag_sheet_name = config['vag_extraction_sheet']

    # Helper for VQ, VAG, Skills
    def update_intermediate_generic(data_dict: Dict, key: str, current_strike: bool, cell_obj: openpyxl.cell.Cell):
        """Updates the dict, preferring non-struck entries."""
        if not key: return
        if key not in data_dict:
            data_dict[key] = {"strike": current_strike, "style_cell": cell_obj}
            logging.debug(f"Sheet '{sheet.title}' Row {cell_obj.row}: Added new item '{key}' with strike={current_strike}")
        elif data_dict[key]["strike"] and not current_strike:
            data_dict[key]["strike"] = False
            data_dict[key]["style_cell"] = cell_obj
            logging.debug(f"Sheet '{sheet.title}' Row {cell_obj.row}: Updated item '{key}' to strike=False")

    processed_cells = 0
    for row_idx in range(1, sheet.max_row + 1):
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value is None or str(cell.value).strip() == "":
                continue

            value_str = str(cell.value).strip()
            strike = bool(cell.font and cell.font.strike)
            processed_cells += 1

            # --- Skill Expression Processing (Specific Handling) ---
            if ">" in value_str:
                formatted_expr = value_str.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                ideal_agent_value = ""
                if ideal_agent_col:
                    if ideal_agent_col <= sheet.max_column:
                        ideal_cell = sheet.cell(row=row_idx, column=ideal_agent_col)
                        ideal_agent_value = str(ideal_cell.value).strip() if ideal_cell.value else ""
                        ideal_agent_value = ideal_agent_value.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                    else:
                        logging.warning(f"Ideal agent column {ideal_agent_col} out of bounds for sheet '{sheet.title}' (max col: {sheet.max_column}) at row {row_idx}")

                concatenated_value = f"{formatted_expr} {ideal_agent_value}".strip()
                if not concatenated_value: continue # Skip if key becomes empty

                skill_expr_dict = intermediate_data["skill_exprs"]
                if concatenated_value not in skill_expr_dict:
                    skill_expr_dict[concatenated_value] = {
                        "strike": strike,
                        "style_cell": cell,
                        "expr": formatted_expr,  # Store separate expression
                        "ideal": ideal_agent_value # Store separate ideal expression
                    }
                    logging.debug(f"Sheet '{sheet.title}' Row {cell.row}: Added new skill_expr '{concatenated_value}' with strike={strike}")
                elif skill_expr_dict[concatenated_value]["strike"] and not strike:
                    skill_expr_dict[concatenated_value]["strike"] = False
                    skill_expr_dict[concatenated_value]["style_cell"] = cell
                    logging.debug(f"Sheet '{sheet.title}' Row {cell.row}: Updated skill_expr '{concatenated_value}' to strike=False")

                # Extract individual skills for the "Skills" collection
                individual_skills = extract_skills(formatted_expr)
                for skill in individual_skills:
                     update_intermediate_generic(intermediate_data["skills"], skill, strike, cell)

            # --- VQ Processing ---
            elif (value_str.startswith("VQ_") or value_str.startswith("vq_") or ("VQ" in value_str)) and ">" not in value_str:
                 normalized_vq = value_str.replace(" ", "").replace('\u00A0', '')
                 update_intermediate_generic(intermediate_data["vqs"], normalized_vq, strike, cell)

            # --- VAG Processing (Only from specific sheet) ---
            elif "VAG_" in value_str and sheet.title == vag_sheet_name:
                 normalized_vag = value_str.replace(" ", "").replace('\u00A0', '')
                 update_intermediate_generic(intermediate_data["vags"], normalized_vag, strike, cell)

    logging.debug(f"Processed {processed_cells} non-empty cells in sheet '{sheet.title}'.")


def collect_routing_entities(
    workbook: openpyxl.workbook.Workbook,
    config: Dict[str, Any]
) -> Tuple[Dict[str, Set[str]], Dict[str, Dict[str, Dict[str, Any]]]]: # Return type changed
    """
    Processes workbook sheets, populates intermediate data, resolves strike status,
    populates final output sheets (incl. separate expr/ideal), and returns
    sheet data (non-struck only) for comparison AND the full intermediate data.
    Modifies the workbook object in place by deleting/adding sheets.

    Returns:
        Tuple containing:
        - sheet_data_for_comparison: Dict[str, Set[str]]
        - intermediate_data: Dict[str, Dict[str, Dict[str, Any]]]
    """
    logging.info("Starting collection and processing of routing entities from workbook sheets.")

    intermediate_data = { "vqs": {}, "skills": {}, "skill_exprs": {}, "vags": {} }

    # --- Phase 1: Process all sheets ---
    for sheet in workbook.worksheets:
        process_sheet(sheet, intermediate_data, config)
    logging.info("Finished processing all sheets, resolved strikethrough status.")

    # --- Phase 2: Create/Populate Output Sheets ---
    logging.info("Populating dedicated output sheets (VQ, Skills, Skill Expr, VAG)...")

    output_sheet_specs = {
        # Headers define columns for the output sheets
        "Skill Expr": {"key": "skill_exprs", "headers": ["Expression", "Ideal Expression", "Concatenated Key", "HasStrikeThrough"]},
        "VQ": {"key": "vqs", "headers": ["VQ Name", "HasStrikeThrough"]},
        "VAG": {"key": "vags", "headers": ["VAG Name", "HasStrikeThrough"]},
        "Skills": {"key": "skills", "headers": ["Skill", "HasStrikeThrough"]}
    }

    # Define sheets to remove (output + comparison)
    # Use sheet title prefixes from specs for comparison sheet names
    comparison_prefixes = {"Skill Expr": "Skill_exprs", "VQ": "Vqs", "VAG": "Vags", "Skills": "Skills"}
    sheets_to_remove = list(output_sheet_specs.keys()) + [f"{comparison_prefixes[t]} Comparison" for t in output_sheet_specs.keys()]
    for sheet_name in sheets_to_remove:
         if sheet_name in workbook.sheetnames:
             try: del workbook[sheet_name]; logging.debug(f"Removed existing sheet: {sheet_name}")
             except Exception as e: logging.warning(f"Could not remove sheet {sheet_name}: {e}")

    # Create and populate new output sheets
    for title, spec in output_sheet_specs.items():
        sheet = workbook.create_sheet(title=title)
        # Write Headers
        for col_idx, header in enumerate(spec["headers"], start=1):
             sheet.cell(row=1, column=col_idx, value=header).font = Font(bold=True) # Make headers bold

        row_num = 2
        data_items = intermediate_data.get(spec["key"], {})

        for item_key, data in sorted(data_items.items()):
             if title == "Skill Expr":
                 # Order: Expression, Ideal Expression, Concatenated Key, HasStrikeThrough
                 sheet.cell(row=row_num, column=1, value=data.get("expr", ""))
                 sheet.cell(row=row_num, column=2, value=data.get("ideal", ""))
                 cell_key = sheet.cell(row=row_num, column=3, value=item_key)
                 sheet.cell(row=row_num, column=4, value=str(data["strike"]))
                 copy_cell_style(data["style_cell"], cell_key) # Copy style to the Key cell
             else:
                 # Standard population for VQ, VAG, Skills
                 cell_a = sheet.cell(row=row_num, column=1, value=item_key)
                 sheet.cell(row=row_num, column=2, value=str(data["strike"]))
                 copy_cell_style(data["style_cell"], cell_a)
             row_num += 1
        logging.debug(f"Populated '{title}' sheet with {row_num - 2} items.")
    logging.info("Finished populating output sheets.")

    # --- Phase 3: Prepare data for comparison (uses concatenated key for skill_exprs) ---
    sheet_data_for_comparison = {
        key: {name for name, data in items.items() if not data["strike"]}
        for key, items in intermediate_data.items()
    }

    logging.info("Prepared final sheet data for comparison (non-struck items only).")
    logging.debug(f"Comparison Data Summary: VQs={len(sheet_data_for_comparison['vqs'])}, Skills={len(sheet_data_for_comparison['skills'])}, SkillExprs={len(sheet_data_for_comparison['skill_exprs'])}, VAGs={len(sheet_data_for_comparison['vags'])}")

    # Return comparison data AND intermediate data
    return sheet_data_for_comparison, intermediate_data


# --- Comparison and Reporting ---
def write_comparison_sheet(
    workbook: openpyxl.workbook.Workbook,
    sheet_data: Dict[str, Set[str]],
    api_data: Dict[str, Dict[str, Any]], # Type hint updated for api_data
    intermediate_data: Dict[str, Dict[str, Dict[str, Any]]] # Added parameter
):
    """
    Compares sheet data (non-struck only) with API data and writes results.
    Adds Expression/Ideal columns to Skill_exprs Comparison sheet.

    Args:
        workbook: The openpyxl Workbook object to write to.
        sheet_data: Dict containing sets of non-struck items found in sheets.
        api_data: Dict containing items fetched from the API.
        intermediate_data: Dict containing detailed items found in sheets.
    """
    logging.info("Starting comparison and writing results to comparison sheets.")
    if not api_data:
        logging.warning("API data is empty or None, skipping comparison writing.")
        return
    if not sheet_data:
        logging.warning("Sheet data for comparison is empty or None, skipping comparison writing.")
        return

    comparison_keys_map = {
        "vqs": "Vqs", "skills": "Skills",
        "skill_exprs": "Skill_exprs", "vags": "Vags"
    }

    for key, sheet_title_prefix in comparison_keys_map.items():
        logging.info(f"Generating comparison sheet for: {key}")
        comparison_sheet_title = f"{sheet_title_prefix} Comparison"
        if comparison_sheet_title in workbook.sheetnames:
             del workbook[comparison_sheet_title]
        sheet = workbook.create_sheet(title=comparison_sheet_title)

        # --- Prepare data for comparison ---
        sheet_items_non_struck = sheet_data.get(key, set())
        api_items_dict = api_data.get(key, {}) # Structure depends on key
        api_items_keys = set(api_items_dict.keys())

        new_in_sheet = sheet_items_non_struck - api_items_keys
        missing_from_sheet_non_struck = api_items_keys - sheet_items_non_struck

        row_num = 2 # Start writing data from row 2

        # --- Set Headers based on key ---
        if key == "skill_exprs":
            headers = ["Concatenated Key", "Expression", "Ideal Expression", "ID (from API)", "Status"]
            col_widths = [40, 40, 30, 15, 30] # Example widths
        else:
            headers = [sheet_title_prefix, "ID (from API)", "Status"]
            col_widths = [40, 15, 30] # Example widths

        for col_idx, header in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            sheet.column_dimensions[openpyxl_cell_utils.get_column_letter(col_idx)].width = col_widths[col_idx-1]


        # --- Write Data Rows ---
        # Write items New in Sheet
        if new_in_sheet:
            logging.debug(f"'{key}' - Found {len(new_in_sheet)} items New in Sheet (Non-Struck).")
            for item in sorted(list(new_in_sheet)):
                if key == "skill_exprs":
                    # Lookup details from intermediate_data (originates from sheet)
                    item_details = intermediate_data['skill_exprs'].get(item, {})
                    sheet.cell(row=row_num, column=1, value=item) # Concatenated Key
                    sheet.cell(row=row_num, column=2, value=item_details.get('expr', '')) # Expression
                    sheet.cell(row=row_num, column=3, value=item_details.get('ideal', '')) # Ideal Expression
                    sheet.cell(row=row_num, column=4, value="N/A") # ID
                    sheet.cell(row=row_num, column=5, value="New in Sheet (Non-Struck)") # Status
                else:
                    # Standard 3-column layout
                    sheet.cell(row=row_num, column=1, value=item)
                    sheet.cell(row=row_num, column=2, value="N/A")
                    sheet.cell(row=row_num, column=3, value="New in Sheet (Non-Struck)")
                row_num += 1
        else:
             logging.debug(f"'{key}' - No items found only in the sheet (non-struck).")

        # Write items Missing from Sheet
        if missing_from_sheet_non_struck:
             logging.debug(f"'{key}' - Found {len(missing_from_sheet_non_struck)} items Missing from Sheet (or only Struck Out).")
             for item in sorted(list(missing_from_sheet_non_struck)):
                if key == "skill_exprs":
                    # Lookup details from api_data (originates from API)
                    # api_items_dict is api_data['skill_exprs'] here
                    api_details = api_items_dict.get(item, {}) # api_details is {'id': ..., 'expr': ..., 'ideal': ...}
                    sheet.cell(row=row_num, column=1, value=item) # Concatenated Key
                    sheet.cell(row=row_num, column=2, value=api_details.get('expr', '')) # Expression
                    sheet.cell(row=row_num, column=3, value=api_details.get('ideal', '')) # Ideal Expression
                    sheet.cell(row=row_num, column=4, value=api_details.get('id', 'ID Not Found')) # ID
                    sheet.cell(row=row_num, column=5, value="Missing in Sheet (or only Struck Out)") # Status
                else:
                    # Standard 3-column layout
                    # api_items_dict is api_data['vqs'], etc. here, value is just the ID string
                    api_id = api_items_dict.get(item, "ID Not Found") # api_id is just the string ID here
                    sheet.cell(row=row_num, column=1, value=item)
                    sheet.cell(row=row_num, column=2, value=api_id)
                    sheet.cell(row=row_num, column=3, value="Missing in Sheet (or only Struck Out)")
                row_num += 1
        else:
            logging.debug(f"'{key}' - No items found only in the API (when compared to non-struck sheet items).")

        logging.info(f"Finished comparison sheet for: {key}")


# --- Main Orchestration ---
def main():
    """Main function to orchestrate the workbook processing and API comparison."""
    try:
        config = load_configuration(CONFIG_FILE)
        source_file_path = config['source_file']
    except (FileNotFoundError, ValueError, configparser.Error) as config_e:
        logging.error(f"Halting execution due to configuration error: {config_e}")
        print(f"FATAL: Configuration Error - {config_e}. See {LOG_FILE} for details.")
        return

    base_name = os.path.splitext(source_file_path)[0]
    final_output_path = base_name + "_processed.xlsx"

    logging.info(f"--- Starting Processing Run ---")
    logging.info(f"Source Workbook: '{source_file_path}'")
    logging.info(f"Output Workbook: '{final_output_path}'")

    try:
        shutil.copyfile(source_file_path, final_output_path)
        logging.info(f"Copied source to '{final_output_path}' for processing.")
    except FileNotFoundError:
         logging.error(f"Source file not found: '{source_file_path}'. Cannot proceed.")
         print(f"FATAL: Source file '{source_file_path}' not found.")
         return
    except Exception as e:
        logging.error(f"Error copying file '{source_file_path}' to '{final_output_path}': {e}", exc_info=True)
        print(f"FATAL: Error copying source file. See {LOG_FILE} for details.")
        return

    workbook = None
    try:
        logging.info(f"Loading workbook: {final_output_path}")
        workbook = openpyxl.load_workbook(final_output_path, read_only=False, data_only=False)

        # --- Step 1: Collect data ---
        # Unpack both comparison data and intermediate data
        sheet_data_for_comparison, intermediate_data = collect_routing_entities(workbook, config)

        # --- Step 2: Fetch API data ---
        api_data = fetch_api_data(config)

        # --- Step 3: Perform comparison and write results ---
        # Pass intermediate_data to the comparison writer
        write_comparison_sheet(workbook, sheet_data_for_comparison, api_data, intermediate_data)

        # --- Step 4: Save the processed workbook ---
        logging.info(f"Saving final processed workbook to: {final_output_path}")
        workbook.save(final_output_path)
        logging.info(f"Successfully saved processed workbook.")
        logging.info(f"--- Processing Complete ---")
        print(f"\nProcessing complete. Output saved to '{final_output_path}'")
        print(f"Log file saved to '{LOG_FILE}'")

    except Exception as e:
        logging.error(f"An unexpected error occurred during the main processing: {e}", exc_info=True)
        print(f"FATAL: An unexpected error occurred during processing. See {LOG_FILE} for details.")
    finally:
         if workbook:
             try: workbook.close(); logging.debug("Workbook closed.")
             except Exception as close_e: logging.warning(f"Error closing workbook: {close_e}")


# --- Script Execution ---
if __name__ == "__main__":
    main()
