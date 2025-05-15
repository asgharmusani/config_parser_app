# -*- coding: utf-8 -*-
"""
Processes data extracted by the ExcelRuleEngine, resolves strike-through status,
creates output sheets in the workbook based ONLY on entities found by the rule engine,
and prepares data structures for the comparison logic.
Handles formatting of sub-entity lists for Excel cell output.
"""

import logging
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import cell as openpyxl_cell_utils # For get_column_letter
from typing import Dict, Any, Optional, Tuple, Set, List

# Import utility functions from utils.py
try:
    from utils import copy_cell_style
    # extract_skills might be used if sub-entity extraction needs further processing here,
    # but primarily it should be handled by the rule engine.
except ImportError:
    logging.error("Failed to import required functions from utils.py in excel_processing.py")
    # Define dummy functions or raise error if utils are critical
    def copy_cell_style(s, t):
        """Dummy function if import fails."""
        pass
    # Consider raising an error here if utils are essential:
    # raise ImportError("Could not import utility functions. Ensure utils.py is present.")

logger = logging.getLogger(__name__) # Use module-specific logger


def resolve_strike_through_and_prepare_intermediate(
    parsed_entities: Dict[str, List[Dict[str, Any]]]
) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Resolves strike-through status for items that might appear multiple times
    (some struck, some not) and prepares the final intermediate data structure.
    If an item is found both struck and not-struck, it's considered not-struck,
    and the style from a non-struck occurrence is preferred.

    Args:
        parsed_entities: Output from ExcelRuleEngine.
            Format: {'EntityName1': [row_data_dict1, row_data_dict2], ...}
            Each row_data_dict contains extracted fields, 'strike' status,
            and '_source_cell_coordinate_', '_source_sheet_title_'.
            The primary identifying value is expected to be under a key like
            the entity name or 'primaryFieldKey' from the rule.

    Returns:
        A dictionary where keys are entity names, and values are dictionaries
        of items. Each item dictionary contains its resolved 'strike' status,
        source cell info, and all other extracted fields.
        Format: {'EntityName1': {'item_key1': {'strike': False, '_source_cell_coordinate_': 'A1', ...}, ...}}
    """
    logger.info("Resolving strike-through status and preparing intermediate data...")
    intermediate_data_resolved: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for entity_name, entity_occurrences in parsed_entities.items():
        if entity_name not in intermediate_data_resolved:
            intermediate_data_resolved[entity_name] = {}
        processed_entity_items = intermediate_data_resolved[entity_name]

        for occurrence_data in entity_occurrences:
            item_key_value = None
            # The rule engine stores the primary value under the key defined by
            # 'primaryFieldKey' or defaults to rule['name'].
            # We need to find this key in occurrence_data.
            # A more robust way: rule engine should explicitly state the primary key name used for output.
            # For now, we try common patterns or the first non-meta key.
            primary_field_key_guess = occurrence_data.get("_rule_primary_field_key", entity_name) # Assume rule engine might add this
            if primary_field_key_guess in occurrence_data:
                potential_key_value = occurrence_data[primary_field_key_guess]
            else: # Fallback
                for k, v_val in occurrence_data.items():
                    if not k.startswith('_') and k not in ['strike', 'expr', 'ideal', 'Expression', 'Ideal Expression', 'Concatenated Key', 'ID', 'Status', 'Item']:
                        potential_key_value = v_val
                        logger.debug(f"Using fallback key '{k}' with value '{v_val}' as identifier for entity '{entity_name}'.")
                        break
            item_key_value = str(potential_key_value) if potential_key_value is not None else None

            if not item_key_value:
                logger.warning(f"Could not determine item key for occurrence in '{entity_name}': {occurrence_data}")
                continue

            current_strike = occurrence_data.get("strike", False)
            # Get source cell coordinate and sheet title for style copying later
            source_sheet_title = occurrence_data.get("_source_sheet_title_")
            source_cell_coordinate = occurrence_data.get("_source_cell_coordinate_")


            if item_key_value not in processed_entity_items:
                processed_entity_items[item_key_value] = occurrence_data.copy()
                processed_entity_items[item_key_value]["strike"] = current_strike
                # Store source info for style, not the full cell object
                if source_sheet_title and source_cell_coordinate:
                    processed_entity_items[item_key_value]["_source_sheet_title_"] = source_sheet_title
                    processed_entity_items[item_key_value]["_source_cell_coordinate_"] = source_cell_coordinate
            else:
                existing_item = processed_entity_items[item_key_value]
                if existing_item.get("strike", False) and not current_strike:
                    existing_item["strike"] = False
                    if source_sheet_title and source_cell_coordinate: # Prefer source info from non-struck cell
                        existing_item["_source_sheet_title_"] = source_sheet_title
                        existing_item["_source_cell_coordinate_"] = source_cell_coordinate
                    for k, v_val in occurrence_data.items():
                        if not k.startswith('_') and k != 'strike':
                            existing_item[k] = v_val
    logger.info("Strike-through resolution complete.")
    return intermediate_data_resolved


def collect_and_write_excel_outputs(
    workbook: openpyxl.workbook.Workbook,
    parsed_entities: Dict[str, List[Dict[str, Any]]], # Output from ExcelRuleEngine
    config: Dict[str, Any], # Application config
    metadata_sheet_name: str,
    sheets_to_remove_config: List[str] # Base list (e.g., just metadata sheet name)
) -> Tuple[Dict[str, Set[str]], Dict[str, Dict[str, Dict[str, Any]]]]:
    """
    Takes parsed entities from the rule engine, resolves strike-through,
    creates output sheets in the workbook based *only* on entities found by the rule engine,
    and prepares data structures for the comparison logic.

    Args:
        workbook: The openpyxl.Workbook object to modify. This workbook is expected to
                  be the one where output sheets will be written.
        parsed_entities: The output from ExcelRuleEngine.process_workbook().
                         Format: {'EntityName1': [row_data_dict1, ...], ...}
        config: The application configuration dictionary.
        metadata_sheet_name: Name for the metadata sheet.
        sheets_to_remove_config: Base list of sheets to ensure are removed.

    Returns:
        A tuple containing:
        - sheet_data_for_comparison: Dict[str, Set[str]] - Sets of non-struck primary keys
                                     for each entity type, used by comparison_logic.
        - intermediate_data_resolved: Dict - The fully processed and strike-resolved data.
    """
    logger.info("Collecting entities and writing Excel output sheets based on rule engine output.")

    # 1. Resolve strike-through and prepare final intermediate data structure
    intermediate_data_resolved = resolve_strike_through_and_prepare_intermediate(parsed_entities)

    # 2. Remove old generated sheets
    all_sheets_to_remove = list(sheets_to_remove_config)
    for entity_name in intermediate_data_resolved.keys():
        if entity_name not in all_sheets_to_remove:
            all_sheets_to_remove.append(entity_name)
        comparison_sheet_name = f"{entity_name} Comparison"
        if comparison_sheet_name not in all_sheets_to_remove:
            all_sheets_to_remove.append(comparison_sheet_name)

    logger.info(f"Ensuring removal of sheets: {all_sheets_to_remove}")
    for sheet_name_to_remove in all_sheets_to_remove:
         if sheet_name_to_remove in workbook.sheetnames:
             try:
                 del workbook[sheet_name_to_remove]
                 logging.debug(f"Removed existing sheet: {sheet_name_to_remove}")
             except Exception as e:
                 logging.warning(f"Could not remove sheet '{sheet_name_to_remove}': {e}")

    # 3. Create and Populate new output sheets from intermediate_data_resolved
    logger.info("Populating dedicated output sheets based on processed entities...")
    for entity_name, items_data in intermediate_data_resolved.items():
        if not items_data:
            logger.info(f"No data found for entity '{entity_name}' after resolution, skipping output sheet creation.")
            continue

        output_sheet = workbook.create_sheet(title=entity_name)
        logger.debug(f"Created output sheet: {entity_name}")

        sample_item = next(iter(items_data.values()), None)
        if not sample_item:
            logger.warning(f"No items to determine headers for entity '{entity_name}'. Skipping sheet population.")
            continue

        headers = sorted([k for k in sample_item.keys() if not k.startswith('_')]) # Exclude internal keys
        if "strike" in headers: headers.remove("strike") # strike is handled by HasStrikeThrough

        # Ensure "HasStrikeThrough" is the last column if not already present (it shouldn't be)
        if "HasStrikeThrough" not in headers:
            headers.append("HasStrikeThrough")
        else: # Move to end if it exists
            headers.remove("HasStrikeThrough")
            headers.append("HasStrikeThrough")


        for col_idx, header_name in enumerate(headers, start=1):
            output_sheet.cell(row=1, column=col_idx, value=header_name).font = Font(bold=True)

        current_row_num = 2
        for item_key, item_details in sorted(items_data.items()):
            for col_idx, header_name in enumerate(headers, start=1):
                cell_value_to_write = None
                if header_name == "HasStrikeThrough":
                    cell_value_to_write = str(item_details.get("strike", False))
                else:
                    raw_cell_value = item_details.get(header_name, "")

                    # --- MODIFICATION START: Handle list of sub-entities ---
                    if isinstance(raw_cell_value, list):
                        # Assuming it's a list of sub-entity dicts like [{"value": "X", "strike": True}, ...]
                        formatted_sub_entities = []
                        for sub_item in raw_cell_value:
                            if isinstance(sub_item, dict):
                                sub_val = sub_item.get("value", "")
                                sub_strike = sub_item.get("strike", False)
                                formatted_sub_entities.append(f"{sub_val}{'(S)' if sub_strike else ''}")
                            else:
                                formatted_sub_entities.append(str(sub_item)) # Fallback for unexpected list items
                        cell_value_to_write = ", ".join(formatted_sub_entities)
                        logger.debug(f"Formatted sub-entity list for header '{header_name}': {cell_value_to_write}")
                    else:
                        cell_value_to_write = raw_cell_value
                    # --- MODIFICATION END ---

                cell_to_write = output_sheet.cell(row=current_row_num, column=col_idx, value=cell_value_to_write)

                # Apply style from the representative cell
                # The primary identifying field's cell will get the style.
                # The key of this field in item_details should be the one that matches item_key.
                # This relies on resolve_strike_through_and_prepare_intermediate ensuring item_key is a value from one of the fields.
                # A more robust way is if item_details stored which of its keys was the primary one.
                # For now, if a header matches the item_key (which is the primary identifier value), style that cell.
                if header_name == item_key:
                    source_sheet_title = item_details.get("_source_sheet_title_")
                    source_cell_coord = item_details.get("_source_cell_coordinate_")
                    if source_sheet_title and source_cell_coord and source_sheet_title in workbook.sheetnames:
                        # We need the original workbook here if it wasn't modified in place by rule engine
                        # For now, assuming 'workbook' arg is the one containing original styles
                        # This might be problematic if rule engine modified the workbook it received.
                        # Best if rule engine passes the original cell object or its style components.
                        # For now, we assume the coordinate is enough to get the cell from the *current* workbook state.
                        try:
                            original_cell_for_style = workbook[source_sheet_title][source_cell_coord]
                            copy_cell_style(original_cell_for_style, cell_to_write)
                        except KeyError:
                            logger.warning(f"Could not find source sheet '{source_sheet_title}' for styling cell {source_cell_coord}")
                        except Exception as e_style:
                            logger.warning(f"Error applying style from {source_sheet_title}!{source_cell_coord}: {e_style}")


            current_row_num += 1
        logging.debug(f"Populated '{entity_name}' output sheet with {current_row_num - 2} items.")
    logging.info("Finished populating output sheets.")

    # 4. Prepare data structures for comparison_logic.py
    sheet_data_for_comparison: Dict[str, Set[str]] = {}
    for entity_name, items_data in intermediate_data_resolved.items():
        sheet_data_for_comparison[entity_name] = {
            item_key for item_key, data in items_data.items() if not data.get("strike", False)
        }

    logging.info("Prepared data for comparison logic.")
    return sheet_data_for_comparison, intermediate_data_resolved

