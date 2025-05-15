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
except ImportError:
    logging.error("Failed to import 'copy_cell_style' from utils.py in excel_processing.py")
    def copy_cell_style(s, t):
        """Dummy function if import fails."""
        pass

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
            '_source_cell_coordinate_', '_source_sheet_title_', and importantly
            '_rule_primary_field_key_' which indicates the name of the key
            holding the primary identifying value for this entity instance.

    Returns:
        A dictionary where keys are entity names, and values are dictionaries
        of items. Each item dictionary is keyed by the item's primary identifier value,
        and contains its resolved 'strike' status, source cell info, and all other fields.
        Format: {'EntityName1': {'item_primary_id_value1': {'strike': False, '_source_...': ..., 'field1': ...}, ...}}
    """
    logger.info("Resolving strike-through status and preparing intermediate data...")
    intermediate_data_resolved: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for entity_name, entity_occurrences_list in parsed_entities.items():
        if entity_name not in intermediate_data_resolved:
            intermediate_data_resolved[entity_name] = {} # Initialize as a DICT
        # This is the dictionary for the current entity type, e.g., intermediate_data_resolved['VQs']
        # It will store items by their unique primary_identifier_value, with resolved strike status.
        entity_item_map = intermediate_data_resolved[entity_name]

        for occurrence_data in entity_occurrences_list:
            # The rule engine should have added '_rule_primary_field_key_' to occurrence_data
            primary_field_key_for_this_entity = occurrence_data.get("_rule_primary_field_key_")

            if not primary_field_key_for_this_entity:
                logger.warning(f"Missing '_rule_primary_field_key_' in occurrence data for entity '{entity_name}'. Cannot determine primary identifier. Data: {occurrence_data}")
                continue

            item_primary_identifier_value = occurrence_data.get(primary_field_key_for_this_entity)

            if item_primary_identifier_value is None: # Check for None explicitly
                logger.warning(f"Primary identifier value is None for key '{primary_field_key_for_this_entity}' in entity '{entity_name}'. Data: {occurrence_data}")
                continue
            
            item_primary_identifier_value_str = str(item_primary_identifier_value)

            current_strike = occurrence_data.get("strike", False)
            source_sheet_title = occurrence_data.get("_source_sheet_title_")
            source_cell_coordinate = occurrence_data.get("_source_cell_coordinate_")

            if item_primary_identifier_value_str not in entity_item_map:
                # First time seeing this item (based on its primary identifier value)
                entity_item_map[item_primary_identifier_value_str] = occurrence_data.copy()
                # Ensure 'strike' and source info are correctly set from this first occurrence
                entity_item_map[item_primary_identifier_value_str]["strike"] = current_strike
                if source_sheet_title and source_cell_coordinate:
                    entity_item_map[item_primary_identifier_value_str]["_source_sheet_title_"] = source_sheet_title
                    entity_item_map[item_primary_identifier_value_str]["_source_cell_coordinate_"] = source_cell_coordinate
            else:
                # Item already exists, check strike-through resolution
                existing_item_details = entity_item_map[item_primary_identifier_value_str]
                if existing_item_details.get("strike", False) and not current_strike:
                    # Existing is struck-out, current is not -> update to not struck-out
                    existing_item_details["strike"] = False
                    # Prefer source info from the non-struck cell for styling
                    if source_sheet_title and source_cell_coordinate:
                        existing_item_details["_source_sheet_title_"] = source_sheet_title
                        existing_item_details["_source_cell_coordinate_"] = source_cell_coordinate
                    # Update other fields from the non-struck occurrence if they differ,
                    # or based on a defined merge strategy (currently overwrites with non-struck).
                    for k, v_val in occurrence_data.items():
                        if not k.startswith('_') and k != 'strike': # Don't overwrite strike or internal meta-keys
                            existing_item_details[k] = v_val
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
        workbook: The openpyxl.Workbook object to modify.
        parsed_entities: The output from ExcelRuleEngine.process_workbook().
        config: The application configuration dictionary.
        metadata_sheet_name: Name for the metadata sheet.
        sheets_to_remove_config: Base list of sheets to ensure are removed.

    Returns:
        A tuple containing:
        - sheet_data_for_comparison: Dict[str, Set[str]]
        - intermediate_data_resolved: Dict
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
    for entity_name, items_data in intermediate_data_resolved.items(): # items_data is Dict[item_key, item_details_dict]
        if not items_data:
            logger.info(f"No data found for entity '{entity_name}' after resolution, skipping output sheet creation.")
            continue

        output_sheet = workbook.create_sheet(title=entity_name)
        logger.debug(f"Created output sheet: {entity_name}")

        sample_item_details = next(iter(items_data.values()), None) # Get one item's details dict
        if not sample_item_details:
            logger.warning(f"No items to determine headers for entity '{entity_name}'. Skipping sheet population.")
            continue

        # Headers for the output sheet are the keys from the item_details_dict, excluding internal ones.
        headers = sorted([k for k in sample_item_details.keys() if not k.startswith('_')])
        if "strike" in headers: headers.remove("strike") # strike status is special

        # Ensure "HasStrikeThrough" is the last column
        if "HasStrikeThrough" not in headers:
            headers.append("HasStrikeThrough")
        else: # Move to end if it somehow got in
            headers.remove("HasStrikeThrough")
            headers.append("HasStrikeThrough")

        # Write Headers
        for col_idx, header_name in enumerate(headers, start=1):
            output_sheet.cell(row=1, column=col_idx, value=header_name).font = Font(bold=True)

        # Write data rows
        current_row_num = 2
        for item_key, item_details_dict in sorted(items_data.items()): # item_key is the primary identifier value
            for col_idx, header_name in enumerate(headers, start=1):
                value_for_excel_cell = None
                if header_name == "HasStrikeThrough":
                    value_for_excel_cell = str(item_details_dict.get("strike", False))
                else:
                    raw_cell_value = item_details_dict.get(header_name, "")

                    if isinstance(raw_cell_value, list):
                        # Format list of sub-entity dicts into a comma-separated string
                        formatted_sub_entities = []
                        for sub_item in raw_cell_value:
                            if isinstance(sub_item, dict) and "value" in sub_item:
                                sub_val = sub_item.get("value", "")
                                sub_strike_status = sub_item.get("strike", False)
                                formatted_sub_entities.append(f"{sub_val}{'(S)' if sub_strike_status else ''}")
                            else:
                                formatted_sub_entities.append(str(sub_item))
                        value_for_excel_cell = ", ".join(formatted_sub_entities)
                    else:
                        value_for_excel_cell = raw_cell_value
                
                cell_to_write = output_sheet.cell(row=current_row_num, column=col_idx, value=value_for_excel_cell)

                # Apply style from the original source cell to the cell containing the primary identifier value
                # The primary identifier value is item_key. We need to find which header corresponds to it.
                # The rule engine stored the name of this primary key in '_rule_primary_field_key_'.
                rule_primary_field_key = item_details_dict.get("_rule_primary_field_key_")
                if rule_primary_field_key and header_name == rule_primary_field_key:
                    source_sheet_title = item_details_dict.get("_source_sheet_title_")
                    source_cell_coord = item_details_dict.get("_source_cell_coordinate_")
                    if source_sheet_title and source_cell_coord and source_sheet_title in workbook.sheetnames:
                        try:
                            # This assumes 'workbook' is the original one for style, or styles were preserved
                            # If 'workbook' is a new one, this won't copy original styles.
                            # The 'excel_rule_engine' should pass the openpyxl cell object if styles are needed.
                            # For now, we assume it was passed via '_style_cell_object_' if that was the design.
                            # The current design stores coordinates.
                            # To copy style, we'd need to load the *original* workbook here, which is complex.
                            # Let's assume for now that style copying is best-effort or might be simplified.
                            # If '_style_cell_object_' was passed by rule engine and resolved:
                            style_cell_obj_from_resolved = item_details_dict.get("_style_cell_object_")
                            if style_cell_obj_from_resolved:
                                copy_cell_style(style_cell_obj_from_resolved, cell_to_write)
                            # else: Cannot copy style if original cell object not available here.
                        except Exception as e_style:
                            logger.warning(f"Could not apply style for {entity_name} item {item_key}, header {header_name}: {e_style}")
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

