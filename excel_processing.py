# -*- coding: utf-8 -*-
"""
Processes data extracted by the ExcelRuleEngine, resolves strike-through status,
creates output sheets in the workbook based ONLY on entities found by the rule engine,
and prepares data structures for the comparison logic.
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
            and '_style_cell_object_'. The primary identifying value is expected
            to be under a key like the entity name or 'primaryFieldKey' from the rule.

    Returns:
        A dictionary where keys are entity names, and values are dictionaries
        of items. Each item dictionary contains its resolved 'strike' status,
        '_style_cell_object_', and all other extracted fields.
        Format: {'EntityName1': {'item_key1': {'strike': False, '_style_cell_object_': ..., 'field1': ...}, ...}}
    """
    logger.info("Resolving strike-through status and preparing intermediate data...")
    intermediate_data_resolved: Dict[str, Dict[str, Dict[str, Any]]] = {}

    # Iterate through each entity type (e.g., "VQs", "Skill_Expressions") from parsed_entities
    for entity_name, entity_occurrences in parsed_entities.items():
        if entity_name not in intermediate_data_resolved:
            intermediate_data_resolved[entity_name] = {}
        # This is a dictionary for the current entity type, e.g., intermediate_data_resolved['VQs']
        # It will store items by their unique key_value, with resolved strike status.
        processed_entity_items = intermediate_data_resolved[entity_name]

        # Iterate through each occurrence of an item for the current entity type
        for occurrence_data in entity_occurrences:
            # Determine the primary key for this item.
            # This relies on the rule engine outputting a consistent primary field.
            # The rule should specify a 'primaryFieldKey'; if not, fallback to the entity 'name'.
            # The value under this key in occurrence_data is the unique identifier for the item.
            item_key_value = None
            # The rule engine should have stored the primary value under the key defined by
            # 'primaryFieldKey' in the rule, or rule['name'] if 'primaryFieldKey' was absent.
            # We need to find this primary identifying value in the occurrence_data.

            # Attempt to find the primary identifying value.
            # This logic assumes the rule engine stores the primary value under a key that matches the entity_name
            # OR under a key specified by `primaryFieldKey` in the rule, which then became a key in occurrence_data.
            # A more robust way is if the rule engine *always* puts the primary identifier value
            # under a consistent, known key (e.g., '_primary_identifier_value_') in occurrence_data.
            # For now, we try the entity_name as a key, then iterate other keys.
            potential_key_value = None
            # Check if the entity_name itself is a key in the occurrence_data.
            # This would happen if the rule's 'primaryFieldKey' was not set or was the same as 'name'.
            if entity_name in occurrence_data:
                potential_key_value = occurrence_data[entity_name]
            else:
                # Fallback: iterate keys to find the first non-internal, non-detail one as the identifier.
                # This is less reliable and depends on the order and nature of keys.
                for k, v_val in occurrence_data.items():
                    # Exclude internal keys and common detail keys that are unlikely to be primary identifiers.
                    if not k.startswith('_') and k not in [
                        'strike', 'expr', 'ideal', 'Expression', 'Ideal Expression',
                        'Concatenated Key', 'ID', 'Status', 'Item' # Common column names
                    ]:
                        potential_key_value = v_val # Assume the value of this field is the identifier
                        logger.debug(f"Using fallback key '{k}' with value '{v_val}' as identifier for entity '{entity_name}'.")
                        break
            item_key_value = str(potential_key_value) if potential_key_value is not None else None


            if not item_key_value:
                logger.warning(f"Could not determine item key for occurrence in '{entity_name}': {occurrence_data}")
                continue # Skip this occurrence if no key can be found

            current_strike = occurrence_data.get("strike", False)
            style_cell = occurrence_data.get("_style_cell_object_") # Style from the cell where this item was identified

            if item_key_value not in processed_entity_items:
                # First time seeing this item key for this entity type
                # Create a copy to avoid modifying the original parsed_entities dicts
                processed_entity_items[item_key_value] = occurrence_data.copy()
                # Ensure 'strike' and 'style_cell' are correctly set from this first occurrence
                processed_entity_items[item_key_value]["strike"] = current_strike
                if style_cell: # Only assign if a style cell was provided
                    processed_entity_items[item_key_value]["_style_cell_object_"] = style_cell
            else:
                # Item key already exists, check strike-through resolution
                existing_item = processed_entity_items[item_key_value]
                # If the existing record for this item was marked as struck-through,
                # but this new occurrence is NOT struck-through, then update the record.
                if existing_item.get("strike", False) and not current_strike:
                    existing_item["strike"] = False # Resolve to not struck-through
                    if style_cell: # Prefer style from the non-struck cell
                        existing_item["_style_cell_object_"] = style_cell
                    # Update other fields from the non-struck occurrence if they differ.
                    # This ensures data from a non-struck instance is preferred.
                    for k, v_val in occurrence_data.items():
                        if k not in ["strike", "_style_cell_object_"]:
                            existing_item[k] = v_val
    logger.info("Strike-through resolution complete.")
    return intermediate_data_resolved


def collect_and_write_excel_outputs(
    workbook: openpyxl.workbook.Workbook,
    parsed_entities: Dict[str, List[Dict[str, Any]]], # Output from ExcelRuleEngine
    config: Dict[str, Any], # Application config (might not be heavily used here)
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
                         Format: {'EntityName1': [row_data_dict1, ...], ...}
                         Each row_data_dict should contain fields extracted by rules.
        config: The application configuration dictionary (currently unused here but passed for future).
        metadata_sheet_name: Name for the metadata sheet.
        sheets_to_remove_config: Base list of sheets to ensure are removed (e.g., ["Metadata"]).

    Returns:
        A tuple containing:
        - sheet_data_for_comparison: Dict[str, Set[str]] - Sets of non-struck primary keys
                                     for each entity type, used by comparison_logic.
        - intermediate_data_resolved: Dict - The fully processed and strike-resolved data.
    """
    logger.info("Collecting entities and writing Excel output sheets based on rule engine output.")

    # 1. Resolve strike-through and prepare final intermediate data structure
    # This aggregates multiple occurrences of the same item and resolves strike status.
    intermediate_data_resolved = resolve_strike_through_and_prepare_intermediate(parsed_entities)

    # 2. Remove old generated sheets
    # Dynamically determine all sheets to remove based on resolved entities from the rule engine
    all_sheets_to_remove = list(sheets_to_remove_config) # Start with base (e.g., Metadata)
    for entity_name in intermediate_data_resolved.keys():
        # Add the entity's output sheet name (which is the entity_name itself)
        if entity_name not in all_sheets_to_remove:
            all_sheets_to_remove.append(entity_name)
        # Add the entity's comparison sheet name
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
        if not items_data: # Skip if no items for this entity after resolution
            logger.info(f"No data found for entity '{entity_name}' after resolution, skipping output sheet creation.")
            continue

        output_sheet = workbook.create_sheet(title=entity_name)
        logger.debug(f"Created output sheet: {entity_name}")

        # Determine headers dynamically from the first item's keys (excluding internal/style keys)
        # This assumes all items of an entity type have a similar structure from the rule engine.
        sample_item = next(iter(items_data.values()), None)
        if not sample_item:
            logger.warning(f"No items to determine headers for entity '{entity_name}'. Skipping sheet population.")
            continue

        # Define a consistent order for headers if possible, or sort them.
        # The 'primaryFieldKey' from the rule should ideally be the first column.
        # For now, simple sort, excluding internal keys.
        # This list of headers will define the columns in the output sheet.
        headers = sorted([k for k in sample_item.keys() if not k.startswith('_') and k != 'strike'])

        # Ensure "HasStrikeThrough" is the last column
        if "HasStrikeThrough" in headers: # Should not be in headers from sample_item keys
            headers.remove("HasStrikeThrough") # Remove if it accidentally got in
        headers.append("HasStrikeThrough") # Add as the last header

        # Write Headers to the new sheet
        for col_idx, header_name in enumerate(headers, start=1):
            output_sheet.cell(row=1, column=col_idx, value=header_name).font = Font(bold=True)

        # Write data rows to the new sheet
        current_row_num = 2
        # Sort items by their primary key for consistent order in the output sheet
        for item_key, item_details in sorted(items_data.items()):
            for col_idx, header_name in enumerate(headers, start=1):
                if header_name == "HasStrikeThrough":
                    cell_value = str(item_details.get("strike", False))
                else:
                    # Get value by header key from the item_details dictionary
                    cell_value = item_details.get(header_name, "")

                cell_to_write = output_sheet.cell(row=current_row_num, column=col_idx, value=cell_value)

                # Apply style from the representative cell if this is the primary identifying field.
                # This assumes the item_key itself is one of the header_name values
                # or that a primary field key was used consistently by the rule engine.
                # For simplicity, let's try to style the cell corresponding to the item_key if it's a header.
                if header_name == item_key: # This condition might need adjustment
                    style_cell_obj = item_details.get("_style_cell_object_")
                    if style_cell_obj:
                        copy_cell_style(style_cell_obj, cell_to_write)
            current_row_num += 1
        logging.debug(f"Populated '{entity_name}' output sheet with {current_row_num - 2} items.")
    logging.info("Finished populating output sheets.")

    # 4. Prepare data structures for comparison_logic.py
    sheet_data_for_comparison: Dict[str, Set[str]] = {}
    for entity_name, items_data in intermediate_data_resolved.items():
        # Collect keys of non-struck items for comparison.
        # The 'item_key' here is the primary identifier for each entity instance.
        sheet_data_for_comparison[entity_name] = {
            item_key for item_key, data in items_data.items() if not data.get("strike", False)
        }

    logging.info("Prepared data for comparison logic.")
    # The intermediate_data_resolved itself contains all details needed for "New in Sheet" items
    # in the comparison sheets (like separate expr/ideal for skill expressions).
    return sheet_data_for_comparison, intermediate_data_resolved

