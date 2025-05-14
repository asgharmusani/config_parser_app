# -*- coding: utf-8 -*-
"""
Core Rule Engine for parsing Excel files based on user-defined rule templates.

This module defines the ExcelRuleEngine class which takes a set of rules
and an Excel workbook, then extracts data according to those rules.
The processing iterates through sheets, reads sheet data into memory once,
then applies rules to the in-memory data.
A cell is "claimed" by the first rule that primarily identifies it.
Stores cell coordinates instead of full cell objects for style reference.
"""

import logging
import re
import openpyxl
from openpyxl.utils import cell as openpyxl_cell_utils # For type hinting and column_index_from_string
from typing import Dict, Any, List, Optional, Union, Set, Tuple

logger = logging.getLogger(__name__) # Use module-specific logger

class ExcelRuleEngine:
    """
    Parses an Excel workbook based on a provided rule template (JSON).
    """

    def __init__(self, rule_template: Dict[str, Any]):
        """
        Initializes the ExcelRuleEngine with a set of rules.

        Args:
            rule_template: A dictionary representing the parsed JSON rule template.
                           Expected structure: {"Entities": [rule1, rule2, ...], "GlobalSettings": {...}}

        Raises:
            ValueError: If the rule template is malformed or missing essential parts.
        """
        logger.info("Initializing ExcelRuleEngine...")
        if not isinstance(rule_template, dict) or "Entities" not in rule_template:
            msg = "Rule template must be a dictionary with an 'Entities' key."
            logger.error(msg)
            raise ValueError(msg)

        self.rules: List[Dict[str, Any]] = rule_template["Entities"]
        if not isinstance(self.rules, list):
            msg = "The 'Entities' key in the rule template must contain a list of rules."
            logger.error(msg)
            raise ValueError(msg)

        # Load global settings with defaults
        self.global_settings = rule_template.get("GlobalSettings", {})
        self.default_skip_sheets = set(self.global_settings.get("defaultSkipSheets", ["Metadata", "Instructions", "Summary"]))
        self.global_default_check_for_strikethrough = self.global_settings.get("defaultCheckForStrikethrough", False)

        self._validate_rules()
        logger.info(f"ExcelRuleEngine initialized with {len(self.rules)} entity rules.")

    def _validate_rules(self):
        """
        Performs basic validation on the loaded entity rules.
        Checks for required keys in each rule and validates their types.
        (No changes from previous version of this method)
        """
        logger.debug("Validating entity rules...")
        for i, rule in enumerate(self.rules):
            if not isinstance(rule, dict):
                raise ValueError(f"Rule at index {i} is not a dictionary.")
            rule_name_for_error = rule.get('name', f'Unnamed rule at index {i}')
            required_keys = ["name", "identifier"] # 'enabled' defaults to True if not present
            for key in required_keys:
                if key not in rule:
                    raise ValueError(f"Rule '{rule_name_for_error}' is missing required key: '{key}'.")
            if not isinstance(rule["name"], str) or not rule["name"]:
                 raise ValueError(f"Rule name for rule at index {i} must be a non-empty string.")

            identifier = rule["identifier"]
            if not isinstance(identifier, dict):
                raise ValueError(f"Identifier for rule '{rule_name_for_error}' must be a dictionary.")
            id_type = identifier.get("type")
            id_value = identifier.get("value")
            valid_id_types = ["startswith", "contains", "exactmatch", "regex"]
            if not id_type or str(id_type).lower() not in valid_id_types: # Ensure type is string before lower()
                raise ValueError(f"Invalid or missing identifier 'type' ('{id_type}') for rule '{rule_name_for_error}'. Must be one of {valid_id_types}.")
            if not isinstance(id_value, str) or not id_value:
                raise ValueError(f"Identifier 'value' for rule '{rule_name_for_error}' must be a non-empty string.")
            if "caseSensitive" in identifier and not isinstance(identifier["caseSensitive"], bool):
                raise ValueError(f"Optional 'identifier.caseSensitive' for rule '{rule_name_for_error}' must be a boolean.")

            if "sheets" in rule and not (rule["sheets"] is None or (isinstance(rule["sheets"], list) and all(isinstance(s, str) for s in rule["sheets"]))):
                raise ValueError(f"Optional 'sheets' key for rule '{rule_name_for_error}' must be null or a list of strings.")
            if "primaryFieldKey" in rule and (not isinstance(rule["primaryFieldKey"], str) or not rule["primaryFieldKey"]):
                raise ValueError(f"Optional 'primaryFieldKey' for rule '{rule_name_for_error}' must be a non-empty string.")

            if "replaceRules" in rule:
                if not isinstance(rule["replaceRules"], list):
                    raise ValueError(f"'replaceRules' for rule '{rule_name_for_error}' must be a list.")
                for rep_rule_idx, rep_rule in enumerate(rule["replaceRules"]):
                    if not (isinstance(rep_rule, dict) and "find" in rep_rule and "replace" in rep_rule and isinstance(rep_rule["find"], str) and isinstance(rep_rule["replace"], str)):
                        raise ValueError(f"Invalid replaceRule at index {rep_rule_idx} for rule '{rule_name_for_error}'.")

            if "fetchAdditionalColumn" in rule:
                add_col_rule = rule["fetchAdditionalColumn"]
                if not isinstance(add_col_rule, dict):
                    raise ValueError(f"'fetchAdditionalColumn' for rule '{rule_name_for_error}' must be a dictionary.")
                required_add_col_keys = ["targetKeyName", "searchHeaderName", "searchInLocations"]
                for key_ac in required_add_col_keys: # Renamed key to key_ac
                    if key_ac not in add_col_rule:
                        raise ValueError(f"'fetchAdditionalColumn' for rule '{rule_name_for_error}' missing key: '{key_ac}'.")
                if not (isinstance(add_col_rule["searchInLocations"], list) and all(isinstance(s, str) for s in add_col_rule["searchInLocations"])):
                     raise ValueError(f"'searchInLocations' for 'fetchAdditionalColumn' in rule '{rule_name_for_error}' must be a list of strings.")
                if "replaceRules" in add_col_rule and not isinstance(add_col_rule["replaceRules"], list):
                        raise ValueError(f"'replaceRules' in 'fetchAdditionalColumn' for rule '{rule_name_for_error}' must be a list.")
                if "valueFromRowOffset" in add_col_rule and not isinstance(add_col_rule["valueFromRowOffset"], int):
                    raise ValueError(f"Optional 'valueFromRowOffset' in 'fetchAdditionalColumn' for rule '{rule_name_for_error}' must be an integer.")

            if "extractSubEntities" in rule:
                sub_entity_rule = rule["extractSubEntities"]
                if not isinstance(sub_entity_rule, dict):
                    raise ValueError(f"'extractSubEntities' for rule '{rule_name_for_error}' must be a dictionary.")
                required_sub_keys = ["subEntityName", "regex"]
                for key_se in required_sub_keys: # Renamed key to key_se
                    if key_se not in sub_entity_rule:
                        raise ValueError(f"'extractSubEntities' for rule '{rule_name_for_error}' missing key: '{key_se}'.")
                if "sourceValueFrom" in sub_entity_rule and sub_entity_rule["sourceValueFrom"] not in ["primaryFieldKey"] and not sub_entity_rule["sourceValueFrom"].startswith("additional."):
                    raise ValueError(f"Invalid 'sourceValueFrom' in 'extractSubEntities' for rule '{rule_name_for_error}'.")
                if "replaceRules" in sub_entity_rule and not isinstance(sub_entity_rule["replaceRules"], list):
                    raise ValueError(f"'replaceRules' in 'extractSubEntities' for rule '{rule_name_for_error}' must be a list.")

            if "constructFields" in rule:
                if not isinstance(rule["constructFields"], list):
                    raise ValueError(f"'constructFields' for rule '{rule_name_for_error}' must be a list.")
                for con_field_idx, con_field_rule in enumerate(rule["constructFields"]):
                    if not isinstance(con_field_rule, dict):
                        raise ValueError(f"Item at index {con_field_idx} in 'constructFields' for rule '{rule_name_for_error}' must be a dictionary.")
                    required_con_keys = ["targetKeyName", "formatString"]
                    for key_cf in required_con_keys: # Renamed key to key_cf
                        if key_cf not in con_field_rule:
                            raise ValueError(f"Item at index {con_field_idx} in 'constructFields' for rule '{rule_name_for_error}' missing key: '{key_cf}'.")
                    if "onMissingSource" in con_field_rule and con_field_rule["onMissingSource"] not in ["skip_field", "empty_string", "error"]:
                        raise ValueError(f"Invalid 'onMissingSource' for 'constructFields' item in rule '{rule_name_for_error}'.")
        logger.debug("All rules passed basic validation.")


    def _apply_replace_rules(self, text_value: str, replace_rules: List[Dict[str, str]]) -> str:
        """Applies a list of find/replace rules to a given text value."""
        if not isinstance(text_value, str):
            return text_value
        processed_value = text_value
        for rule_item in replace_rules:
            find_str = rule_item.get("find", "")
            replace_str = rule_item.get("replace", "")
            processed_value = processed_value.replace(find_str, replace_str)
        return processed_value

    def _match_identifier(self, cell_value_str: str, identifier_rule: Dict[str, Any]) -> bool:
        """
        Checks if a cell value matches the given identifier rule.
        Non-regex matches (startsWith, contains, exactMatch) are case-insensitive
        by default, unless 'caseSensitive': true is in the identifier_rule.
        """
        id_type = str(identifier_rule.get("type", "")).lower()
        id_value_from_rule = identifier_rule.get("value", "")
        case_sensitive = identifier_rule.get("caseSensitive", False)

        if not id_value_from_rule:
            logger.debug(f"Identifier rule missing 'value': {identifier_rule}")
            return False

        cell_val_to_compare = cell_value_str if (id_type == "regex" or case_sensitive) else cell_value_str.lower()
        id_val_to_compare = id_value_from_rule if (id_type == "regex" or case_sensitive) else id_value_from_rule.lower()

        if id_type == "startswith":
            return cell_val_to_compare.startswith(id_val_to_compare)
        elif id_type == "contains":
            return id_val_to_compare in cell_val_to_compare
        elif id_type == "exactmatch":
            return cell_val_to_compare == id_val_to_compare
        elif id_type == "regex":
            try:
                return bool(re.search(id_value_from_rule, cell_value_str)) # Regex uses original case from rule
            except re.error as e:
                logger.warning(f"Invalid regex '{id_value_from_rule}' in identifier rule: {e}")
                return False
        else:
            logger.warning(f"Unknown identifier type: '{id_type}' in rule: {identifier_rule}")
            return False

    def _find_additional_column_header_once_per_sheet(
            self,
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            fetch_additional_column_rule: Dict[str, Any],
            sheet_header_cache: Dict[str, Optional[int]] # Cache for this specific sheet
        ) -> Optional[int]:
        """
        Finds the column index for an additional column's header for a given sheet.
        Caches the result per header name to avoid re-searching the same header on the same sheet.
        """
        search_header_name = fetch_additional_column_rule.get("searchHeaderName")
        search_in_locations = fetch_additional_column_rule.get("searchInLocations", [])

        if not search_header_name: # Ensure header name is provided
            logger.warning("fetchAdditionalColumn rule missing 'searchHeaderName'.")
            return None

        cache_key = search_header_name # Cache by header name for this sheet
        if cache_key in sheet_header_cache:
            return sheet_header_cache[cache_key]

        found_column_idx = None
        for loc in search_in_locations:
            try:
                if re.fullmatch(r'[A-Z]+', loc, re.IGNORECASE): # Column letter
                    col_idx_from_letter = openpyxl_cell_utils.column_index_from_string(loc)
                    if col_idx_from_letter <= sheet.max_column:
                        header_cell_value = sheet.cell(row=1, column=col_idx_from_letter).value
                        if header_cell_value and search_header_name in str(header_cell_value):
                            found_column_idx = col_idx_from_letter
                            break
                elif re.fullmatch(r'[A-Z]+[1-9][0-9]*', loc, re.IGNORECASE): # Cell address
                    col_str, row_str = openpyxl_cell_utils.coordinate_to_tuple(loc)
                    header_row_idx = int(row_str)
                    header_col_idx = openpyxl_cell_utils.column_index_from_string(col_str)
                    if header_row_idx <= sheet.max_row and header_col_idx <= sheet.max_column:
                        header_cell_value = sheet.cell(row=header_row_idx, column=header_col_idx).value
                        if header_cell_value and search_header_name in str(header_cell_value):
                            found_column_idx = header_col_idx
                            break
            except Exception as e:
                logger.warning(f"Error processing searchIn location '{loc}' for header '{search_header_name}' on sheet '{sheet.title}': {e}")
                continue

        sheet_header_cache[cache_key] = found_column_idx
        if found_column_idx:
            logger.debug(f"Header '{search_header_name}' found in column {found_column_idx} for sheet '{sheet.title}'. Caching.")
        else:
            logger.debug(f"Header '{search_header_name}' not found in specified locations for sheet '{sheet.title}'. Caching as None.")
        return found_column_idx


    def _fetch_additional_column_data_from_row(
        self,
        current_row_idx: int,
        sheet: openpyxl.worksheet.worksheet.Worksheet,
        found_column_idx: int,
        replace_rules: List[Dict[str, str]],
        value_from_row_offset: int = 0 # Added offset
    ) -> Optional[str]:
        """
        Fetches and cleans data from a specific cell, potentially offset from the current row,
        given the pre-determined column index of the additional data.
        """
        target_row_idx = current_row_idx + value_from_row_offset
        if not (1 <= target_row_idx <= sheet.max_row):
            logger.warning(f"Target row {target_row_idx} (current: {current_row_idx} + offset: {value_from_row_offset}) is out of bounds for sheet '{sheet.title}'.")
            return None

        if found_column_idx > sheet.max_column:
            logger.warning(f"Attempting to fetch from column {found_column_idx} (max: {sheet.max_column}) for target row {target_row_idx} in sheet '{sheet.title}'.")
            return None

        additional_cell_value = sheet.cell(row=target_row_idx, column=found_column_idx).value
        if additional_cell_value is not None:
            value_str = str(additional_cell_value).strip()
            if replace_rules:
                value_str = self._apply_replace_rules(value_str, replace_rules)
            return value_str
        return None

    def _extract_sub_entities(
        self,
        source_text: str,
        sub_entity_rule: Dict[str, Any],
        primary_cell_strikethrough: bool
    ) -> List[Dict[str, Any]]:
        """ Extracts sub-entities from a source text using regex. """
        if not source_text or not isinstance(source_text, str): return []
        regex_pattern = sub_entity_rule.get("regex")
        replace_rules = sub_entity_rule.get("replaceRules", [])
        apply_primary_strike = sub_entity_rule.get("checkForStrikethrough", False)
        final_strike_status = primary_cell_strikethrough if apply_primary_strike else False
        if not regex_pattern: logger.warning("Missing 'regex' in 'extractSubEntities' rule."); return []
        extracted_values = []
        try:
            matches = re.findall(regex_pattern, source_text)
            for match_value in matches:
                value_to_clean = match_value[0] if isinstance(match_value, tuple) and match_value else match_value
                cleaned_value = str(value_to_clean).strip()
                if replace_rules: cleaned_value = self._apply_replace_rules(cleaned_value, replace_rules)
                if cleaned_value: extracted_values.append({"value": cleaned_value, "strike": final_strike_status})
        except re.error as e: logger.error(f"Invalid regex '{regex_pattern}' in 'extractSubEntities': {e}")
        except Exception as e: logger.error(f"Error extracting sub-entities with regex '{regex_pattern}': {e}", exc_info=True)
        return extracted_values

    def _construct_field(
            self,
            format_string: str,
            current_entity_data: Dict[str, Any],
            primary_key_name_for_this_rule: str,
            on_missing_source: str
        ) -> Optional[str]:
        """ Constructs a new field value based on a format string and existing entity data. """
        placeholder_pattern = re.compile(r'{([^}]+)}')
        def replace_match(match):
            field_name_to_lookup = match.group(1).strip()
            if field_name_to_lookup == "_primary_":
                return str(current_entity_data.get(primary_key_name_for_this_rule, ""))
            elif field_name_to_lookup in current_entity_data:
                return str(current_entity_data.get(field_name_to_lookup, ""))
            else: # Source field is missing
                if on_missing_source == "empty_string": logger.debug(f"Missing source field '{field_name_to_lookup}' for constructField, using empty string."); return ""
                elif on_missing_source == "error": logger.error(f"Missing source field '{field_name_to_lookup}' for constructField (onMissingSource=error)."); raise KeyError(f"Missing source field '{field_name_to_lookup}' for constructField (onMissingSource=error).")
                elif on_missing_source == "skip_field": logger.debug(f"Missing source field '{field_name_to_lookup}' for constructField, will skip field."); raise ValueError("_SKIP_CONSTRUCT_FIELD_")
                return match.group(0)
        try:
            return placeholder_pattern.sub(replace_match, format_string)
        except ValueError as e:
            if str(e) == "_SKIP_CONSTRUCT_FIELD_": return None
            raise
        except KeyError as e: raise e


    def process_workbook(self, workbook: openpyxl.workbook.Workbook) -> Dict[str, List[Dict[str, Any]]]:
        """
        Processes the entire workbook based on the loaded rules.
        Iterates sheets first, reads sheet data into memory, then applies rules.
        A cell is "claimed" as a primary entity by the first rule that identifies it.

        Args:
            workbook: An openpyxl Workbook object.

        Returns:
            A dictionary where keys are entity names (from rules) and values are lists
            of extracted row data dictionaries.
        """
        logger.info(f"Processing workbook with {len(self.rules)} rules using sheet-first optimized approach.")
        parsed_entities: Dict[str, List[Dict[str, Any]]] = {}
        sheet_header_location_cache: Dict[str, Dict[str, Optional[int]]] = {}
        claimed_primary_cells: Set[Tuple[str, int, int]] = set()

        # Initialize result structure for each enabled rule
        for rule in self.rules:
            if rule.get("enabled", True): # Default to enabled if key is missing
                parsed_entities[rule["name"]] = []

        # 1. Iterate through SHEETS first
        for sheet in workbook.worksheets:
            # Apply global skip sheets
            is_globally_skipped = sheet.title in self.default_skip_sheets
            if is_globally_skipped:
                is_explicitly_included_by_a_rule = False
                for rule_check in self.rules:
                    if rule_check.get("enabled", True):
                        rule_sheets_filter_check = rule_check.get("sheets")
                        if rule_sheets_filter_check and sheet.title in rule_sheets_filter_check:
                            is_explicitly_included_by_a_rule = True
                            break
                if not is_explicitly_included_by_a_rule:
                    logger.info(f"Skipping sheet: {sheet.title} (in global skip list and not explicitly included by any active rule).")
                    continue

            logger.info(f"Processing sheet: {sheet.title}")
            if sheet.title not in sheet_header_location_cache:
                sheet_header_location_cache[sheet.title] = {}

            # OPTIMIZATION: Read sheet data into memory once per sheet
            # Stores tuples of (cell_object_for_style, cell_value_string)
            logger.info(f"Processing sheet: {sheet.title} Stores tuples of ")
            sheet_data_in_memory: List[List[Tuple[Optional[openpyxl.cell.Cell], str]]] = []
            for row_idx_read in range(1, sheet.max_row + 1):
                current_row_data_mem = []
                for col_idx_read in range(1, sheet.max_column + 1):
                    cell_obj_mem = sheet.cell(row=row_idx_read, column=col_idx_read)
                    cell_val_mem = str(cell_obj_mem.value).strip() if cell_obj_mem.value is not None else ""
                    current_row_data_mem.append((cell_obj_mem, cell_val_mem))
                sheet_data_in_memory.append(current_row_data_mem)
            logger.debug(f"Read {len(sheet_data_in_memory)} rows from sheet '{sheet.title}' into memory.")


            # 2. Iterate through in-memory SHEET DATA (Rows and Cells)
            for row_idx_mem, row_content in enumerate(sheet_data_in_memory):
                actual_row_idx = row_idx_mem + 1 # 1-based for Excel

                for col_idx_mem, (cell_obj_from_mem, cell_value_str_from_mem) in enumerate(row_content):
                    actual_col_idx = col_idx_mem + 1 # 1-based for Excel

                    cell_coordinate_tuple = (sheet.title, actual_row_idx, actual_col_idx)
                    if cell_coordinate_tuple in claimed_primary_cells:
                        continue # This cell already claimed as primary for another rule

                    if not cell_value_str_from_mem: # Skip empty string values from memory
                        continue

                    # 3. Iterate through RULES for the current cell
                    for rule in self.rules:
                        if not rule.get("enabled", True):
                            continue

                        rule_sheets_filter = rule.get("sheets")
                        if rule_sheets_filter is not None and sheet.title not in rule_sheets_filter:
                            continue

                        # 4. Apply CURRENT rule's identifier to the cell's string value
                        if self._match_identifier(cell_value_str_from_mem, rule["identifier"]):
                            logger.debug(f"MATCH: Rule '{rule['name']}', Sheet '{sheet.title}', Cell {cell_obj_from_mem.coordinate if cell_obj_from_mem else (actual_row_idx, actual_col_idx)}, Value '{cell_value_str_from_mem}'")
                            claimed_primary_cells.add(cell_coordinate_tuple) # Claim this cell for this rule

                            primary_value = cell_value_str_from_mem
                            rule_check_strike = rule["identifier"].get("checkForStrikethrough", self.global_default_check_for_strikethrough)
                            primary_strike_status = bool(cell_obj_from_mem and cell_obj_from_mem.font and cell_obj_from_mem.font.strike if rule_check_strike else False)

                            if "replaceRules" in rule:
                                primary_value = self._apply_replace_rules(primary_value, rule["replaceRules"])

                            entity_data: Dict[str, Any] = {}
                            primary_key_name = rule.get("primaryFieldKey", rule["name"])
                            entity_data[primary_key_name] = primary_value
                            entity_data["strike"] = primary_strike_status
                            # --- MODIFICATION: Store coordinate instead of full cell object ---
                            entity_data["_source_sheet_title_"] = sheet.title # For retrieving cell later
                            entity_data["_source_cell_coordinate_"] = cell_obj_from_mem.coordinate if cell_obj_from_mem else f"R{actual_row_idx}C{actual_col_idx}"
                            # --- END MODIFICATION ---

                            if "fetchAdditionalColumn" in rule:
                                add_col_config = rule["fetchAdditionalColumn"]
                                header_col_idx_found = self._find_additional_column_header_once_per_sheet(
                                    sheet, add_col_config, sheet_header_location_cache[sheet.title]
                                )
                                if header_col_idx_found:
                                    offset = add_col_config.get("valueFromRowOffset", 0)
                                    additional_value = self._fetch_additional_column_data_from_row(
                                        actual_row_idx, sheet, header_col_idx_found, add_col_config.get("replaceRules", []), offset
                                    )
                                    if additional_value is not None:
                                        target_key = add_col_config.get("targetKeyName")
                                        if target_key: entity_data[target_key] = additional_value
                                        else: logger.warning(f"Missing 'targetKeyName' in fetchAdditionalColumn for rule '{rule['name']}'.")

                            if "extractSubEntities" in rule:
                                sub_entity_config = rule["extractSubEntities"]
                                source_value_for_sub_extraction = ""
                                source_value_key_from_rule = sub_entity_config.get("sourceValueFrom", "primaryFieldKey")
                                if source_value_key_from_rule == "primaryFieldKey":
                                    source_value_for_sub_extraction = entity_data.get(primary_key_name, "")
                                elif source_value_key_from_rule.startswith("additional."):
                                    additional_key = source_value_key_from_rule.split("additional.", 1)[1]
                                    source_value_for_sub_extraction = entity_data.get(additional_key, "")
                                else: logger.warning(f"Invalid 'sourceValueFrom' ('{source_value_key_from_rule}') for sub-entity extraction in rule '{rule['name']}'.")
                                if source_value_for_sub_extraction:
                                    sub_entities = self._extract_sub_entities(source_value_for_sub_extraction, sub_entity_config, primary_strike_status)
                                    if sub_entities:
                                        sub_entity_list_key = sub_entity_config.get("subEntityName")
                                        if sub_entity_list_key: entity_data[sub_entity_list_key] = sub_entities
                                        else: logger.warning(f"Missing 'subEntityName' in extractSubEntities rule for '{rule['name']}'.")

                            if "constructFields" in rule:
                                for construct_rule_config in rule["constructFields"]:
                                    target_key = construct_rule_config.get("targetKeyName")
                                    format_string = construct_rule_config.get("formatString")
                                    on_missing = construct_rule_config.get("onMissingSource", "skip_field")
                                    if target_key and format_string:
                                        try:
                                            constructed_value = self._construct_field(format_string, entity_data, primary_key_name, on_missing)
                                            if constructed_value is not None: entity_data[target_key] = constructed_value
                                        except KeyError as e: logger.error(f"Error constructing field '{target_key}' for rule '{rule['name']}': {e}")
                                    else: logger.warning(f"Missing 'targetKeyName' or 'formatString' in constructFields for rule '{rule['name']}'.")

                            parsed_entities[rule["name"]].append(entity_data)
                            break # This cell has been claimed by this rule for primary identification.
        logger.info("Finished processing workbook with rule engine.")
        return parsed_entities

