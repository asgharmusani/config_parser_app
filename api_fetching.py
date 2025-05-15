# -*- coding: utf-8 -*-
"""
Handles fetching data from external APIs as specified in Excel Rule Templates.

Core Function:
- fetch_and_process_api_data_for_entity: Fetches data from a rule-specific API URL,
                                         filters items based on the rule's identifier,
                                         processes the response into a comparable format,
                                         and determines the maximum numeric ID from that response.
"""

import logging
import requests
import re # For normalizing expressions if needed
import json # For handling potential JSON decode errors
from typing import Dict, Any, Tuple, Optional, List

# Import the shared identifier matching logic from utils.py
try:
    from utils import match_identifier_logic
except ImportError:
    logging.error("Failed to import match_identifier_logic from utils.py in api_fetching.py")
    # Define a dummy function if utils.py or the function is missing, to allow startup
    def match_identifier_logic(value_to_check_str: str, identifier_rule: Dict[str, Any]) -> bool:
        """Dummy identifier matching function if import fails."""
        logging.error("Dummy match_identifier_logic called. Real function not imported from utils.py.")
        # Depending on desired fallback, this could return True to process all items,
        # or False to process no items if the real logic is critical.
        return False # Safest fallback is to not match if logic is missing.

logger = logging.getLogger(__name__) # Use module-specific logger


# fetch_max_ids_from_config_urls function has been REMOVED as Max IDs are now
# aggregated in processing_routes.py based on results from fetch_and_process_api_data_for_entity.

# --- Function to Fetch and Process Data for a Specific Entity Rule ---
def fetch_and_process_api_data_for_entity(
    api_url: str,
    entity_name: str, # The 'name' of the entity from the rule
    rule_definition: Dict[str, Any], # The full rule definition for this entity
    config: Dict[str, Any] # Global app config (for 'api_timeout')
) -> Tuple[Dict[str, Any], int]:
    """
    Fetches data from the API URL specified in an entity rule.
    Filters the API items based on the rule's identifier.
    Processes matching items into the format expected by the comparison logic.
    Calculates the maximum numeric ID found in the *filtered* API response.

    Args:
        api_url: The direct API URL from the rule's 'comparisonApiUrl'.
        entity_name: The name of the entity rule (e.g., "VQs", "Skill_Expressions").
        rule_definition: The dictionary containing the full rule for this entity,
                         including 'identifier' and optional 'apiProcessingHints'.
        config: The global application configuration (for 'api_timeout').

    Returns:
        A tuple containing:
        - processed_api_data: Dictionary structured for comparison:
                              {primary_api_identifier: details_dict_or_id_string}.
                              Returns empty dict if fetching or processing fails.
        - max_id_from_this_api: The highest numeric ID found in this *filtered and processed*
                                API response for this entity (or 0).
    """
    logger.info(f"Fetching API data for entity '{entity_name}' from URL: {api_url}")
    # Get timeout from global app config, with a default
    timeout = config.get('api_timeout', 15)
    processed_api_data: Dict[str, Any] = {}
    max_id_from_this_api = 0 # Initialize max ID for this specific API call

    if not api_url:
        logger.warning(f"No API URL provided for entity '{entity_name}'. Skipping API fetch for comparison.")
        return processed_api_data, max_id_from_this_api

    try:
        response = requests.get(api_url, timeout=timeout)
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
        raw_api_response_list = response.json() # Assuming API returns a list of items

        # Ensure the API response is a list
        if not isinstance(raw_api_response_list, list):
            logger.error(f"API response for '{entity_name}' from {api_url} is not a list. Response type: {type(raw_api_response_list)}. Response: {raw_api_response_list}")
            return processed_api_data, max_id_from_this_api
        logging.info(f"Successfully fetched {len(raw_api_response_list)} raw items for entity '{entity_name}'.")

        # Get processing hints from the rule, with defaults
        # These hints guide how to extract key fields from the API response items.
        hints = rule_definition.get("apiProcessingHints", {})
        id_field_in_api = hints.get("idField", "id") # Field in API item that holds the ID
        name_field_in_api = hints.get("nameField", "name") # Field in API item used as primary identifier/name for simple entities
        # Field from API item to use for identifier matching against the rule's identifier
        api_identifier_source_field = hints.get("apiIdentifierField", name_field_in_api)

        # Fields specific to complex entities like skill expressions
        expression_field_in_api = hints.get("expressionField", "expression")
        ideal_field_in_api = hints.get("idealField", "IdealExpression") # Matches 'IdealExpression' from Genesys API

        # Get the identifier rule from the main rule_definition (this is rule['identifier'])
        entity_identifier_rule = rule_definition.get("identifier")
        if not entity_identifier_rule:
            logger.error(f"No 'identifier' found in rule definition for entity '{entity_name}'. Cannot filter API data.")
            return processed_api_data, max_id_from_this_api

        # Determine if this entity type needs complex processing (like skill expressions)
        # This is a heuristic based on the entity name from the rule.
        # A more robust method could be a flag in the rule, e.g., "apiDataStructure": "complex"
        is_complex_entity = "expression" in entity_name.lower() or \
                            "skill_expr" in entity_name.lower()
        if is_complex_entity:
            # For complex types, the identifier usually matches on the expression field from the API
            api_identifier_source_field = hints.get("apiIdentifierField", expression_field_in_api)


        # Process each item received from the API
        for api_item in raw_api_response_list:
            # Handle responses where data might be nested under a 'data' key, or be the item itself
            item_data = api_item.get('data', api_item)
            if not isinstance(item_data, dict): # Ensure item_data is a dictionary to call .get()
                logger.warning(f"Skipping API item for '{entity_name}' as item_data is not a dictionary: {item_data}")
                continue

            # --- Filter API item based on the rule's identifier ---
            value_to_match_in_api = item_data.get(api_identifier_source_field)
            if value_to_match_in_api is None:
                logger.debug(f"API item for '{entity_name}' missing identifier source field '{api_identifier_source_field}'. Skipping. Item: {item_data}")
                continue

            # Use the shared match_identifier_logic from utils.
            # The entity_identifier_rule is the 'identifier' object from the excelrule_template.json
            if not match_identifier_logic(str(value_to_match_in_api), entity_identifier_rule):
                logger.debug(f"API item for '{entity_name}' did not match rule identifier. Value checked: '{value_to_match_in_api}'. Rule: {entity_identifier_rule}. Item: {item_data}")
                continue
            # --- End API Item Filtering ---

            # If we reach here, the API item matches the rule's identifier. Now process it.
            item_id_val = item_data.get(id_field_in_api)
            if item_id_val is None: # ID is crucial for comparison and max ID calculation
                logger.warning(f"Skipping matched API item for '{entity_name}' due to missing ID (expected field: '{id_field_in_api}'): {item_data}")
                continue

            item_id_str = str(item_id_val)
            # Update max_id_from_this_api if current ID is numeric and larger
            if item_id_str.isdigit():
                try:
                    max_id_from_this_api = max(max_id_from_this_api, int(item_id_str))
                except ValueError: # Should not happen due to isdigit, but safety
                    logger.warning(f"Could not convert API ID '{item_id_str}' to int for max calculation (entity: {entity_name}).")


            # Structure the data for comparison_logic.py
            if is_complex_entity:
                # For skill expressions, we expect 'expression' and 'IdealExpression'
                expr_val = item_data.get(expression_field_in_api, "") or ""
                ideal_val = item_data.get(ideal_field_in_api, "") or ""

                # Normalize and create a key similar to how it's done by ExcelRuleEngine
                norm_expr = expr_val.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
                norm_ideal = ideal_val.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")

                # The comparison key for skill expressions is usually combined
                api_item_key = norm_expr
                if norm_ideal: # Only add ideal if it exists
                    api_item_key = f"{norm_expr} {norm_ideal}".strip()

                if not api_item_key: # Skip if key is empty after normalization
                    logger.warning(f"Skipping complex API item for '{entity_name}' (matched identifier) due to empty key after normalization: {item_data}")
                    continue

                # Store detailed dictionary for complex entities
                processed_api_data[api_item_key] = {
                    'id': item_id_str,
                    'expr': norm_expr,
                    'ideal': norm_ideal
                }
            else:
                # For simpler entities (VQs, Skills, VAGs), use name_field as key and ID as value
                item_name_val = item_data.get(name_field_in_api)
                if item_name_val is None:
                    logger.warning(f"Skipping simple API item for '{entity_name}' (matched identifier) due to missing name (expected field: '{name_field_in_api}'): {item_data}")
                    continue

                # Normalize the name field to match Excel processing (remove spaces, NBSP)
                api_item_key = str(item_name_val).replace(" ", "").replace('\u00A0', '')
                if not api_item_key: # Skip if key is empty after normalization
                    logger.warning(f"Skipping simple API item for '{entity_name}' (matched identifier) due to empty key after normalization: {item_data}")
                    continue
                # Store just the ID string for simpler entities
                processed_api_data[api_item_key] = item_id_str

        logger.info(f"Processed {len(processed_api_data)} matching items for entity '{entity_name}' from API. Max ID in this filtered response: {max_id_from_this_api}")

    except requests.exceptions.Timeout:
        logger.error(f"API request timed out for entity '{entity_name}' URL ({api_url}) after {timeout} seconds.")
    except requests.exceptions.RequestException as e:
        logger.error(f"API fetch failed for entity '{entity_name}' URL ({api_url}): {e}")
    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON for entity '{entity_name}' from URL ({api_url}): {e}")
    except Exception as e:
        # Catch any other unexpected errors during API processing
        logger.error(f"Unexpected error processing API data for entity '{entity_name}' from {api_url}: {e}", exc_info=True)

    return processed_api_data, max_id_from_this_api

