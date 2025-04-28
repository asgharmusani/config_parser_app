# -*- coding: utf-8 -*-
"""
Handles fetching data from external APIs (DN and Agent Group).

Core Function:
- fetch_api_data: Calls configured API endpoints, processes the responses,
                  calculates separate maximum numeric IDs for DN and Agent Group
                  entities, and returns the structured data and max IDs.
"""

import logging
import requests
import re # Used for simple validation
from typing import Dict, Any, Tuple

logger = logging.getLogger(__name__) # Use module-specific logger

# --- API Data Fetching ---
def fetch_api_data(config: Dict[str, Any]) -> Tuple[Dict[str, Dict[str, Any]], int, int]:
    """
    Fetches routing entity data from APIs specified in the configuration.
    Calculates the maximum numeric ID found SEPARATELY for DN and Agent Group sources.
    Stores detailed info for skill expressions.

    Args:
        config: The loaded application configuration dictionary, containing keys like
                'dn_url', 'agent_group_url', and 'api_timeout'.

    Returns:
        A tuple containing:
        - api_data: Dictionary containing the fetched API data. Structure varies by key:
                    {
                        "vqs": {norm_name: id_str},
                        "skills": {norm_name: id_str},
                        "vags": {norm_name: id_str},
                        "skill_exprs": {concat_key: {"id": id_str, "expr": expr_str, "ideal": ideal_str}}
                    }
        - max_dn_id: The highest numeric ID found in the DN API data (or 0).
        - max_ag_id: The highest numeric ID found in the Agent Group API data (or 0).
    """
    logging.info("Fetching API data...")
    # Get API details from config, using .get() with defaults for safety
    dn_url = config.get('dn_url')
    agent_group_url = config.get('agent_group_url')
    timeout = config.get('api_timeout', 15) # Default timeout if not specified

    # Validate URLs exist
    if not dn_url or not agent_group_url:
        msg = "API URLs ('dn_url', 'agent_group_url') not found in configuration."
        logger.error(msg)
        print(f"ERROR: {msg}")
        # Return empty data and zero IDs if URLs are missing
        return {"vqs": {}, "skills": {}, "skill_exprs": {}, "vags": {}}, 0, 0

    # Initialize data structures and max ID counters
    api_data = {"vqs": {}, "skills": {}, "skill_exprs": {}, "vags": {}}
    max_dn_id = 0
    max_ag_id = 0

    # Initialize JSON variables in case requests fail
    dn_json = []
    ag_json = []

    # --- Fetch DN (VQ) data ---
    try:
        logging.debug(f"Fetching DN data from {dn_url} with timeout={timeout}s")
        dn_response = requests.get(dn_url, timeout=timeout)
        # Raise an exception for bad status codes (4xx or 5xx)
        dn_response.raise_for_status()
        # Parse the JSON response
        dn_json = dn_response.json()
        logging.info(f"Successfully fetched DN response ({len(dn_json)} items).")
    except requests.exceptions.Timeout:
         logging.error(f"API request timed out for DN URL ({dn_url}) after {timeout} seconds.")
         print(f"ERROR: API request timed out for DN URL. Check URL and network.")
         # Continue to try fetching AG data, max_dn_id remains 0
    except requests.exceptions.RequestException as e:
        logging.error(f"API fetch failed for DN URL ({dn_url}): {e}")
        print(f"ERROR: Failed to fetch data from DN API. Check URL and network. Details in log.")
        # Continue to try fetching AG data, max_dn_id remains 0
    except json.JSONDecodeError as e:
        logging.error(f"Failed to decode JSON response from DN URL ({dn_url}): {e}")
        print(f"ERROR: Invalid JSON received from DN API. Check API endpoint.")
        # Continue to try fetching AG data

    # --- Fetch Agent Group data ---
    try:
        logging.debug(f"Fetching Agent Group data from {agent_group_url} with timeout={timeout}s")
        ag_response = requests.get(agent_group_url, timeout=timeout)
        ag_response.raise_for_status()
        ag_json = ag_response.json()
        logging.info(f"Successfully fetched Agent Group response ({len(ag_json)} items).")
    except requests.exceptions.Timeout:
         logging.error(f"API request timed out for Agent Group URL ({agent_group_url}) after {timeout} seconds.")
         print(f"ERROR: API request timed out for Agent Group URL. Check URL and network.")
         # Return current api_data and calculated max IDs (max_ag_id will be 0)
         return api_data, max_dn_id, max_ag_id
    except requests.exceptions.RequestException as e:
        logging.error(f"API fetch failed for Agent Group URL ({agent_group_url}): {e}")
        print(f"ERROR: Failed to fetch data from Agent Group API. Check URL and network. Details in log.")
         # Return current api_data and calculated max IDs (max_ag_id will be 0)
        return api_data, max_dn_id, max_ag_id
    except json.JSONDecodeError as e:
        logging.error(f"Failed to decode JSON response from Agent Group URL ({agent_group_url}): {e}")
        print(f"ERROR: Invalid JSON received from Agent Group API. Check API endpoint.")
        # Return current api_data and calculated max IDs (max_ag_id will be 0)
        return api_data, max_dn_id, max_ag_id


    # --- Process DN (VQ) data ---
    vq_count = 0
    for item in dn_json: # Iterate through the fetched list
        # Safely get nested data, defaulting to empty dict if 'data' key is missing
        data = item.get('data', {})
        vq_name = data.get('name')
        vq_id = data.get('id') # API might return int or string ID

        # Ensure both name and ID are present and not None
        if vq_name and vq_id is not None:
            # Normalize name: remove spaces and non-breaking spaces (\u00A0)
            normalized_vq = vq_name.replace(" ", "").replace('\u00A0', '')
            id_str = str(vq_id) # Store ID as string for consistency

            # Store in api_data dictionary
            api_data["vqs"][normalized_vq] = id_str

            # --- Calculate Max DN ID ---
            # Check if the ID is numeric and update max_dn_id if it's larger
            if id_str.isdigit():
                try:
                    max_dn_id = max(max_dn_id, int(id_str)) # Update DN max ID
                except ValueError: # Should not happen due to isdigit, but safety check
                    logging.warning(f"Could not convert DN ID '{id_str}' to int for max calculation.")
            # --- End Calculate Max DN ID ---

            logging.debug(f"Processed VQ: Name='{normalized_vq}', ID='{id_str}'")
            vq_count += 1
        else:
            # Log if essential data is missing for an item
            logging.warning(f"Skipping DN item due to missing name or id: {item}")
    logging.info(f"Processed {vq_count} VQs from API. Max DN ID found: {max_dn_id}")


    # --- Process Agent Group (Skill, Skill Expr, VAG) data ---
    skill_count, expr_count, vag_count = 0, 0, 0
    skipped_ag_count = 0
    for item in ag_json: # Iterate through the fetched list
        data = item.get('data', {})
        ag_id = data.get('id')
        expression = data.get('expression', '') or '' # Ensure string, default to empty
        ideal_expression = data.get('IdealExpression', '') or '' # Ensure string, default to empty

        # Skip if ID is missing
        if ag_id is None:
            logging.warning(f"Skipping AG item due to missing id: {item}")
            skipped_ag_count += 1
            continue
        ag_id_str = str(ag_id) # Store ID as string

        # --- Calculate Max AG ID ---
        # Check if the ID is numeric and update max_ag_id if it's larger
        if ag_id_str.isdigit():
            try:
                max_ag_id = max(max_ag_id, int(ag_id_str)) # Update AG max ID
            except ValueError:
                 logging.warning(f"Could not convert AG ID '{ag_id_str}' to int for max calculation.")
        # --- End Calculate Max AG ID ---

        # Normalize expressions: remove spaces, add spaces around operators for consistency
        # This helps match sheet data which is similarly normalized
        norm_expr = expression.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")
        norm_ideal = ideal_expression.replace(" ", "").replace('\u00A0', '').replace("|", " | ").replace("&", " & ")

        # Create the concatenated key used for comparison, mirroring sheet processing
        combined_expr = norm_expr
        is_skill_expr = ">" in norm_expr
        if is_skill_expr and norm_ideal: # Only combine if it's a skill expression AND has an ideal part
             combined_expr = f"{norm_expr} {norm_ideal}".strip()

        # Categorize and store based on expression content
        if is_skill_expr: # Skill Expression (contains '>')
            # Store dict with details needed for comparison sheet
            api_data["skill_exprs"][combined_expr] = {
                'id': ag_id_str,
                'expr': norm_expr,   # Store the base expression from API
                'ideal': norm_ideal # Store the ideal part from API
            }
            logging.debug(f"Processed Skill Expr: Key='{combined_expr}', ID='{ag_id_str}'")
            expr_count += 1
        elif "VAG_" in norm_expr: # VAG (check for prefix)
            api_data["vags"][norm_expr] = ag_id_str # Store only ID for VAGs
            logging.debug(f"Processed VAG: Name='{norm_expr}', ID='{ag_id_str}'")
            vag_count += 1
        elif norm_expr: # Potentially a Simple Skill (if not VAG or Skill Expr)
             # Check it contains actual characters, not just operators/spaces
             if re.search(r'[a-zA-Z0-9]', norm_expr):
                 api_data["skills"][norm_expr] = ag_id_str # Store only ID for Skills
                 logging.debug(f"Processed Skill: Name='{norm_expr}', ID='{ag_id_str}'")
                 skill_count += 1
             else:
                 # Log if expression becomes empty/invalid after normalization
                 logging.warning(f"Skipping AG item - skill name seems empty or invalid after normalization: {item}")
                 skipped_ag_count += 1
        else:
             # Log if original expression was empty
             logging.warning(f"Skipping AG item with empty expression: {item}")
             skipped_ag_count += 1

    logging.info(f"Processed Agent Groups from API: Skills={skill_count}, SkillExprs={expr_count}, VAGs={vag_count}. Skipped={skipped_ag_count}.")
    logging.info(f"Finished parsing API data. Max Agent Group ID found: {max_ag_id}")

    # Return the structured data and BOTH max IDs
    return api_data, max_dn_id, max_ag_id
