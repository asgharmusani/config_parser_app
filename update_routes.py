# -*- coding: utf-8 -*-
"""
Flask Blueprint for applying configuration templates to selected data rows
and simulating the preparation for a database update.
"""

import os
import json
import logging
import re # For placeholder replacement
from flask import Blueprint, request, jsonify, current_app, abort # Import current_app to access shared config

# --- Constants ---
TEMPLATE_DIR = './config_templates/' # Directory where JSON templates are stored
LOG_FILE_UI = 'ui_viewer.log'        # Assuming shared log file with main app
# Define the keys used to identify rows in different comparison sheets
# This helps retrieve the correct data from the cache
IDENTIFIER_KEYS = {
    "Skill_exprs Comparison": "Concatenated Key",
    "Vqs Comparison": "Item",
    "Skills Comparison": "Item",
    "Vags Comparison": "Item",
}


# --- Logging ---
# Use the root logger configured in the main app (ui_viewer.py)
logger = logging.getLogger()

# --- Blueprint Definition ---
# Create a Blueprint named 'updates'. The main app (ui_viewer.py) will register this.
update_bp = Blueprint('updates', __name__)

# --- Helper Function: Placeholder Replacement ---
def replace_placeholders(template_data: any, row_data: dict) -> any:
    """
    Recursively traverses a template structure (dict, list, or string)
    and replaces placeholders like {row.ColumnName} with values from row_data.

    Args:
        template_data: The template structure (can be dict, list, string, etc.).
        row_data: The dictionary containing data for the current row.

    Returns:
        The template structure with placeholders replaced.
    """
    # Regex to find placeholders like {row.ColumnName} or {row.Some Key}
    # It captures the part inside the curly braces after 'row.'
    placeholder_pattern = re.compile(r'{row\.([^}]+)}')

    # If the template_data is a string, perform replacement
    if isinstance(template_data, str):
        # Use a function with findall to handle multiple placeholders correctly
        def replace_match(match):
            col_name = match.group(1).strip() # Get the column name (e.g., "Concatenated Key")
            # Get value from row_data, default to empty string if not found or None
            replacement = row_data.get(col_name, "")
            return str(replacement) # Ensure replacement is a string

        # Substitute all found placeholders in the string
        return placeholder_pattern.sub(replace_match, template_data)

    # If it's a dictionary, recurse through its values
    elif isinstance(template_data, dict):
        return {
            key: replace_placeholders(value, row_data)
            for key, value in template_data.items()
        }

    # If it's a list, recurse through its items
    elif isinstance(template_data, list):
        return [replace_placeholders(item, row_data) for item in template_data]

    # Otherwise (numbers, booleans, None), return the value as is
    else:
        return template_data


# --- Backend Route for Applying Configuration ---
@update_bp.route('/api/apply-configuration', methods=['POST'])
def apply_configuration():
    """
    API endpoint to apply a selected template to selected row identifiers.
    It loads the template, retrieves row data (from cache), generates
    JSON payloads by replacing placeholders, and logs the results (simulating DB update).

    Expects JSON payload:
    {
        "templateName": "template_filename.json",
        "selectedRowsData": ["identifier1", "identifier2", ...] // List of unique identifiers
    }

    Returns:
        JSON response indicating success, partial success, or failure.
    """
    logger.info("Request received for /api/apply-configuration")
    try:
        # --- 1. Get and Validate Request Data ---
        request_data = request.get_json()
        if not request_data:
            logger.warning("Apply configuration request received with invalid/empty JSON payload.")
            return jsonify({"error": "Invalid JSON payload received."}), 400 # Bad Request

        template_name = request_data.get('templateName')
        # selected_row_identifiers is expected to be a list of strings (e.g., 'Item' or 'Concatenated Key' values)
        selected_row_identifiers = request_data.get('selectedRowsData')

        # Validate required fields
        if not template_name or selected_row_identifiers is None:
            logger.warning("Apply configuration request missing 'templateName' or 'selectedRowsData'.")
            return jsonify({"error": "Missing 'templateName' or 'selectedRowsData'."}), 400
        if not isinstance(selected_row_identifiers, list):
             logger.warning("'selectedRowsData' received is not a list.")
             return jsonify({"error": "'selectedRowsData' must be a list."}), 400

        logger.info(f"Attempting to apply template '{template_name}' to {len(selected_row_identifiers)} selected item(s).")

        # --- 2. Load the Specified Template ---
        # Basic security check on filename
        if '..' in template_name or template_name.startswith('/'):
            logger.error(f"Invalid template name requested: {template_name}")
            abort(400, description="Invalid template name.") # Bad Request

        template_path = os.path.join(TEMPLATE_DIR, template_name)
        if not os.path.exists(template_path) or not os.path.isfile(template_path):
            logger.error(f"Template file not found at path: {template_path}")
            return jsonify({"error": f"Template '{template_name}' not found."}), 404 # Not Found

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_json = json.load(f) # Load the template structure
            logger.debug(f"Successfully loaded template '{template_name}'.")
        except Exception as e:
            logger.error(f"Error reading or parsing template file {template_name}: {e}", exc_info=True)
            return jsonify({"error": f"Failed to load or parse template '{template_name}'."}), 500 # Server Error

        # --- 3. Retrieve Data for Selected Rows (from cached Excel data) ---
        # This simulates fetching data based on identifiers.
        # In a real app, you might query a database here using the identifiers.
        all_excel_data = current_app.config.get('EXCEL_DATA', {}) # Get data cached by ui_viewer.py
        rows_to_process = []
        processed_identifiers = set() # Track identifiers found to avoid duplicates
        identifier_key_map = {} # Stores which key ('Item' or 'Concatenated Key') matched the identifier

        # Iterate through all comparison sheets in the cached data
        for sheet_name, sheet_data in all_excel_data.items():
            # Determine the primary identifier key for this sheet type
            id_key = IDENTIFIER_KEYS.get(sheet_name)
            if not id_key:
                logging.warning(f"No identifier key defined for sheet '{sheet_name}', skipping.")
                continue

            # Find rows matching the selected identifiers in this sheet
            for row in sheet_data:
                row_identifier = row.get(id_key)
                # Check if this row's identifier is in the requested list and not already processed
                if row_identifier in selected_row_identifiers and row_identifier not in processed_identifiers:
                    rows_to_process.append(row) # Add the full row dictionary
                    processed_identifiers.add(row_identifier) # Mark as processed
                    identifier_key_map[row_identifier] = id_key # Remember which key matched

        found_count = len(rows_to_process)
        missing_identifiers = set(selected_row_identifiers) - processed_identifiers
        missing_count = len(missing_identifiers)
        logger.info(f"Retrieved data for {found_count} of {len(selected_row_identifiers)} selected identifiers from cached data.")
        if missing_count > 0:
             logger.warning(f"Could not find cached data for identifiers: {missing_identifiers}")


        # --- 4. Generate Payloads using Template and Row Data ---
        final_payloads = []
        processing_errors = []
        for row_data in rows_to_process:
            row_id_for_log = row_data.get(identifier_key_map.get(row_data.get('Item', row_data.get('Concatenated Key'))), 'UNKNOWN_ID')
            try:
                # Use the helper function to replace placeholders in the template
                generated_payload = replace_placeholders(template_json, row_data)
                final_payloads.append(generated_payload)
                logging.debug(f"Generated payload for row identifier '{row_id_for_log}'.") # Avoid logging full payload at debug level
            except Exception as e:
                # Catch errors during placeholder replacement for a specific row
                logger.error(f"Error processing template for row identifier '{row_id_for_log}': {e}", exc_info=True)
                processing_errors.append(f"Error processing row '{row_id_for_log}': {e}")

        # --- 5. SIMULATE Database Update ---
        # In a real application, you would now send the `final_payloads` list
        # to your actual database update API endpoint or service.
        logger.info(f"--- SIMULATING DATABASE UPDATE (START) ---")
        logger.info(f"Template: '{template_name}'")
        logger.info(f"Generated {len(final_payloads)} payloads to be sent for update.")
        if final_payloads:
             # Log the generated payloads (or a summary) for verification.
             # Be cautious about logging sensitive data in production environments.
             log_limit = 5 # Limit how many full payloads are logged
             logger.info(f"Example Payloads (limit {log_limit}):")
             for i, payload in enumerate(final_payloads[:log_limit]):
                 try:
                     # Log payload as pretty-printed JSON string
                     logger.info(f"Payload {i+1}: {json.dumps(payload, indent=2)}")
                 except Exception as json_e:
                     logger.error(f"Error converting payload {i+1} to JSON for logging: {json_e}")
                     logger.info(f"Payload {i+1} (raw): {payload}") # Log raw dict as fallback

             if len(final_payloads) > log_limit:
                 logger.info(f"... and {len(final_payloads) - log_limit} more payloads generated but not logged in detail.")
        # --- SIMULATION END --- You would add your actual DB update call here ---


        # --- 6. Construct and Return Response ---
        response_status_code = 200 # Default OK
        response_data = {
            "message": f"Processed {found_count} of {len(selected_row_identifiers)} selected items using template '{template_name}'. Payloads generated and logged (simulation).",
            "status": "Success",
            "processed_count": found_count,
            "payload_count": len(final_payloads),
            "errors": []
        }

        if missing_count > 0:
             response_data["message"] += f" Could not find data for {missing_count} identifiers: {list(missing_identifiers)}."
             response_data["status"] = "Partial Success / Missing Data"
             response_status_code = 207 # Multi-Status
             logger.warning(response_data["message"])

        if processing_errors:
             response_data["errors"] = processing_errors
             response_data["message"] += f" Encountered {len(processing_errors)} errors during payload generation."
             # If there were already missing items, keep 207, otherwise use 207 for processing errors too
             response_data["status"] = "Partial Success / Errors" if response_status_code == 200 else response_data["status"]
             response_status_code = 207 # Multi-Status
             logger.error(f"Processing errors occurred: {processing_errors}")


        return jsonify(response_data), response_status_code

    except Exception as e:
        # Catch any other unexpected errors during the process
        logger.error(f"Unexpected error in /api/apply-configuration: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred."}), 500 # Internal Server Error

