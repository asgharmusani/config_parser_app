# -*- coding: utf-8 -*-
"""
Flask Blueprint for applying configuration templates to selected data rows
and simulating the preparation for a database update.
"""

import os
import json
import logging
import re
from flask import Blueprint, request, jsonify, current_app

# --- Constants ---
TEMPLATE_DIR = './config_templates/'
LOG_FILE_UI = 'ui_viewer.log' # Assuming shared log file

# --- Logging ---
logger = logging.getLogger()

# --- Blueprint Definition ---
update_bp = Blueprint('updates', __name__)

# --- Placeholder/Helper ---
def replace_placeholders(template_dict: dict, row_data: dict) -> dict:
    """Recursively replaces {row.ColumnName} placeholders in a template dict."""
    output_dict = {}
    placeholder_pattern = re.compile(r'{row\.([^}]+)}') # Matches {row.ColumnName}

    for key, value in template_dict.items():
        if isinstance(value, str):
            # Find all placeholders in the string value
            new_value = value
            for match in placeholder_pattern.finditer(value):
                placeholder = match.group(0) # e.g., {row.ColumnName}
                col_name = match.group(1).strip() # e.g., ColumnName
                # Get value from row_data (case-insensitive fallback might be useful)
                replacement = row_data.get(col_name)
                # Handle None values - replace with empty string or null? Let's use empty string for simplicity
                if replacement is None:
                    replacement = ""
                # Replace only the placeholder part - important if string has mixed content
                # Convert replacement to string as we are replacing in a string context
                new_value = new_value.replace(placeholder, str(replacement))
            output_dict[key] = new_value
        elif isinstance(value, dict):
            output_dict[key] = replace_placeholders(value, row_data) # Recurse for nested dicts
        elif isinstance(value, list):
             # Process lists (optional - handle dicts within lists if needed)
             # For now, just copy lists as is or handle simple string replacements
             output_dict[key] = [
                 replace_placeholders(item, row_data) if isinstance(item, dict)
                 else (item.replace(placeholder, str(row_data.get(match.group(1).strip(), ""))) if isinstance(item, str) and placeholder_pattern.search(item) else item)
                 for item in value
                 for match in placeholder_pattern.finditer(str(item)) # Basic replacement in list strings
                 for placeholder in [match.group(0)]
             ] if value else [] # Handle empty lists
             # This list handling is basic, might need refinement based on expected list content
        else:
            output_dict[key] = value # Copy numbers, booleans, etc., directly
    return output_dict


# --- Backend Route ---
@update_bp.route('/api/apply-configuration', methods=['POST'])
def apply_configuration():
    """
    Applies a selected template to selected row identifiers, generates JSON payloads.
    SIMULATES database update by logging payloads.
    """
    try:
        request_data = request.get_json()
        if not request_data:
            return jsonify({"error": "Invalid JSON payload received."}), 400

        template_name = request_data.get('templateName')
        # selected_rows_data should contain the actual data or identifiers
        # Let's assume it contains identifiers (e.g., the unique key/item name)
        selected_row_identifiers = request_data.get('selectedRowsData') # Expecting list of identifiers

        if not template_name or selected_row_identifiers is None: # Check for None explicitly for list
            return jsonify({"error": "Missing 'templateName' or 'selectedRowsData'."}), 400
        if not isinstance(selected_row_identifiers, list):
             return jsonify({"error": "'selectedRowsData' must be a list."}), 400

        logger.info(f"Received request to apply template '{template_name}' to {len(selected_row_identifiers)} items.")

        # --- 1. Load the Template ---
        if '..' in template_name or template_name.startswith('/'):
            abort(400, description="Invalid template name.")
        template_path = os.path.join(TEMPLATE_DIR, template_name)
        if not os.path.exists(template_path):
            logger.error(f"Template file not found: {template_path}")
            return jsonify({"error": f"Template '{template_name}' not found."}), 404

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_json = json.load(f)
            logger.debug(f"Loaded template '{template_name}'.")
        except Exception as e:
            logger.error(f"Error reading or parsing template {template_name}: {e}", exc_info=True)
            return jsonify({"error": f"Failed to load or parse template '{template_name}'."}), 500

        # --- 2. Get Data for Selected Rows (from cache/Excel data) ---
        # This part needs access to the data loaded by the main app.
        # We assume the main app stores it in current_app.config['EXCEL_DATA']
        # The structure is {'Sheet Name': [{'Item':.., 'ID':.., 'Status':..}, ...]}
        # We need to find the rows matching the identifiers across *all* sheets in the cache.
        # This simulation assumes identifiers are unique across comparison types.
        # A real app would likely query a database using these identifiers.

        all_excel_data = current_app.config.get('EXCEL_DATA', {})
        rows_to_process = []
        processed_identifiers = set() # Avoid duplicates if ID appears in multiple sheets
        identifier_key_map = {} # Map identifier back to its primary key name ('Item' or 'Concatenated Key')

        for sheet_name, sheet_data in all_excel_data.items():
            # Determine the identifier key for this sheet type
            is_skill_expr = (sheet_name == "Skill_exprs Comparison")
            id_key = "Concatenated Key" if is_skill_expr else "Item"

            for row in sheet_data:
                row_identifier = row.get(id_key)
                if row_identifier in selected_row_identifiers and row_identifier not in processed_identifiers:
                    rows_to_process.append(row) # Add the whole row dict
                    processed_identifiers.add(row_identifier)
                    identifier_key_map[row_identifier] = id_key # Store which key was used

        found_count = len(rows_to_process)
        missing_count = len(selected_row_identifiers) - found_count
        logger.info(f"Found data for {found_count} selected identifiers. Missing data for {missing_count} identifiers.")
        if missing_count > 0:
             logger.warning(f"Could not find data for identifiers: {set(selected_row_identifiers) - processed_identifiers}")


        # --- 3. Generate Payloads ---
        final_payloads = []
        processing_errors = []
        for row_data in rows_to_process:
            try:
                # Generate the payload using the template and row data
                generated_payload = replace_placeholders(template_json, row_data)
                final_payloads.append(generated_payload)
                logger.debug(f"Generated payload for row identifier '{row_data.get(identifier_key_map.get(row_data.get('Item', row_data.get('Concatenated Key'))), 'UNKNOWN')}': {generated_payload}")
            except Exception as e:
                row_id = row_data.get(identifier_key_map.get(row_data.get('Item', row_data.get('Concatenated Key'))), 'UNKNOWN')
                logger.error(f"Error processing template for row identifier '{row_id}': {e}", exc_info=True)
                processing_errors.append(f"Error processing row '{row_id}': {e}")

        # --- 4. SIMULATE Database Update ---
        # In a real application, you would send `final_payloads` to your database update API/service here.
        logger.info(f"--- SIMULATING DATABASE UPDATE ---")
        logger.info(f"Generated {len(final_payloads)} payloads to send for update using template '{template_name}'.")
        if final_payloads:
             # Log first few payloads for inspection (be careful with sensitive data in real logs)
             log_limit = 3
             logger.info(f"Example Payloads (limit {log_limit}):")
             for i, payload in enumerate(final_payloads[:log_limit]):
                 logger.info(f"Payload {i+1}: {json.dumps(payload)}") # Log generated JSON
             if len(final_payloads) > log_limit:
                 logger.info(f"... and {len(final_payloads) - log_limit} more payloads.")
        logger.info(f"--- END SIMULATION ---")


        # --- 5. Return Response ---
        response_message = f"Processed {found_count} of {len(selected_row_identifiers)} selected items using template '{template_name}'. Payloads generated and logged (simulation)."
        if missing_count > 0:
             response_message += f" Could not find data for {missing_count} items."
        if processing_errors:
             response_message += f" Encountered {len(processing_errors)} errors during processing."
             return jsonify({
                 "message": response_message,
                 "status": "Partial Success / Errors",
                 "processed_count": found_count,
                 "payload_count": len(final_payloads),
                 "errors": processing_errors
             }), 207 # Multi-Status
        else:
             return jsonify({
                 "message": response_message,
                 "status": "Success",
                 "processed_count": found_count,
                 "payload_count": len(final_payloads)
             }), 200

    except Exception as e:
        logger.error(f"Unexpected error in /api/apply-configuration: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred."}), 500
