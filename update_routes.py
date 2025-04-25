# -*- coding: utf-8 -*-
"""
Flask Blueprint for applying configuration templates to selected data rows.

Handles simulating the template application and confirming the update (simulated).
Uses separate Max IDs (DN vs Agent Group) calculated from loaded data
for the {func.next_id} placeholder.
"""

import os
import json
import logging
import re # For placeholder replacement
from flask import Blueprint, request, jsonify, current_app, abort # Import current_app to access shared config
# Import the placeholder replacement function from template_routes
try:
    # This assumes template_routes.py is in the same directory or accessible via Python path
    from template_routes import replace_placeholders, TEMPLATE_DIR
except ImportError:
    # Fallback if running standalone or import issue (should not happen in normal execution)
    logging.error("Could not import 'replace_placeholders' or 'TEMPLATE_DIR' from template_routes.py")
    # Define a dummy function or exit if critical
    def replace_placeholders(template_data, row_data, current_row_next_id=None):
        """Dummy placeholder function if import fails."""
        return template_data
    TEMPLATE_DIR = './config_templates/' # Define fallback constant


# --- Constants ---
LOG_FILE_UI = 'ui_viewer.log' # Assuming shared log file
# Keys used to identify rows in different comparison sheets (must match ui_viewer.py)
# This mapping helps retrieve the correct identifier from row data.
IDENTIFIER_KEYS = {
    "Skill_exprs Comparison": "Concatenated Key",
    "Vqs Comparison": "Item",
    "Skills Comparison": "Item",
    "Vags Comparison": "Item",
}
# Define which comparison sheets belong to which ID group
DN_SHEETS = {"Vqs Comparison"}
AGENT_GROUP_SHEETS = {"Skills Comparison", "Vags Comparison", "Skill_exprs Comparison"}


# --- Logging ---
# Use the root logger configured in the main app (ui_viewer.py)
logger = logging.getLogger()

# --- Blueprint Definition ---
# Create a Blueprint named 'updates'. The main app (ui_viewer.py) will register this.
update_bp = Blueprint('updates', __name__)

# --- ID Generation Helper (MODIFIED for separate counters) ---
class IdGenerator:
    """
    Generates sequential IDs within a single request context, maintaining
    separate sequences for DN (VQ) and Agent Group entities based on
    Max IDs read from the application config.
    """
    def __init__(self):
        """Initializes the generator by fetching Max IDs from app config."""
        # Fetch the Max IDs calculated and stored by ui_viewer.py
        # Use fallbacks if the values aren't found in config
        max_dn_id = current_app.config.get('MAX_DN_ID', 0) # Default to 0 if not found
        max_ag_id = current_app.config.get('MAX_AG_ID', 0) # Default to 0 if not found

        # The next ID should be one greater than the maximum found for each type
        self._next_dn_id = max_dn_id + 1
        self._next_ag_id = max_ag_id + 1

        logger.info(f"ID Generator initialized. Next DN ID: {self._next_dn_id} (based on max: {max_dn_id}), Next AG ID: {self._next_ag_id} (based on max: {max_ag_id})")

    def get_next_dn_id(self) -> int:
        """Returns the next sequential DN ID and increments the DN counter."""
        next_id = self._next_dn_id
        self._next_dn_id += 1
        logger.debug(f"Generated next DN ID: {next_id}")
        return next_id

    def get_next_ag_id(self) -> int:
        """Returns the next sequential Agent Group ID and increments the AG counter."""
        next_id = self._next_ag_id
        self._next_ag_id += 1
        logger.debug(f"Generated next AG ID: {next_id}")
        return next_id

# --- Backend Route for SIMULATING Configuration Application (MODIFIED) ---
@update_bp.route('/api/simulate-configuration', methods=['POST'])
def simulate_configuration():
    """
    API endpoint to apply a selected template to selected row identifiers,
    generate JSON payloads using placeholders (including func.next_id per row,
    using appropriate ID sequence), and RETURN the generated payloads for review.

    Expects JSON payload:
    {
        "templateName": "template_filename.json",
        "selectedRowsData": ["identifier1", "identifier2", ...] // List of unique identifiers
    }

    Returns:
        JSON response containing the list of generated payloads and status.
    """
    logger.info("Request received for /api/simulate-configuration")
    try:
        # --- 1. Get and Validate Request Data ---
        request_data = request.get_json()
        if not request_data:
            logger.warning("Simulate config request: Invalid/empty JSON payload.")
            return jsonify({"error": "Invalid JSON payload received."}), 400

        template_name = request_data.get('templateName')
        selected_row_identifiers = request_data.get('selectedRowsData')

        if not template_name or selected_row_identifiers is None:
            logger.warning("Simulate config request missing 'templateName' or 'selectedRowsData'.")
            return jsonify({"error": "Missing 'templateName' or 'selectedRowsData'."}), 400
        if not isinstance(selected_row_identifiers, list):
             logger.warning("'selectedRowsData' is not a list.")
             return jsonify({"error": "'selectedRowsData' must be a list."}), 400

        logger.info(f"Simulating template '{template_name}' for {len(selected_row_identifiers)} items.")

        # --- 2. Load the Specified Template ---
        if '..' in template_name or template_name.startswith('/'):
            logger.error(f"Invalid template name requested: {template_name}")
            abort(400, description="Invalid template name.")

        template_path = os.path.join(TEMPLATE_DIR, template_name)
        if not os.path.exists(template_path) or not os.path.isfile(template_path):
            logger.error(f"Template file not found: {template_path}")
            return jsonify({"error": f"Template '{template_name}' not found."}), 404

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_json = json.load(f)
            logger.debug(f"Loaded template '{template_name}'.")
        except Exception as e:
            logger.error(f"Error reading/parsing template {template_name}: {e}", exc_info=True)
            return jsonify({"error": f"Failed to load/parse template '{template_name}'."}), 500

        # --- 3. Retrieve Data for Selected Rows (from cache) ---
        all_excel_data = current_app.config.get('EXCEL_DATA', {})
        rows_to_process = [] # Will store tuples of (row_data, entity_type)
        processed_identifiers = set()
        identifier_key_map = {}

        # Iterate through sheets to find matching rows and determine their type
        for sheet_name, sheet_data in all_excel_data.items():
            id_key = IDENTIFIER_KEYS.get(sheet_name)
            if not id_key: continue

            # Determine entity type based on sheet name
            entity_type = None
            if sheet_name in DN_SHEETS:
                entity_type = 'dn'
            elif sheet_name in AGENT_GROUP_SHEETS:
                entity_type = 'agent_group'
            else:
                logger.warning(f"Sheet '{sheet_name}' not mapped to DN or AG type. Cannot determine ID sequence.")
                continue # Skip rows from this sheet for ID generation

            for row in sheet_data:
                row_identifier = row.get(id_key)
                if row_identifier in selected_row_identifiers and row_identifier not in processed_identifiers:
                    rows_to_process.append((row, entity_type)) # Store row data AND its type
                    processed_identifiers.add(row_identifier)
                    identifier_key_map[row_identifier] = id_key

        found_count = len(rows_to_process)
        missing_identifiers = set(selected_row_identifiers) - processed_identifiers
        missing_count = len(missing_identifiers)
        logger.info(f"Retrieved data for {found_count} of {len(selected_row_identifiers)} identifiers from cache.")
        if missing_count > 0:
             logger.warning(f"Could not find cached data for identifiers: {missing_identifiers}")

        # --- 4. Generate Payloads ---
        generated_payloads = []
        processing_errors = []
        id_generator = IdGenerator() # Initialize ID generator for this batch

        # Process each row found
        for row_data, entity_type in rows_to_process:
            # Determine identifier for logging
            row_id_for_log = "UNKNOWN_ID"
            id_key_for_row = identifier_key_map.get(row_data.get("Item", row_data.get("Concatenated Key")))
            if id_key_for_row:
                row_id_for_log = row_data.get(id_key_for_row, "UNKNOWN_ID")

            try:
                # --- MODIFICATION: Generate ID based on type BEFORE replacing placeholders ---
                current_row_id = None
                if entity_type == 'dn':
                    current_row_id = id_generator.get_next_dn_id()
                elif entity_type == 'agent_group':
                    current_row_id = id_generator.get_next_ag_id()
                else:
                    # Should not happen if sheet was skipped earlier, but handle defensively
                    logger.warning(f"Cannot generate ID for row '{row_id_for_log}' - unknown entity type.")
                    # Placeholder {func.next_id} will result in an error string if used

                logger.debug(f"Using next_id={current_row_id} (type: {entity_type}) for row '{row_id_for_log}'")

                # Generate the payload, passing the specific ID for this row
                generated_payload = replace_placeholders(
                    template_data=template_json,
                    row_data=row_data,
                    current_row_next_id=current_row_id # Pass the generated ID
                )
                # --- END MODIFICATION ---

                generated_payloads.append(generated_payload)
                logging.debug(f"Generated simulation payload for row identifier '{row_id_for_log}'.")
            except Exception as e:
                logger.error(f"Error processing template for row identifier '{row_id_for_log}': {e}", exc_info=True)
                processing_errors.append(f"Error processing row '{row_id_for_log}': {e}")

        # --- 5. Construct and Return Simulation Response ---
        response_status_code = 200
        response_data = {
            "message": f"Simulation complete for template '{template_name}'. Generated {len(generated_payloads)} payloads for review.",
            "status": "Simulation Success",
            "processed_count": found_count,
            "payloads": generated_payloads, # Include generated payloads
            "errors": []
        }

        if missing_count > 0:
             response_data["message"] += f" Could not find data for {missing_count} identifiers: {list(missing_identifiers)}."
             response_data["status"] = "Simulation Partial Success / Missing Data"
             response_status_code = 207
             logger.warning(response_data["message"])

        if processing_errors:
             response_data["errors"] = [str(e) for e in processing_errors]
             response_data["message"] += f" Encountered {len(processing_errors)} errors during payload generation."
             response_data["status"] = "Simulation Partial Success / Errors" if response_status_code == 200 else response_data["status"]
             response_status_code = 207
             logger.error(f"Simulation processing errors occurred: {processing_errors}")

        logger.info(f"Simulation successful. Returning {len(generated_payloads)} generated payloads.")
        return jsonify(response_data), response_status_code

    except Exception as e:
        logger.error(f"Unexpected error in /api/simulate-configuration: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred during simulation."}), 500


# --- Backend Route for CONFIRMING the Update ---
@update_bp.route('/api/confirm-update', methods=['POST'])
def confirm_update():
    """
    API endpoint to receive previously generated payloads (from simulation)
    and perform the final (simulated) database update action by logging them.

    Expects JSON payload:
    {
        "payloads": [ { ... payload 1 ... }, { ... payload 2 ... } ]
    }

    Returns:
        JSON response indicating success or failure of the simulated update.
    """
    logger.info("Request received for /api/confirm-update")
    try:
        # --- 1. Get and Validate Request Data ---
        request_data = request.get_json()
        if not request_data:
            logger.warning("Confirm update request received with invalid/empty JSON payload.")
            return jsonify({"error": "Invalid JSON payload received."}), 400

        payloads_to_commit = request_data.get('payloads')

        if payloads_to_commit is None or not isinstance(payloads_to_commit, list):
            logger.warning("Confirm update request missing 'payloads' list or it's not a list.")
            return jsonify({"error": "Missing 'payloads' list in request or invalid format."}), 400

        logger.info(f"Received {len(payloads_to_commit)} payloads for final (simulated) update.")

        # --- 2. SIMULATE Database Update ---
        # In a real application, iterate through payloads_to_commit and send
        # each one to the appropriate database update service/API.
        commit_errors = []
        commit_success_count = 0

        logger.info(f"--- SIMULATING FINAL DATABASE UPDATE (START) ---")
        for i, payload in enumerate(payloads_to_commit):
            try:
                # Simulate sending payload to DB - just log it here
                logger.info(f"Simulating DB Update for Payload {i+1}: {json.dumps(payload, indent=2)}")
                # --- Replace log statement with actual DB call ---
                # db_response = your_db_update_function(payload)
                # if not db_response.success:
                #     commit_errors.append(f"Payload {i+1}: {db_response.error_message}")
                # else:
                #     commit_success_count += 1
                # -------------------------------------------------
                commit_success_count += 1 # Increment success count for simulation
            except Exception as db_err:
                logger.error(f"Simulated DB update FAILED for Payload {i+1}: {db_err}", exc_info=True)
                commit_errors.append(f"Payload {i+1}: {db_err}")
        logger.info(f"--- SIMULATING FINAL DATABASE UPDATE (END) ---")


        # --- 3. Construct and Return Response ---
        response_status_code = 200 # Default OK
        response_data = {
            "message": f"Simulated update completed for {commit_success_count} of {len(payloads_to_commit)} payloads.",
            "status": "Update Simulation Success",
            "success_count": commit_success_count,
            "error_count": len(commit_errors),
            "errors": [str(e) for e in commit_errors] # Convert errors to strings
        }

        if commit_errors:
            response_data["status"] = "Update Simulation Partial Success / Errors"
            response_status_code = 207 # Multi-Status

        if commit_success_count == 0 and len(payloads_to_commit) > 0:
             response_data["status"] = "Update Simulation Failed"
             response_status_code = 500 # Internal Server Error

        return jsonify(response_data), response_status_code

    except Exception as e:
        logger.error(f"Unexpected error in /api/confirm-update: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred during update confirmation."}), 500

