# -*- coding: utf-8 -*-
"""
Flask Blueprint for handling backend processing tasks and API endpoints.

Includes API endpoints for:
- Uploading and processing the Excel file (triggering comparison).
- Simulating configuration template application.
- Confirming and finalizing (simulated) updates.
- Updating the application configuration.
"""

import os
import json
import logging
import re
import shutil # For file operations if needed
import openpyxl # Needed for loading workbook in processing route
from flask import (
    Blueprint, request, jsonify, current_app, abort, flash, redirect, url_for
)
from werkzeug.utils import secure_filename # For secure file uploads
# --- MODIFICATION START: Import Font ---
from openpyxl.styles import Font # Required for processing logic
# --- MODIFICATION END ---
from typing import Dict, Any, Optional, Tuple, Set, List

# Import utility functions and constants
try:
    # Assuming utils.py is in the parent directory or accessible via Python path
    from utils import IdGenerator, replace_placeholders, read_comparison_data # Added read_comparison_data
    # Assuming config.py defines TEMPLATE_DIR or it's passed via app config
    # For simplicity here, we might redefine or get from app.config
    TEMPLATE_DIR = './config_templates/'
except ImportError as e:
    logging.error(f"Failed to import required functions/constants for processing_routes: {e}")
    # Define dummy functions or raise error if critical
    class IdGenerator:
        def __init__(self, *args, **kwargs): pass
        def get_next_dn_id(self): return 0
        def get_next_ag_id(self): return 0
    def replace_placeholders(template_data, row_data, current_row_next_id=None): return template_data
    def read_comparison_data(filename: str) -> bool: return False # Dummy returns false
    TEMPLATE_DIR = './config_templates/'

# Import functions from the refactored processing modules
# These modules contain the core logic previously in excel_comparator.py
try:
    from config import save_config # Function to save updated config
    from excel_processing import collect_routing_entities # Function to process workbook sheets
    from api_fetching import fetch_api_data # Function to call external APIs
    from comparison_logic import write_comparison_sheets # Function to write comparison results
    # Import constants needed within this blueprint's functions
    METADATA_SHEET_NAME = "Metadata"
    MAX_DN_ID_LABEL_CELL = "A1"
    MAX_DN_ID_VALUE_CELL = "B1"
    MAX_AG_ID_LABEL_CELL = "A2"
    MAX_AG_ID_VALUE_CELL = "B2"
    DN_SHEETS = {"Vqs Comparison"}
    AGENT_GROUP_SHEETS = {"Skills Comparison", "Vags Comparison", "Skill_exprs Comparison"}

except ImportError as e:
     logging.error(f"Failed to import core processing functions: {e}. Processing endpoints will fail.")
     # Define dummy functions to allow app to run, but log error
     def save_config(p, s): raise NotImplementedError("save_config not imported")
     def collect_routing_entities(w, c, m): raise NotImplementedError("collect_routing_entities not imported")
     def fetch_api_data(c): raise NotImplementedError("fetch_api_data not imported")
     def write_comparison_sheets(w, s, a, i): raise NotImplementedError("write_comparison_sheets not imported")
     # Redefine constants as fallbacks
     METADATA_SHEET_NAME = "Metadata"; MAX_DN_ID_LABEL_CELL = "A1"; MAX_DN_ID_VALUE_CELL = "B1"; MAX_AG_ID_LABEL_CELL = "A2"; MAX_AG_ID_VALUE_CELL = "B2"
     DN_SHEETS = set(); AGENT_GROUP_SHEETS = set()


# --- Constants ---
LOG_FILE_UI = 'ui_viewer.log' # Assuming shared log file
UPLOAD_FOLDER = './uploads' # Define a folder to store uploaded files temporarily
ALLOWED_EXTENSIONS = {'xlsx'}

# Keys used to identify rows in different comparison sheets (must match ui_viewer.py)
IDENTIFIER_KEYS = {
    "Skill_exprs Comparison": "Concatenated Key",
    "Vqs Comparison": "Item", # Assuming 'Item' is the first column header read
    "Skills Comparison": "Item",
    "Vags Comparison": "Item",
}


# --- Logging ---
# Use the root logger configured in the main app (app.py)
logger = logging.getLogger(__name__) # Use module-specific logger

# --- Blueprint Definition ---
# Create a Blueprint named 'processing'. The main app (app.py) will register this with '/api' prefix.
processing_bp = Blueprint('processing', __name__)

# --- Helper Functions ---
def allowed_file(filename):
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- API Routes ---

@processing_bp.route('/upload-process', methods=['POST'])
def upload_and_process_file():
    """
    Handles upload of the *original* Excel file, triggers the comparison process
    (using imported logic from excel_comparator parts), saves the
    '*_processed.xlsx' output, and reloads the UI data cache from the output file.
    Redirects back to the results viewer on completion or error.
    """
    logger.info("Received request to upload and process file.")
    # --- File Upload Handling ---
    if 'excelFile' not in request.files:
        flash('No file part in the request.', 'error')
        logger.warning("File upload request missing 'excelFile' part.")
        return redirect(url_for('ui.upload_config_page')) # Redirect back to upload page

    file = request.files['excelFile']
    # If the user does not select a file, the browser submits an
    # empty file without a filename.
    if file.filename == '':
        flash('No selected file.', 'warning')
        logger.warning("File upload request received with no selected file.")
        return redirect(url_for('ui.upload_config_page')) # Redirect back to upload page

    # Check if the file is allowed and process it
    if file and allowed_file(file.filename):
        # Ensure the upload folder exists
        if not os.path.exists(UPLOAD_FOLDER):
            try:
                os.makedirs(UPLOAD_FOLDER)
                logger.info(f"Created upload directory: {UPLOAD_FOLDER}")
            except OSError as e:
                 logger.error(f"Could not create upload directory {UPLOAD_FOLDER}: {e}")
                 flash(f"Server error: Could not create upload directory.", 'error')
                 return redirect(url_for('ui.upload_config_page'))


        # Save the uploaded *original* file temporarily or with a unique name
        original_filename = secure_filename(file.filename) # Sanitize filename
        # Define path for the output file (will be created by processing)
        processed_filename = f"{os.path.splitext(original_filename)[0]}_processed.xlsx"
        processed_filepath = os.path.join(UPLOAD_FOLDER, processed_filename)
        # Save the original uploaded file (this will be the input for processing)
        original_filepath = os.path.join(UPLOAD_FOLDER, original_filename) # Path to save original

        try:
            file.save(original_filepath)
            logger.info(f"Uploaded original file saved to: {original_filepath}")
        except Exception as e:
            logger.error(f"Error saving uploaded file: {e}", exc_info=True)
            flash(f"An error occurred saving the uploaded file: {e}", 'error')
            return redirect(url_for('ui.upload_config_page'))

        # --- Trigger Core Comparison Logic (using original_filepath as input) ---
        logger.info(f"Starting core comparison process using: {original_filepath}")
        config = current_app.config.get('APP_SETTINGS', {}) # Get loaded app config

        workbook = None # Initialize workbook variable
        try:
            # Step 0: Create a working copy for processing (output path)
            # This ensures the original upload isn't modified directly
            shutil.copyfile(original_filepath, processed_filepath)
            logger.info(f"Copied uploaded file to '{processed_filepath}' for processing.")

            # Load the copied workbook for processing
            logging.info(f"Loading workbook: {processed_filepath}")
            workbook = openpyxl.load_workbook(processed_filepath, read_only=False, data_only=False)

            # Step 1: Fetch API data AND Separate Max IDs
            api_data, max_dn_id, max_ag_id = fetch_api_data(config)
            if not any(api_data.values()):
                logging.warning("API data fetch resulted in empty/partial datasets.")
                # Store potentially empty API data for ID generator fallback?
                # current_app.config['API_DATA_CACHE'] = api_data # Optional

            # Step 2: Collect data from Excel sheets (modifies workbook)
            sheet_data_for_comparison, intermediate_data = collect_routing_entities(
                workbook, config, METADATA_SHEET_NAME
            )

            # Step 3: Perform comparison and write comparison sheets (modifies workbook)
            write_comparison_sheets(
                workbook, sheet_data_for_comparison, api_data, intermediate_data
            )

            # Step 4: Write Metadata sheet (modifies workbook)
            if METADATA_SHEET_NAME in workbook.sheetnames:
                metadata_sheet = workbook[METADATA_SHEET_NAME]
            else:
                metadata_sheet = workbook.create_sheet(title=METADATA_SHEET_NAME)
            # Write labels and values
            metadata_sheet[MAX_DN_ID_LABEL_CELL] = "Max DN API ID Found"
            metadata_sheet[MAX_DN_ID_LABEL_CELL].font = Font(bold=True) # Use imported Font
            metadata_sheet[MAX_DN_ID_VALUE_CELL] = max_dn_id
            metadata_sheet[MAX_AG_ID_LABEL_CELL] = "Max AgentGroup API ID Found"
            metadata_sheet[MAX_AG_ID_LABEL_CELL].font = Font(bold=True) # Use imported Font
            metadata_sheet[MAX_AG_ID_VALUE_CELL] = max_ag_id
            logging.info(f"Wrote Max IDs (DN:{max_dn_id}, AG:{max_ag_id}) to '{METADATA_SHEET_NAME}'.")

            # Step 5: Save the processed workbook
            workbook.save(processed_filepath)
            logger.info(f"Successfully saved processed workbook to: {processed_filepath}")

            # --- Update App Cache ---
            # Clear old cache and reload data from the newly generated processed file
            logger.info("Reloading application data cache from processed file...")
            current_app.config['EXCEL_DATA'] = {}
            current_app.config['EXCEL_FILENAME'] = None
            current_app.config['COMPARISON_SHEETS'] = []
            current_app.config['SHEET_HEADERS'] = {}
            current_app.config['MAX_DN_ID'] = 0
            current_app.config['MAX_AG_ID'] = 0

            # Use the read_comparison_data function from utils
            if read_comparison_data(processed_filepath): # Read the new file into cache
                 flash(f"File '{original_filename}' processed successfully. Max IDs updated.", 'success')
                 logger.info("Application cache updated successfully.")
                 # Redirect to the first results page after processing
                 available_sheets = current_app.config.get('COMPARISON_SHEETS', [])
                 if available_sheets:
                     return redirect(url_for('ui.view_comparison', comparison_type=available_sheets[0]))
                 else:
                     # If no comparison sheets were generated (maybe error or empty results)
                     flash("Processing complete, but no comparison sheets were generated.", "warning")
                     return redirect(url_for('ui.upload_config_page'))
            else:
                 # This case indicates an error reading the file we just created
                 flash(f"File '{original_filename}' processed, but failed to reload data into UI cache. Check logs.", 'error')
                 logger.error("Failed to reload data cache after processing.")
                 return redirect(url_for('ui.upload_config_page')) # Stay on upload page

        except Exception as proc_err:
            # Catch errors during the core processing logic
            logger.error(f"Error during file processing: {proc_err}", exc_info=True)
            flash(f"Error processing file '{original_filename}': {proc_err}", 'error')
            # Redirect back to upload page on processing error
            return redirect(url_for('ui.upload_config_page'))
        finally:
            # Ensure workbook is closed
            if workbook:
                try:
                    workbook.close()
                    logger.debug("Processing workbook closed.")
                except Exception as close_e:
                     logging.warning(f"Error closing processing workbook: {close_e}")
            # Clean up the original uploaded file? Or keep it?
            # For now, keep both original and processed in uploads/
            # if os.path.exists(original_filepath):
            #     try:
            #         os.remove(original_filepath)
            #         logger.info(f"Removed original upload file: {original_filepath}")
            #     except OSError as rm_err:
            #         logger.warning(f"Could not remove original upload file {original_filepath}: {rm_err}")


    else:
        # File type not allowed
        flash('Invalid file type. Please upload an .xlsx file.', 'error')
        logger.warning(f"Invalid file type uploaded: {file.filename}")
        return redirect(url_for('ui.upload_config_page'))


@processing_bp.route('/update-config', methods=['POST'])
def update_config():
    """
    API endpoint to receive updated configuration data from the UI
    and save it back to the config.ini file.
    """
    logger.info("Received request to update configuration.")
    try:
        # Extract form data - keys must match the 'name' attributes in the HTML form
        settings_to_save = {
            # 'source_file': request.form.get('source_file'), # Removed source_file
            'dn_url': request.form.get('dn_url'),
            'agent_group_url': request.form.get('agent_group_url'),
            'api_timeout': request.form.get('timeout', type=int, default=15), # Get as int
            'ideal_agent_header_text': request.form.get('ideal_agent_header_text'),
            'ideal_agent_fallback_cell': request.form.get('ideal_agent_fallback_cell'),
            'vag_extraction_sheet': request.form.get('vag_extraction_sheet'),
        }
        # Remove None values if any fields were missing in the form (optional)
        settings_to_save = {k: v for k, v in settings_to_save.items() if v is not None}

        # Get config file path from app settings or use default
        # Ensure 'config_file_path' is set during app creation if not using default
        config_path = current_app.config.get('APP_SETTINGS', {}).get('config_file_path', 'config.ini')

        # Use the save_config function (imported from config.py)
        save_config(config_path, settings_to_save)

        # Update the config cache in the running app
        current_app.config['APP_SETTINGS'].update(settings_to_save) # Merge updates
        logger.info("Configuration saved and application cache updated.")
        flash('Configuration saved successfully to config.ini.', 'success')

    except (IOError, ValueError, Exception) as e:
        logger.error(f"Error saving configuration: {e}", exc_info=True)
        flash(f'Error saving configuration: {e}', 'error')

    # Redirect back to the config page
    return redirect(url_for('ui.upload_config_page'))


@processing_bp.route('/simulate-configuration', methods=['POST'])
def simulate_configuration():
    """
    API endpoint to simulate applying a template to selected rows.
    Generates JSON payloads using placeholders and returns them for review.
    (Logic moved from old update_routes.py)
    """
    logger.info("Request received for /api/simulate-configuration")
    try:
        # --- 1. Get and Validate Request Data ---
        request_data = request.get_json()
        if not request_data:
            logger.warning("Simulate config request: Invalid/empty JSON payload.")
            return jsonify({"error": "Invalid JSON payload received."}), 400

        template_name = request_data.get('templateName')
        selected_row_identifiers = request_data.get('selectedRowsData') # Identifiers from first column

        if not template_name or selected_row_identifiers is None:
            logger.warning("Simulate config request missing 'templateName' or 'selectedRowsData'.")
            return jsonify({"error": "Missing 'templateName' or 'selectedRowsData'."}), 400
        if not isinstance(selected_row_identifiers, list):
             logger.warning("'selectedRowsData' received is not a list.")
             return jsonify({"error": "'selectedRowsData' must be a list."}), 400

        logger.info(f"Attempting to simulate template '{template_name}' for {len(selected_row_identifiers)} selected item(s).")

        # --- 2. Load the Specified Template ---
        if '..' in template_name or template_name.startswith('/'):
            logger.error(f"Invalid template name requested: {template_name}")
            abort(400, description="Invalid template name.")

        template_path = os.path.join(TEMPLATE_DIR, template_name)
        if not os.path.exists(template_path) or not os.path.isfile(template_path):
            logger.error(f"Template file not found at path: {template_path}")
            return jsonify({"error": f"Template '{template_name}' not found."}), 404

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_json = json.load(f)
            logger.debug(f"Successfully loaded template '{template_name}'.")
        except Exception as e:
            logger.error(f"Error reading/parsing template {template_name}: {e}", exc_info=True)
            return jsonify({"error": f"Failed to load/parse template '{template_name}'."}), 500

        # --- 3. Retrieve Data for Selected Rows (from cache using dynamic headers) ---
        all_excel_data = current_app.config.get('EXCEL_DATA', {})
        sheet_headers_map = current_app.config.get('SHEET_HEADERS', {}) # Get cached headers
        rows_to_process = [] # Will store tuples of (row_data, entity_type)
        processed_identifiers = set()

        # Iterate through sheets to find matching rows and determine their type
        for sheet_name, sheet_data in all_excel_data.items():
            headers = sheet_headers_map.get(sheet_name)
            if not headers:
                logging.warning(f"Headers not found in cache for sheet '{sheet_name}', skipping row lookup.")
                continue

            id_key = headers[0] # Use the first header as the identifier key

            # Determine entity type based on sheet name for ID generation
            entity_type = None
            if sheet_name in DN_SHEETS:
                entity_type = 'dn'
            elif sheet_name in AGENT_GROUP_SHEETS:
                entity_type = 'agent_group'
            else:
                logger.warning(f"Sheet '{sheet_name}' not mapped to DN or AG type. Cannot determine ID sequence.")

            for row in sheet_data:
                row_identifier = row.get(id_key)
                if row_identifier in selected_row_identifiers and row_identifier not in processed_identifiers:
                    rows_to_process.append((row, entity_type)) # Store row data AND its type
                    processed_identifiers.add(row_identifier)

        # Log counts
        found_count = len(rows_to_process)
        missing_identifiers = set(selected_row_identifiers) - processed_identifiers
        missing_count = len(missing_identifiers)
        logger.info(f"Retrieved data for {found_count} of {len(selected_row_identifiers)} identifiers from cached data.")
        if missing_count > 0:
             logger.warning(f"Could not find cached data for identifiers: {missing_identifiers}")


        # --- 4. Generate Payloads using Template and Row Data ---
        generated_payloads = []
        processing_errors = []
        # Initialize the ID generator using Max IDs from app config
        id_generator = IdGenerator(
            max_dn_id=current_app.config.get('MAX_DN_ID', 0),
            max_ag_id=current_app.config.get('MAX_AG_ID', 0)
        )

        # Process each row found in the cache
        for row_data, entity_type in rows_to_process:
            # Determine identifier for logging (use first header value)
            first_header = list(row_data.keys())[0] if row_data else 'UNKNOWN_KEY'
            row_id_for_log = row_data.get(first_header, "UNKNOWN_ID")

            try:
                # Generate the next ID based on entity type *before* processing the row template
                current_row_id = None
                if entity_type == 'dn':
                    current_row_id = id_generator.get_next_dn_id()
                elif entity_type == 'agent_group':
                    current_row_id = id_generator.get_next_ag_id()
                else:
                    logger.warning(f"Cannot generate ID for row '{row_id_for_log}' - unknown entity type '{entity_type}'.")

                logger.debug(f"Using next_id={current_row_id} (type: {entity_type}) for row '{row_id_for_log}'")

                # Generate the payload using the template, row data, and the pre-generated ID
                generated_payload = replace_placeholders(
                    template_data=template_json,
                    row_data=row_data, # Pass the row dict with header keys
                    current_row_next_id=current_row_id # Pass the specific ID for this row
                )

                generated_payloads.append(generated_payload)
                logging.debug(f"Generated simulation payload for row identifier '{row_id_for_log}'.")
            except Exception as e:
                logger.error(f"Error processing template for row identifier '{row_id_for_log}': {e}", exc_info=True)
                processing_errors.append(f"Error processing row '{row_id_for_log}': {e}")

        # --- 5. Construct and Return Simulation Response ---
        response_status_code = 200 # Default OK
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
@processing_bp.route('/confirm-update', methods=['POST'])
def confirm_update():
    """
    API endpoint to receive previously generated payloads (from simulation)
    and perform the final (simulated) database update action by logging them.

    Expects JSON payload: {"payloads": [ { ... }, { ... } ]}

    Returns: JSON response indicating success or failure of the simulated update.
    """
    logger.info("Request received for /api/confirm-update")
    try:
        # --- 1. Get and Validate Request Data ---
        request_data = request.get_json()
        if not request_data:
            logger.warning("Confirm update request: Invalid/empty JSON payload.")
            return jsonify({"error": "Invalid JSON payload received."}), 400

        payloads_to_commit = request_data.get('payloads')

        if payloads_to_commit is None or not isinstance(payloads_to_commit, list):
            logger.warning("Confirm update request missing 'payloads' list or invalid format.")
            return jsonify({"error": "Missing 'payloads' list or invalid format."}), 400

        logger.info(f"Received {len(payloads_to_commit)} payloads for final (simulated) update.")

        # --- 2. SIMULATE Database Update ---
        commit_errors = []
        commit_success_count = 0

        logger.info(f"--- SIMULATING FINAL DATABASE UPDATE (START) ---")
        for i, payload in enumerate(payloads_to_commit):
            try:
                # *** Replace this block with your actual database update logic ***
                logger.info(f"Simulating DB Update for Payload {i+1}: {json.dumps(payload, indent=2)}")
                commit_success_count += 1 # Assume success for simulation
                # *** End of placeholder for actual DB logic ***
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
            "errors": [str(e) for e in commit_errors]
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

