# -*- coding: utf-8 -*-
"""
Flask Blueprint for handling backend processing tasks and API endpoints.

Includes API endpoints for:
- Uploading the original source Excel file.
- Triggering the comparison process on an uploaded file (using rule-defined API URLs for comparison).
- Loading data from an existing processed file.
- Simulating configuration template application.
- Confirming and finalizing (simulated) updates.
- Updating the application configuration.
"""

import os
import json
import logging
import re
import shutil # For file operations
import openpyxl # Needed for loading workbook in processing route
from flask import (
    Blueprint, request, jsonify, current_app, abort, flash, redirect, url_for
)
from werkzeug.utils import secure_filename # For secure file uploads
from openpyxl.styles import Font # Required for processing logic
from typing import Dict, Any, Optional, Tuple, Set, List

# Import utility functions and constants
try:
    # Assuming utils.py is in the parent directory or accessible via Python path
    from utils import IdGenerator, replace_placeholders, read_comparison_data
    # TEMPLATE_DIR is for DB update templates
    TEMPLATE_DIR = './config_templates/'
    # EXCEL_RULE_TEMPLATE_DIR for Excel processing rules
    EXCEL_RULE_TEMPLATE_DIR = './excel_rule_templates/'
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
    EXCEL_RULE_TEMPLATE_DIR = './excel_rule_templates/'


# Import functions from the refactored processing modules
try:
    from config import save_config # Function to save updated config
    from excel_processing import collect_and_write_excel_outputs
    from api_fetching import fetch_and_process_api_data_for_entity # Updated import
    from comparison_logic import write_comparison_sheets
    # Import constants needed within this blueprint's functions
    METADATA_SHEET_NAME = "Metadata"
    MAX_DN_ID_LABEL_CELL = "A1"
    MAX_DN_ID_VALUE_CELL = "B1"
    MAX_AG_ID_LABEL_CELL = "A2"
    MAX_AG_ID_VALUE_CELL = "B2"
    # Sheet type definitions for ID generation (based on comparison sheet names)
    # These are used by IdGenerator in this module when processing rows for simulation
    DN_SHEETS = {"Vqs Comparison"} # Example, this might need to be more dynamic based on rule's idPoolType
    AGENT_GROUP_SHEETS = {"Skills Comparison", "Vags Comparison", "Skill_exprs Comparison"}

except ImportError as e:
     logging.error(f"Failed to import core processing functions: {e}. Processing endpoints will fail.")
     # Define dummy functions to allow app to run, but log error
     def save_config(p, s): raise NotImplementedError("save_config not imported")
     def collect_and_write_excel_outputs(w, p_e, c, m, s_t_r_c): raise NotImplementedError("collect_and_write_excel_outputs not imported")
     def fetch_and_process_api_data_for_entity(u, en, r, c): return ({}, 0) # Dummy returns data and max_id
     def write_comparison_sheets(w, s, a, i): raise NotImplementedError("write_comparison_sheets not imported")
     # Redefine constants as fallbacks
     METADATA_SHEET_NAME = "Metadata"; MAX_DN_ID_LABEL_CELL = "A1"; MAX_DN_ID_VALUE_CELL = "B1"; MAX_AG_ID_LABEL_CELL = "A2"; MAX_AG_ID_VALUE_CELL = "B2"
     DN_SHEETS = set(); AGENT_GROUP_SHEETS = set()


# --- Constants ---
LOG_FILE_UI = 'ui_viewer.log' # Assuming shared log file
UPLOAD_FOLDER = './uploads' # Define a folder to store uploaded and processed files
ALLOWED_EXTENSIONS = {'xlsx'}


# --- Logging ---
# Use the root logger configured in the main app (app.py)
logger = logging.getLogger(__name__) # Use module-specific logger

# --- Blueprint Definition ---
# Create a Blueprint named 'processing'. The main app (app.py) will register this with '/api' prefix.
processing_bp = Blueprint('processing', __name__)

# --- Helper Functions ---
def allowed_file(filename: str) -> bool:
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# map_rule_entity_name_to_api_key is no longer needed as API URLs are per rule.

# --- API Routes ---

@processing_bp.route('/upload-original-file', methods=['POST'])
def upload_original_file():
    """
    Handles upload of the original source Excel file.
    Saves the file and stores its path for later processing.
    """
    logger.info("Received request to upload original file.")
    if 'sourceExcelFile' not in request.files:
        logger.warning("File upload request missing 'sourceExcelFile' part.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['sourceExcelFile']
    if file.filename == '':
        logger.warning("File upload request received with no selected file.")
        return jsonify({"error": "No selected file."}), 400

    if file and allowed_file(file.filename):
        if not os.path.exists(UPLOAD_FOLDER):
            try:
                os.makedirs(UPLOAD_FOLDER)
                logger.info(f"Created upload directory: {UPLOAD_FOLDER}")
            except OSError as e:
                 logger.error(f"Could not create upload directory {UPLOAD_FOLDER}: {e}")
                 return jsonify({"error": f"Server error: Could not create upload directory."}), 500

        original_filename = secure_filename(file.filename)
        original_filepath = os.path.join(UPLOAD_FOLDER, original_filename)

        try:
            file.save(original_filepath)
            logger.info(f"Uploaded original file saved to: {original_filepath}")
            current_app.config['LAST_UPLOADED_ORIGINAL_FILE'] = original_filepath
            return jsonify({
                "message": f"File '{original_filename}' uploaded successfully. Ready to process.",
                "original_filename": original_filename
                }), 200
        except Exception as e:
            logger.error(f"Error saving uploaded file: {e}", exc_info=True)
            return jsonify({"error": f"An error occurred saving the uploaded file: {e}"}), 500
    else:
        logger.warning(f"Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Please upload an .xlsx file."}), 400


@processing_bp.route('/run-comparison', methods=['POST'])
def run_comparison():
    """
    Triggers the comparison process using the last uploaded original file
    and the selected Excel rule template.
    Generates the '*_processed.xlsx' file and loads its data into the cache.
    API data for comparison is fetched based on 'comparisonApiUrl' in rules.
    Max IDs for ID generation are aggregated from these specific API calls.
    """
    logger.info("Received request to run comparison process.")
    request_data = request.get_json()
    if not request_data:
        logger.warning("Run comparison request missing JSON body.")
        return jsonify({"error": "Missing request data."}), 400

    excel_rule_template_name = request_data.get('excelRuleTemplateName')
    if not excel_rule_template_name:
        logger.warning("Run comparison request missing 'excelRuleTemplateName'.")
        return jsonify({"error": "Excel rule template name not provided."}), 400

    original_filepath = current_app.config.get('LAST_UPLOADED_ORIGINAL_FILE')
    if not original_filepath or not os.path.exists(original_filepath):
        logger.error("Run comparison triggered, but no valid original file path found in config.")
        return jsonify({"error": "No valid source file has been uploaded recently or file not found."}), 400

    original_filename = os.path.basename(original_filepath)
    processed_filename = f"{os.path.splitext(original_filename)[0]}_processed.xlsx"
    processed_filepath = os.path.join(UPLOAD_FOLDER, processed_filename)

    logger.info(f"Starting comparison: original='{original_filepath}', rule='{excel_rule_template_name}'")
    app_config_settings = current_app.config.get('APP_SETTINGS', {})

    # --- Load Excel Rule Template ---
    rule_template_path = os.path.join(EXCEL_RULE_TEMPLATE_DIR, excel_rule_template_name)
    if not os.path.exists(rule_template_path):
        logger.error(f"Excel rule template file not found: {rule_template_path}")
        return jsonify({"error": f"Excel rule template '{excel_rule_template_name}' not found."}), 404
    try:
        with open(rule_template_path, 'r', encoding='utf-8') as f:
            rule_template_json = json.load(f)
        logger.info(f"Successfully loaded Excel rule template: {excel_rule_template_name}")
    except Exception as e:
        logger.error(f"Error loading or parsing Excel rule template '{excel_rule_template_name}': {e}", exc_info=True)
        return jsonify({"error": f"Could not load or parse Excel rule template: {e}"}), 500

    # --- Instantiate Rule Engine ---
    try:
        from excel_rule_engine import ExcelRuleEngine
        rule_engine = ExcelRuleEngine(rule_template_json)
    except Exception as e:
        logger.error(f"Error initializing ExcelRuleEngine: {e}", exc_info=True)
        return jsonify({"error": f"Failed to initialize Excel rule engine: {e}"}), 500

    # --- Main Processing Logic ---
    output_workbook = None
    try:
        # 1. Make a working copy of the original file for processing
        shutil.copyfile(original_filepath, processed_filepath)
        logger.info(f"Copied original file to '{processed_filepath}' for processing.")

        # 2. Load the original workbook for the rule engine to read
        original_workbook_for_rules = openpyxl.load_workbook(original_filepath, read_only=False, data_only=False)
        parsed_entities = rule_engine.process_workbook(original_workbook_for_rules)
        original_workbook_for_rules.close()

        # --- 3. Fetch API data PER ENTITY RULE for comparison AND Aggregate Max IDs ---
        api_data_for_comparison = {}
        overall_max_dn_id = 0 # Initialize overall max IDs
        overall_max_ag_id = 0

        if rule_template_json and "Entities" in rule_template_json:
            for entity_rule in rule_template_json["Entities"]:
                if not entity_rule.get("enabled", True):
                    continue

                entity_name = entity_rule["name"]
                api_url_for_comparison = entity_rule.get("comparisonApiUrl")
                # Get the idPoolType hint from the rule
                id_pool_type = entity_rule.get("idPoolType") # e.g., "dn" or "agent_group"

                if api_url_for_comparison:
                    logger.info(f"Fetching API data for comparison for entity '{entity_name}' using URL: {api_url_for_comparison}")
                    # fetch_and_process_api_data_for_entity now returns (processed_data, max_id_from_this_api_call)
                    processed_data, max_id_this_api = fetch_and_process_api_data_for_entity(
                        api_url_for_comparison,
                        entity_name,
                        entity_rule, # Pass the whole rule for hints like 'apiProcessingHints'
                        app_config_settings # Pass global app config for timeout
                    )
                    api_data_for_comparison[entity_name] = processed_data

                    # --- MODIFICATION START: Aggregate Max IDs based on idPoolType ---
                    if id_pool_type == 'dn':
                        overall_max_dn_id = max(overall_max_dn_id, max_id_this_api)
                    elif id_pool_type == 'agent_group':
                        overall_max_ag_id = max(overall_max_ag_id, max_id_this_api)
                    elif max_id_this_api > 0: # If pool type not specified but IDs found
                        logger.warning(f"Max ID {max_id_this_api} found for entity '{entity_name}' from API '{api_url_for_comparison}', "
                                       f"but 'idPoolType' was not specified or recognized in the rule. This ID will not contribute to IdGenerator pools.")
                    # --- MODIFICATION END ---
                else:
                    # If no comparisonApiUrl, this entity won't be compared against an API
                    api_data_for_comparison[entity_name] = {}
                    logger.info(f"No 'comparisonApiUrl' defined for entity '{entity_name}'. It will not be compared against API data.")
        logger.info(f"Aggregated Max IDs after all rule-specific API calls: DN Max ID={overall_max_dn_id}, AG Max ID={overall_max_ag_id}")
        # --- End API Fetching ---

        # 4. Load the workbook that will be modified (the copy for output)
        output_workbook = openpyxl.load_workbook(processed_filepath, read_only=False, data_only=False)

        # 5. Collect and Write Excel Outputs (Output sheets based on rules)
        sheets_to_remove_base = [METADATA_SHEET_NAME]
        sheet_data_for_comparison, intermediate_data_resolved = collect_and_write_excel_outputs(
            output_workbook, parsed_entities, app_config_settings, METADATA_SHEET_NAME, sheets_to_remove_base
        )

        # 6. Write Comparison Sheets
        # Pass the api_data_for_comparison (now keyed by rule entity names)
        write_comparison_sheets(
            output_workbook, sheet_data_for_comparison, api_data_for_comparison, intermediate_data_resolved
        )

        # 7. Write Metadata sheet with aggregated Max IDs
        if METADATA_SHEET_NAME in output_workbook.sheetnames:
            metadata_sheet = output_workbook[METADATA_SHEET_NAME]
        else:
            metadata_sheet = output_workbook.create_sheet(title=METADATA_SHEET_NAME)
        metadata_sheet[MAX_DN_ID_LABEL_CELL] = "Max DN API ID Found"
        metadata_sheet[MAX_DN_ID_LABEL_CELL].font = Font(bold=True)
        metadata_sheet[MAX_DN_ID_VALUE_CELL] = overall_max_dn_id # Use aggregated max
        metadata_sheet[MAX_AG_ID_LABEL_CELL] = "Max AgentGroup API ID Found"
        metadata_sheet[MAX_AG_ID_LABEL_CELL].font = Font(bold=True)
        metadata_sheet[MAX_AG_ID_VALUE_CELL] = overall_max_ag_id # Use aggregated max
        logging.info(f"Wrote Aggregated Max IDs (DN:{overall_max_dn_id}, AG:{overall_max_ag_id}) to '{METADATA_SHEET_NAME}'.")

        # 8. Save the processed workbook
        output_workbook.save(processed_filepath)
        logger.info(f"Successfully saved processed workbook to: {processed_filepath}")

        # --- 9. Update App Cache from the *newly generated* processed file ---
        logger.info("Reloading application data cache from processed file...")
        current_app.config['EXCEL_DATA'] = {}
        current_app.config['EXCEL_FILENAME'] = None
        current_app.config['COMPARISON_SHEETS'] = []
        current_app.config['SHEET_HEADERS'] = {}
        current_app.config['MAX_DN_ID'] = 0 # Will be updated by read_comparison_data
        current_app.config['MAX_AG_ID'] = 0 # Will be updated by read_comparison_data

        if read_comparison_data(processed_filepath): # This function is now in utils.py
             logger.info("Application cache updated successfully.")
             first_sheet = current_app.config.get('COMPARISON_SHEETS', [None])[0]
             return jsonify({
                 "message": f"File '{original_filename}' processed successfully using rule '{excel_rule_template_name}'.",
                 "processed_file": processed_filename,
                 "redirect_url": url_for('ui.view_comparison', comparison_type=first_sheet) if first_sheet else url_for('ui.upload_config_page')
                 }), 200
        else:
             logger.error("Failed to reload data cache after processing.")
             return jsonify({"error": f"File '{original_filename}' processed, but failed to reload data into UI cache. Check logs."}), 500

    except Exception as proc_err:
        logger.error(f"Error during file processing with rule '{excel_rule_template_name}': {proc_err}", exc_info=True)
        return jsonify({"error": f"Error processing file '{original_filename}' with rule '{excel_rule_template_name}': {proc_err}"}), 500
    finally:
        if output_workbook:
            try:
                output_workbook.close()
                logger.debug("Output workbook closed.")
            except Exception as close_e:
                 logging.warning(f"Error closing output workbook: {close_e}")


@processing_bp.route('/load-processed-file', methods=['POST'])
def load_processed_file():
    """
    Loads data from an existing *_processed.xlsx file (selected by user)
    into the application cache.
    """
    logger.info("Request received to load existing processed file.")
    request_data = request.get_json()
    if not request_data or 'filename' not in request_data:
        logger.warning("Load processed file request missing filename.")
        return jsonify({"error": "Filename not provided."}), 400

    filename = secure_filename(request_data['filename'])
    filepath = os.path.join(UPLOAD_FOLDER, filename)

    logger.info(f"Attempting to load data from: {filepath}")

    if not os.path.exists(filepath):
        logger.error(f"Processed file not found: {filepath}")
        return jsonify({"error": f"File '{filename}' not found in uploads directory."}), 404

    # Clear existing cache before loading
    current_app.config['EXCEL_DATA'] = {}
    current_app.config['EXCEL_FILENAME'] = None
    current_app.config['COMPARISON_SHEETS'] = []
    current_app.config['SHEET_HEADERS'] = {}
    current_app.config['MAX_DN_ID'] = 0
    current_app.config['MAX_AG_ID'] = 0

    if read_comparison_data(filepath): # This function is in utils.py
        logger.info(f"Successfully loaded data from '{filename}' into cache.")
        first_sheet = current_app.config.get('COMPARISON_SHEETS', [None])[0]
        return jsonify({
            "message": f"Successfully loaded data from '{filename}'.",
            "redirect_url": url_for('ui.view_comparison', comparison_type=first_sheet) if first_sheet else url_for('ui.upload_config_page')
        }), 200
    else:
        logger.error(f"Failed to read data from '{filename}'. Check logs.")
        # Reset cache again on failure
        current_app.config['EXCEL_DATA'] = {}
        current_app.config['EXCEL_FILENAME'] = None
        current_app.config['COMPARISON_SHEETS'] = []
        current_app.config['SHEET_HEADERS'] = {}
        current_app.config['MAX_DN_ID'] = 0
        current_app.config['MAX_AG_ID'] = 0
        return jsonify({"error": f"Failed to read data from '{filename}'. Check logs."}), 500


@processing_bp.route('/update-config', methods=['POST'])
def update_config():
    """
    API endpoint to receive updated configuration data from the UI
    and save it back to the config.ini file. Handles new 'log_level'.
    API URLs (dn_url, agent_group_url) are no longer part of global config.
    """
    logger.info("Received request to update configuration.")
    try:
        settings_to_save = {
            # 'dn_url' and 'agent_group_url' are removed from global config
            'api_timeout': request.form.get('timeout', type=int, default=15),
            'ideal_agent_header_text': request.form.get('ideal_agent_header_text'),
            'ideal_agent_fallback_cell': request.form.get('ideal_agent_fallback_cell'),
            'vag_extraction_sheet': request.form.get('vag_extraction_sheet'),
            'log_level_str': request.form.get('log_level') # Get the log level string
        }
        # Ensure essential keys have fallbacks
        if 'api_timeout' not in settings_to_save or settings_to_save['api_timeout'] is None:
            settings_to_save['api_timeout'] = 15
        if 'log_level_str' not in settings_to_save or settings_to_save['log_level_str'] is None:
            settings_to_save['log_level_str'] = 'INFO'


        config_path = current_app.config.get('CONFIG_FILE_PATH', 'config.ini')
        save_config(config_path, settings_to_save)

        # Update the config cache in the running app
        current_app.config['APP_SETTINGS'].update(settings_to_save)
        logger.info("Configuration saved and application cache updated.")
        flash('Configuration saved successfully to config.ini. Restart may be needed for some changes.', 'success')

    except (IOError, ValueError, Exception) as e:
        logger.error(f"Error saving configuration: {e}", exc_info=True)
        flash(f'Error saving configuration: {e}', 'error')
    return redirect(url_for('ui.upload_config_page'))


@processing_bp.route('/simulate-configuration', methods=['POST'])
def simulate_configuration():
    """
    API endpoint to simulate applying a template to selected rows.
    Generates JSON payloads using placeholders and returns them for review.
    """
    logger.info("Request received for /api/simulate-configuration")
    try:
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
             logger.warning("'selectedRowsData' received is not a list.")
             return jsonify({"error": "'selectedRowsData' must be a list."}), 400

        logger.info(f"Attempting to simulate template '{template_name}' for {len(selected_row_identifiers)} selected item(s).")

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

        all_excel_data = current_app.config.get('EXCEL_DATA', {})
        sheet_headers_map = current_app.config.get('SHEET_HEADERS', {})
        rows_to_process = []
        processed_identifiers = set()

        for sheet_name, sheet_data in all_excel_data.items():
            headers = sheet_headers_map.get(sheet_name)
            if not headers: continue
            id_key = headers[0] # Assumes first header is the primary identifier for selection
            # Determine entity type based on the sheet name (which is now the rule entity name)
            # This mapping is heuristic for IdGenerator. A more robust way is to pass entity_name from rule.
            entity_type = 'dn' if "vq" in sheet_name.lower() else ('agent_group' if any(s_type in sheet_name.lower() for s_type in ["skill", "vag", "expr"]) else None)

            for row in sheet_data:
                row_identifier = row.get(id_key)
                if row_identifier in selected_row_identifiers and row_identifier not in processed_identifiers:
                    rows_to_process.append((row, entity_type))
                    processed_identifiers.add(row_identifier)

        found_count = len(rows_to_process)
        missing_identifiers = set(selected_row_identifiers) - processed_identifiers
        missing_count = len(missing_identifiers)
        logger.info(f"Retrieved data for {found_count} of {len(selected_row_identifiers)} identifiers from cached data.")
        if missing_count > 0: logger.warning(f"Could not find cached data for identifiers: {missing_identifiers}")

        generated_payloads = []
        processing_errors = []
        id_generator = IdGenerator(
            max_dn_id=current_app.config.get('MAX_DN_ID', 0),
            max_ag_id=current_app.config.get('MAX_AG_ID', 0)
        )

        for row_data, entity_type in rows_to_process:
            first_header = list(row_data.keys())[0] if row_data else 'UNKNOWN_KEY'
            row_id_for_log = row_data.get(first_header, "UNKNOWN_ID")
            try:
                current_row_id = None
                if entity_type == 'dn': current_row_id = id_generator.get_next_dn_id()
                elif entity_type == 'agent_group': current_row_id = id_generator.get_next_ag_id()
                else: logger.warning(f"Cannot generate ID for row '{row_id_for_log}' - unknown entity type '{entity_type}'.")
                logger.debug(f"Using next_id={current_row_id} (type: {entity_type}) for row '{row_id_for_log}'")
                generated_payload = replace_placeholders(template_json, row_data, current_row_id)
                generated_payloads.append(generated_payload)
            except Exception as e:
                logger.error(f"Error processing template for row '{row_id_for_log}': {e}", exc_info=True)
                processing_errors.append(f"Error processing row '{row_id_for_log}': {e}")

        response_status_code = 200
        response_data = { "message": f"Simulation complete for template '{template_name}'. Generated {len(generated_payloads)} payloads for review.", "status": "Simulation Success", "processed_count": found_count, "payloads": generated_payloads, "errors": [] }
        if missing_count > 0: response_data["message"] += f" Could not find data for {missing_count} identifiers: {list(missing_identifiers)}."; response_data["status"] = "Simulation Partial Success / Missing Data"; response_status_code = 207
        if processing_errors: response_data["errors"] = [str(e) for e in processing_errors]; response_data["message"] += f" Encountered {len(processing_errors)} errors during payload generation."; response_data["status"] = "Simulation Partial Success / Errors" if response_status_code == 200 else response_data["status"]; response_status_code = 207
        logger.info(f"Simulation successful. Returning {len(generated_payloads)} generated payloads.")
        return jsonify(response_data), response_status_code

    except Exception as e:
        logger.error(f"Unexpected error in /api/simulate-configuration: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred during simulation."}), 500


@processing_bp.route('/confirm-update', methods=['POST'])
def confirm_update():
    """
    API endpoint to receive previously generated payloads and perform (simulated) DB update.
    """
    logger.info("Request received for /api/confirm-update")
    try:
        request_data = request.get_json()
        if not request_data: logger.warning("Confirm update: Invalid/empty JSON."); return jsonify({"error": "Invalid JSON payload."}), 400
        payloads_to_commit = request_data.get('payloads')
        if payloads_to_commit is None or not isinstance(payloads_to_commit, list): logger.warning("Confirm update: Missing 'payloads' list."); return jsonify({"error": "Missing 'payloads' list or invalid format."}), 400
        logger.info(f"Received {len(payloads_to_commit)} payloads for final (simulated) update.")
        commit_errors = []; commit_success_count = 0
        logger.info(f"--- SIMULATING FINAL DATABASE UPDATE (START) ---")
        for i, payload in enumerate(payloads_to_commit):
            try: logger.info(f"Simulating DB Update for Payload {i+1}: {json.dumps(payload, indent=2)}"); commit_success_count += 1
            except Exception as db_err: logger.error(f"Simulated DB update FAILED for Payload {i+1}: {db_err}", exc_info=True); commit_errors.append(f"Payload {i+1}: {db_err}")
        logger.info(f"--- SIMULATING FINAL DATABASE UPDATE (END) ---")
        response_status_code = 200
        response_data = { "message": f"Simulated update completed for {commit_success_count} of {len(payloads_to_commit)} payloads.", "status": "Update Simulation Success", "success_count": commit_success_count, "error_count": len(commit_errors), "errors": [str(e) for e in commit_errors] }
        if commit_errors: response_data["status"] = "Update Simulation Partial Success / Errors"; response_status_code = 207
        if commit_success_count == 0 and len(payloads_to_commit) > 0: response_data["status"] = "Update Simulation Failed"; response_status_code = 500
        return jsonify(response_data), response_status_code
    except Exception as e:
        logger.error(f"Unexpected error in /api/confirm-update: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred during update confirmation."}), 500

