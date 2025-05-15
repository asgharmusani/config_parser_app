# -*- coding: utf-8 -*-
"""
Flask Blueprint for managing Excel Processing Rule Templates (.json files).

This blueprint handles routes related to:
- Displaying the Excel rule template management UI page.
- Listing existing Excel rule template files.
- Retrieving the content of a specific Excel rule template file.
- Saving new or updated Excel rule template files.
- Deleting Excel rule template files.
"""

import os
import json
import logging
from flask import (
    Blueprint, request, jsonify, render_template, abort, current_app, session, url_for
)
from typing import Dict, Any, Optional, List

# --- Constants ---
# Directory where Excel processing rule templates are stored
EXCEL_RULE_TEMPLATE_DIR = './excel_rule_templates/' # Ensure this matches the directory created
LOG_FILE_UI = 'ui_viewer.log'        # Assuming shared log file with main app

# --- Logging ---
# Use the root logger configured in the main app (app.py)
logger = logging.getLogger(__name__) # Use module-specific logger

# --- Blueprint Definition ---
# Create a Blueprint named 'excel_rules'. The main app (app.py) will register this.
# Point to the main 'templates' folder where HTML files reside.
excel_rule_bp = Blueprint('excel_rules', __name__, template_folder='../templates')


# --- Backend Routes ---

@excel_rule_bp.route('/') # Route relative to the blueprint prefix (e.g., '/excel-rules/')
def excel_rule_manager_page():
    """
    Renders the main UI page for managing Excel processing rule templates.
    Uses the 'excel_rule_manager.html' template.
    Determines the correct 'Back' link URL based on session data.
    """
    logger.info("Rendering Excel Rule Template Manager page.")

    # Determine the URL for the 'Back' link
    last_viewed_comparison = session.get('last_viewed_comparison')
    back_url = url_for('ui.upload_config_page') # Default back URL

    if last_viewed_comparison:
        # If a comparison page was last viewed, try to generate URL back to it
        try:
            # Ensure the comparison type is valid before generating URL
            # This check might be more robust if we verify against current_app.config['COMPARISON_SHEETS']
            if last_viewed_comparison in current_app.config.get('COMPARISON_SHEETS', []):
                back_url = url_for('ui.view_comparison', comparison_type=last_viewed_comparison)
                logger.debug(f"Setting back URL to last viewed comparison: {last_viewed_comparison}")
            else:
                logger.warning(f"Last viewed comparison '{last_viewed_comparison}' not in current available sheets. Defaulting back URL.")
        except Exception as e:
            # Handle cases where url_for might fail
            logger.warning(f"Could not build URL for last viewed comparison '{last_viewed_comparison}': {e}. Defaulting back URL.")
            # back_url remains the default url_for('ui.upload_config_page')
    else:
        logger.debug("Setting back URL to upload/config page (no last viewed comparison found in session).")

    # Render the excel_rule_manager.html file
    return render_template('excel_rule_manager.html', back_url=back_url)


@excel_rule_bp.route('/list', methods=['GET'])
def list_excel_rule_templates():
    """
    API endpoint to list available .json Excel rule template filenames.

    Returns:
        JSON list of filenames on success.
        JSON error object on failure.
    """
    logger.info("Request received to list Excel rule templates.")
    try:
        # Ensure the template directory exists, create if not
        if not os.path.exists(EXCEL_RULE_TEMPLATE_DIR):
            os.makedirs(EXCEL_RULE_TEMPLATE_DIR)
            logger.info(f"Created Excel rule template directory: {EXCEL_RULE_TEMPLATE_DIR}")
            return jsonify([]) # Return empty list if directory was just created

        # List files ending with .json (case-insensitive) that are actual files
        files = [
            f for f in os.listdir(EXCEL_RULE_TEMPLATE_DIR)
            if f.lower().endswith('.json') and os.path.isfile(os.path.join(EXCEL_RULE_TEMPLATE_DIR, f))
        ]
        logger.debug(f"Found Excel rule template files: {files}")
        # Return the sorted list of filenames as JSON
        return jsonify(sorted(files))
    except Exception as e:
        logger.error(f"Error listing Excel rule templates in {EXCEL_RULE_TEMPLATE_DIR}: {e}", exc_info=True)
        # Return a server error response
        return jsonify({"error": "Failed to list Excel rule templates"}), 500

@excel_rule_bp.route('/get/<path:filename>', methods=['GET'])
def get_excel_rule_template(filename):
    """
    API endpoint to retrieve the JSON content of a specific Excel rule template file.

    Args:
        filename: The name of the template file (including .json extension).

    Returns:
        JSON content of the file on success.
        JSON error object or aborts on failure (400, 404, 500).
    """
    logger.info(f"Request received to get Excel rule template: {filename}")
    # Basic security check to prevent path traversal
    if '..' in filename or filename.startswith('/'):
        logger.warning(f"Attempted access to invalid path for Excel rule template: {filename}")
        abort(400, description="Invalid filename.") # Bad request

    try:
        filepath = os.path.join(EXCEL_RULE_TEMPLATE_DIR, filename)
        # Check if the file exists and is actually a file
        if not os.path.exists(filepath) or not os.path.isfile(filepath):
             logger.warning(f"Excel rule template file not found: {filepath}")
             abort(404, description="Excel rule template not found.") # Not found

        # Read and parse the JSON file content
        with open(filepath, 'r', encoding='utf-8') as f:
            content = json.load(f)
        logger.debug(f"Successfully read Excel rule template content from: {filepath}")
        return jsonify(content) # Return JSON content
    except json.JSONDecodeError:
        # Handle case where the file is not valid JSON
        logger.error(f"Invalid JSON in Excel rule template file: {filepath}")
        return jsonify({"error": "Excel rule template file contains invalid JSON."}), 500 # Server error
    except Exception as e:
        # Handle other file reading errors
        logger.error(f"Error reading Excel rule template file {filename}: {e}", exc_info=True)
        return jsonify({"error": "Failed to read Excel rule template file."}), 500 # Server error

@excel_rule_bp.route('/save', methods=['POST'])
def save_excel_rule_template():
    """
    API endpoint to save JSON content to an Excel rule template file.
    Creates a new file or overwrites an existing one.

    Expects JSON payload: {"filename": "my_rule_template.json", "content": { ... }}

    Returns:
        JSON success message on success.
        JSON error object or aborts on failure (400, 500).
    """
    logger.info("Request received to save Excel rule template.")
    try:
        # Get JSON data from the request body
        data = request.get_json()
        if not data or 'filename' not in data or 'content' not in data:
            logger.warning("Save Excel rule template request missing filename or content.")
            abort(400, description="Missing filename or content in request.") # Bad request

        filename = data['filename']
        content = data['content'] # Expecting already parsed JSON object/dict from JS validation

        # --- Filename Validation ---
        if not isinstance(filename, str) or not filename:
             abort(400, description="Filename must be a non-empty string.")
        if not filename.lower().endswith('.json'):
             abort(400, description="Filename must end with .json")
        # Prevent path traversal and invalid characters typically disallowed in filenames
        if '..' in filename or filename.startswith('/') or os.path.dirname(filename) or any(c in filename for c in r'<>:"|?*\\/'):
             logger.warning(f"Attempted save with invalid Excel rule template filename: {filename}")
             abort(400, description="Invalid characters or path in filename.")

        # Ensure the template directory exists
        if not os.path.exists(EXCEL_RULE_TEMPLATE_DIR):
            os.makedirs(EXCEL_RULE_TEMPLATE_DIR)
            logger.info(f"Created Excel rule template directory during save: {EXCEL_RULE_TEMPLATE_DIR}")

        filepath = os.path.join(EXCEL_RULE_TEMPLATE_DIR, filename)
        base_name = filename.replace('.json', '')

        # Check if overwriting or creating new to provide accurate logging/message
        is_update = os.path.exists(filepath)
        action = "Updated" if is_update else "Created"

        # Write the JSON content to the file, pretty-printed with indent=2
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(content, f, indent=2)

        logger.info(f"{action} Excel rule template file: {filepath}")
        # Return success message with appropriate HTTP status code
        return jsonify({"message": f"Excel rule template '{base_name}' saved successfully."}), 200 if is_update else 201 # OK or Created

    except Exception as e:
        # Handle file writing errors or other unexpected issues
        logger.error(f"Error saving Excel rule template file: {e}", exc_info=True)
        return jsonify({"error": "Failed to save Excel rule template file."}), 500 # Server error

# Allow both DELETE and POST for deletion for flexibility (browsers sometimes block DELETE)
@excel_rule_bp.route('/delete/<path:filename>', methods=['DELETE', 'POST'])
def delete_excel_rule_template(filename):
    """
    API endpoint to delete a specific Excel rule template file.

    Args:
        filename: The name of the template file to delete.

    Returns:
        JSON success message on success.
        JSON error object or aborts on failure (400, 404, 500).
    """
    logger.info(f"Request received to delete Excel rule template: {filename}")
    # Basic security check
    if '..' in filename or filename.startswith('/'):
        logger.warning(f"Attempted delete with invalid path for Excel rule template: {filename}")
        abort(400, description="Invalid filename.") # Bad request
    try:
        filepath = os.path.join(EXCEL_RULE_TEMPLATE_DIR, filename)
        base_name = filename.replace('.json', '')

        # Check if file exists before attempting deletion
        if os.path.exists(filepath) and os.path.isfile(filepath):
            os.remove(filepath)
            logger.info(f"Deleted Excel rule template file: {filepath}")
            return jsonify({"message": f"Excel rule template '{base_name}' deleted successfully."}), 200 # OK
        else:
            # File not found
            logger.warning(f"Attempted to delete non-existent Excel rule template: {filepath}")
            abort(404, description="Excel rule template not found.") # Not found

    except Exception as e:
        # Handle file deletion errors
        logger.error(f"Error deleting Excel rule template file {filename}: {e}", exc_info=True)
        return jsonify({"error": "Failed to delete Excel rule template file."}), 500 # Server error

# Note: The proxy_api_fetch route is specific to the DB update template manager (template_routes.py)
# and is not needed for Excel rule templates unless a similar use case arises.
# If needed for fetching example Excel structures or something similar, it could be added here.

