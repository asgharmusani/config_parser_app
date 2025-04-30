# -*- coding: utf-8 -*-
"""
Flask Blueprint for managing configuration templates (.json files).

This blueprint handles routes related to:
- Displaying the template management UI page.
- Listing existing template files.
- Retrieving the content of a specific template file.
- Saving new or updated template files.
- Deleting template files.
- Proxying requests to fetch reference API data (to avoid CORS).
- Includes the placeholder replacement logic used during simulation/update.
"""

import os
import json
import logging
import datetime
import requests
import re # Added for placeholder logic
from flask import (
    Blueprint, request, jsonify, render_template, abort, current_app, session, url_for # Added render_template, session, url_for
)
from typing import Dict, Any, Optional, List # Added List

# --- Constants ---
TEMPLATE_DIR = './config_templates/' # Directory where JSON templates are stored
LOG_FILE_UI = 'ui_viewer.log'        # Assuming shared log file with main app

# --- Logging ---
# Use the root logger configured in the main app (app.py)
logger = logging.getLogger(__name__) # Use module-specific logger

# --- Blueprint Definition ---
# Create a Blueprint named 'templates'. The main app (app.py) will register this.
# Point to the main 'templates' folder where HTML files reside.
template_bp = Blueprint('templates', __name__, template_folder='../templates')

# --- Helper Function: Placeholder Replacement ---
# This function is part of the template blueprint as it defines how templates are interpreted.
def replace_placeholders(template_data: Any, row_data: dict, current_row_next_id: Optional[int] = None) -> Any:
    """
    Recursively traverses a template structure (dict, list, or string)
    and replaces placeholders with values from row_data or the pre-generated ID.

    Supported Placeholders:
    - {row.ColumnName}: Replaced with the value from the corresponding column in row_data.
                        Lookup is case-insensitive for the ColumnName part.
    - {func.next_id}: Replaced with the current_row_next_id value.

    Also handles simple string concatenation like "prefix:{row.ColumnName}".

    Args:
        template_data: The template structure (can be dict, list, string, etc.).
        row_data: The dictionary containing data for the current row (keys are actual headers).
        current_row_next_id: The pre-generated sequential ID for the current row.

    Returns:
        The template structure with placeholders replaced.
    """
    # Regex to find placeholders like {row.ColumnName} or {func.FunctionName}
    placeholder_pattern = re.compile(r'{(\w+)\.([^}]+)}') # Captures type (row/func) and name

    # --- Inner replacement function ---
    def perform_replace(text: str) -> str:
        """Performs replacements on a single string."""
        if not isinstance(text, str):
            # Return non-strings as is
            return text

        # Function to handle each match found by the regex
        def replace_match(match):
            placeholder_type = match.group(1).lower() # 'row' or 'func'
            placeholder_name = match.group(2).strip() # 'ColumnName' or 'next_id'

            if placeholder_type == 'row':
                # --- Case-insensitive lookup ---
                found_key = None
                for key in row_data.keys():
                    # Compare lowercased keys
                    if key.lower() == placeholder_name.lower():
                        found_key = key
                        break

                if found_key:
                    replacement = row_data.get(found_key, "") # Use the actual key found
                else:
                    replacement = "" # Default to empty if no matching key found
                    logger.warning(f"Placeholder {{row.{placeholder_name}}} not found in row data keys: {list(row_data.keys())}")
                # --- End Case-insensitive lookup ---
                return str(replacement) # Ensure replacement is a string

            elif placeholder_type == 'func':
                # Handle the {func.next_id} placeholder
                if placeholder_name == 'next_id':
                    if current_row_next_id is not None:
                        # Use the ID pre-generated for this specific row
                        return str(current_row_next_id)
                    else:
                        # Log a warning if the placeholder is used but no ID was provided
                        logger.warning(f"Placeholder {{func.next_id}} used but no ID provided for this row.")
                        return "{ERROR:next_id_missing}" # Indicate error in output
                else:
                    # Handle unknown function placeholders
                    logger.warning(f"Unknown function placeholder: {match.group(0)}")
                    return match.group(0) # Return the unknown placeholder itself
            else:
                 # Handle unknown placeholder types (neither row nor func)
                 logger.warning(f"Unknown placeholder type in template: {match.group(0)}")
                 return match.group(0) # Return the placeholder itself

        # Use re.sub with the handler function to replace all occurrences in the string
        return placeholder_pattern.sub(replace_match, text)
    # --- End of inner replacement function ---

    # --- Main logic for traversing template data ---
    # Process strings using the inner function
    if isinstance(template_data, str):
        return perform_replace(template_data)
    # Recursively process dictionaries
    elif isinstance(template_data, dict):
        return {
            key: replace_placeholders(value, row_data, current_row_next_id)
            for key, value in template_data.items()
        }
    # Recursively process lists
    elif isinstance(template_data, list):
        return [
            replace_placeholders(item, row_data, current_row_next_id)
            for item in template_data
        ]
    # Return numbers, booleans, None, etc., directly without modification
    else:
        return template_data


# --- Backend Routes ---

@template_bp.route('/') # Route relative to the blueprint prefix ('/templates/')
def template_manager_page():
    """
    Renders the main template management UI page using an external HTML file.
    Determines the correct 'Back' link URL based on session data.
    """
    logger.info("Rendering template manager page.")

    # Determine the URL for the 'Back' link
    last_viewed = session.get('last_viewed_comparison')
    if last_viewed:
        # If a comparison page was last viewed, generate URL back to it
        try:
            back_url = url_for('ui.view_comparison', comparison_type=last_viewed)
            logger.debug(f"Setting back URL to last viewed comparison: {last_viewed}")
        except Exception as e:
            # Handle cases where url_for might fail (e.g., invalid last_viewed value)
            logger.warning(f"Could not build URL for last viewed comparison '{last_viewed}': {e}. Defaulting back URL.")
            back_url = url_for('ui.upload_config_page')
    else:
        # Otherwise, link back to the main upload/config page
        back_url = url_for('ui.upload_config_page')
        logger.debug("Setting back URL to upload/config page (no last viewed comparison found in session).")

    # Render the template_manager.html file located in the main 'templates' folder
    # Pass the calculated back_url to the template context
    return render_template('template_manager.html', back_url=back_url)


# --- API Endpoints for Templates ---

@template_bp.route('/list', methods=['GET'])
def list_templates():
    """
    API endpoint to list available .json template filenames from the template directory.

    Returns:
        JSON list of filenames on success.
        JSON error object on failure.
    """
    logger.info("Request received to list templates.")
    try:
        # Ensure the template directory exists, create if not
        if not os.path.exists(TEMPLATE_DIR):
            os.makedirs(TEMPLATE_DIR)
            logger.info(f"Created template directory: {TEMPLATE_DIR}")
            return jsonify([]) # Return empty list if directory was just created

        # List files ending with .json (case-insensitive) that are actual files
        files = [
            f for f in os.listdir(TEMPLATE_DIR)
            if f.lower().endswith('.json') and os.path.isfile(os.path.join(TEMPLATE_DIR, f))
        ]
        logger.debug(f"Found template files: {files}")
        # Return the sorted list of filenames as JSON
        return jsonify(sorted(files))
    except Exception as e:
        logger.error(f"Error listing templates in {TEMPLATE_DIR}: {e}", exc_info=True)
        # Return a server error response
        return jsonify({"error": "Failed to list templates"}), 500

@template_bp.route('/get/<path:filename>', methods=['GET'])
def get_template(filename):
    """
    API endpoint to retrieve the JSON content of a specific template file.

    Args:
        filename: The name of the template file (including .json extension).

    Returns:
        JSON content of the file on success.
        JSON error object or aborts on failure (400, 404, 500).
    """
    logger.info(f"Request received to get template: {filename}")
    # Basic security check to prevent path traversal
    if '..' in filename or filename.startswith('/'):
        logger.warning(f"Attempted access to invalid path: {filename}")
        abort(400, description="Invalid filename.") # Bad request

    try:
        filepath = os.path.join(TEMPLATE_DIR, filename)
        # Check if the file exists and is actually a file
        if not os.path.exists(filepath) or not os.path.isfile(filepath):
             logger.warning(f"Template file not found: {filepath}")
             abort(404, description="Template not found.") # Not found

        # Read and parse the JSON file content
        with open(filepath, 'r', encoding='utf-8') as f:
            content = json.load(f)
        logger.debug(f"Successfully read template content from: {filepath}")
        return jsonify(content) # Return JSON content
    except json.JSONDecodeError:
        # Handle case where the file is not valid JSON
        logger.error(f"Invalid JSON in template file: {filepath}")
        return jsonify({"error": "Template file contains invalid JSON."}), 500 # Server error
    except Exception as e:
        # Handle other file reading errors
        logger.error(f"Error reading template file {filename}: {e}", exc_info=True)
        return jsonify({"error": "Failed to read template file."}), 500 # Server error

@template_bp.route('/save', methods=['POST'])
def save_template():
    """
    API endpoint to save JSON content to a template file.
    Creates a new file or overwrites an existing one.

    Expects JSON payload: {"filename": "my_template.json", "content": { ... }}

    Returns:
        JSON success message on success.
        JSON error object or aborts on failure (400, 500).
    """
    logger.info("Request received to save template.")
    try:
        # Get JSON data from the request body
        data = request.get_json()
        if not data or 'filename' not in data or 'content' not in data:
            logger.warning("Save template request missing filename or content.")
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
             logger.warning(f"Attempted save with invalid filename: {filename}")
             abort(400, description="Invalid characters or path in filename.")

        # Ensure the template directory exists
        if not os.path.exists(TEMPLATE_DIR):
            os.makedirs(TEMPLATE_DIR)
            logger.info(f"Created template directory during save: {TEMPLATE_DIR}")

        filepath = os.path.join(TEMPLATE_DIR, filename)
        base_name = filename.replace('.json', '')

        # Check if overwriting or creating new to provide accurate logging/message
        is_update = os.path.exists(filepath)
        action = "Updated" if is_update else "Created"

        # Write the JSON content to the file, pretty-printed with indent=2
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(content, f, indent=2)

        logger.info(f"{action} template file: {filepath}")
        # Return success message with appropriate HTTP status code
        return jsonify({"message": f"Template '{base_name}' saved successfully."}), 200 if is_update else 201 # OK or Created

    except Exception as e:
        # Handle file writing errors or other unexpected issues
        logger.error(f"Error saving template file: {e}", exc_info=True)
        return jsonify({"error": "Failed to save template file."}), 500 # Server error

# Allow both DELETE and POST for deletion for flexibility (browsers sometimes block DELETE)
@template_bp.route('/delete/<path:filename>', methods=['DELETE', 'POST'])
def delete_template(filename):
    """
    API endpoint to delete a specific template file.

    Args:
        filename: The name of the template file to delete.

    Returns:
        JSON success message on success.
        JSON error object or aborts on failure (400, 404, 500).
    """
    logger.info(f"Request received to delete template: {filename}")
    # Basic security check
    if '..' in filename or filename.startswith('/'):
        logger.warning(f"Attempted delete with invalid path: {filename}")
        abort(400, description="Invalid filename.") # Bad request
    try:
        filepath = os.path.join(TEMPLATE_DIR, filename)
        base_name = filename.replace('.json', '')

        # Check if file exists before attempting deletion
        if os.path.exists(filepath) and os.path.isfile(filepath):
            os.remove(filepath)
            logger.info(f"Deleted template file: {filepath}")
            return jsonify({"message": f"Template '{base_name}' deleted successfully."}), 200 # OK
        else:
            # File not found
            logger.warning(f"Attempted to delete non-existent template: {filepath}")
            abort(404, description="Template not found.") # Not found

    except Exception as e:
        # Handle file deletion errors
        logger.error(f"Error deleting template file {filename}: {e}", exc_info=True)
        return jsonify({"error": "Failed to delete template file."}), 500 # Server error


@template_bp.route('/proxy_api_fetch', methods=['POST'])
def proxy_api_fetch():
    """
    Server-side proxy to fetch data from a user-provided URL.
    Helps avoid potential CORS (Cross-Origin Resource Sharing) issues
    if the target API doesn't allow direct browser requests.

    Expects JSON payload: {"url": "http://target-api.com/data"}

    Returns:
        JSON containing the fetched data {"data": ...} on success.
        JSON error object on failure.
    """
    logger.info("Request received for proxy API fetch.")
    try:
        data = request.get_json()
        url = data.get('url')
        if not url:
            logger.warning("Proxy fetch request missing URL.")
            return jsonify({"error": "URL is required."}), 400 # Bad request

        # Basic validation could be added here (e.g., check for http/https)
        logger.info(f"Proxy fetching reference API: {url}")

        # Make the request to the target URL
        # Consider adding headers if needed by the target API
        response = requests.get(url, timeout=10) # Use a reasonable timeout
        response.raise_for_status() # Raise HTTPError for bad status codes (4xx, 5xx)

        # Attempt to parse JSON, but return raw text if it fails
        try:
            api_data = response.json()
            logger.debug("Proxy fetch successful, parsed JSON response.")
        except requests.exceptions.JSONDecodeError:
            api_data = response.text
            logger.debug("Proxy fetch successful, but response was not valid JSON, returning as text.")

        # Return the fetched data (JSON or text) nested under a 'data' key
        return jsonify({"data": api_data}), 200 # OK

    except requests.exceptions.Timeout:
        logger.warning(f"Proxy fetch timed out for URL: {url}")
        return jsonify({"error": "Request to target API timed out."}), 504 # Gateway Timeout
    except requests.exceptions.RequestException as e:
        # Handle connection errors, invalid URLs, non-2xx status codes, etc.
        logger.error(f"Proxy fetch failed for URL {url}: {e}")
        return jsonify({"error": f"Failed to fetch data from the target API: {e}"}), 502 # Bad Gateway
    except Exception as e:
        # Catch any other unexpected errors
        logger.error(f"Unexpected error during proxy fetch: {e}", exc_info=True)
        return jsonify({"error": "An internal server error occurred during proxy fetch."}), 500 # Server error

