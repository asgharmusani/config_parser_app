# -*- coding: utf-8 -*-
"""
Flask Blueprint for managing configuration templates (.json files).
Handles listing, viewing, saving, deleting, fetching reference API data,
and includes the placeholder replacement logic.
"""

import os
import json
import logging
import datetime
import requests
import re # Added for placeholder logic
from flask import Blueprint, request, jsonify, render_template_string, abort, current_app # Added current_app
from typing import Dict, Any, Optional, List # Added List

# --- Constants ---
TEMPLATE_DIR = './config_templates/'
LOG_FILE_UI = 'ui_viewer.log'

# --- Logging ---
logger = logging.getLogger()

# --- Blueprint Definition ---
template_bp = Blueprint('templates', __name__, template_folder=None)

# --- Helper Function: Placeholder Replacement ---
# (No changes needed in this function from previous version)
def replace_placeholders(template_data: Any, row_data: dict, current_row_next_id: Optional[int] = None) -> Any:
    """
    Recursively traverses a template structure (dict, list, or string)
    and replaces placeholders with values from row_data or the pre-generated ID.
    """
    placeholder_pattern = re.compile(r'{(\w+)\.([^}]+)}') # Captures type (row/func) and name

    def perform_replace(text: str) -> str:
        """Performs replacements on a single string."""
        if not isinstance(text, str):
            return text

        def replace_match(match):
            placeholder_type = match.group(1).lower()
            placeholder_name = match.group(2).strip()

            if placeholder_type == 'row':
                replacement = row_data.get(placeholder_name, "")
                return str(replacement)
            elif placeholder_type == 'func':
                if placeholder_name == 'next_id':
                    if current_row_next_id is not None:
                        return str(current_row_next_id)
                    else:
                        logger.warning(f"Placeholder {{func.next_id}} used but no ID provided for this row.")
                        return "{ERROR:next_id_missing}"
                else:
                    logger.warning(f"Unknown function placeholder: {match.group(0)}")
                    return match.group(0)
            else:
                 logger.warning(f"Unknown placeholder type in template: {match.group(0)}")
                 return match.group(0)

        return placeholder_pattern.sub(replace_match, text)

    if isinstance(template_data, str):
        return perform_replace(template_data)
    elif isinstance(template_data, dict):
        return {
            key: replace_placeholders(value, row_data, current_row_next_id)
            for key, value in template_data.items()
        }
    elif isinstance(template_data, list):
        return [
            replace_placeholders(item, row_data, current_row_next_id)
            for item in template_data
        ]
    else:
        return template_data


# --- HTML Template for Template Manager (Updated Examples & Help Text) ---
TEMPLATE_MANAGER_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Configuration Templates</title>
    <script src="https://cdn.tailwindcss.com?plugins=typography"></script>
    <script src="https://unpkg.com/lucide@latest/dist/umd/lucide.js"></script>
     <style>
        /* Basic styling */
        .json-display { background-color: #f3f4f6; border: 1px solid #d1d5db; padding: 1rem; border-radius: 0.375rem; max-height: 300px; overflow-y: auto; white-space: pre-wrap; word-wrap: break-word; font-family: monospace; font-size: 0.875rem; }
        textarea { font-family: monospace; font-size: 0.875rem; min-height: 300px; }
        code { background-color: #e5e7eb; padding: 0.1em 0.3em; border-radius: 0.25em; font-size: 0.9em; }
        .help-text ul { margin-top: 0.25rem; }
        .help-text li { margin-bottom: 0.25rem;}
    </style>
</head>
<body class="bg-gray-100 font-sans">
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">
        <h1 class="text-3xl font-bold text-gray-800 mb-4">Configuration Template Manager</h1>
         <div class="mb-6"> <a href="{{ url_for('index') }}" class="text-indigo-600 hover:text-indigo-800 text-sm">&larr; Back to Comparison Viewer</a> </div>
        <div id="messageArea" class="mb-4"></div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
            {# Column 1: Template Lister #}
            <div class="md:col-span-1 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Existing Templates</h2>
                <ul id="templateList" class="list-none space-y-2"> <li class="text-gray-500 italic">Loading...</li> </ul>
            </div>

            {# Column 2: Template Editor/Creator #}
            <div class="md:col-span-2 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Create / Edit Template</h2>

                {# Reference API Fetcher #}
                <div class="mb-4 p-3 border rounded bg-gray-50">
                    <h3 class="text-sm font-medium text-gray-600 mb-2">Reference API Data (Optional)</h3>
                    <div class="flex items-center space-x-2 mb-2">
                        <input type="text" id="refApiUrl" placeholder="Enter API URL (e.g., Agent Group URL)" class="flex-grow border border-gray-300 rounded-md p-1.5 text-sm focus:ring-indigo-500 focus:border-indigo-500">
                        <button onclick="fetchReferenceApi()" class="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1.5 rounded-md text-sm">Fetch Example</button>
                    </div>
                    <div id="refApiResult" class="json-display text-xs" style="display: none;"></div>
                    <div id="refApiError" class="text-red-600 text-xs mt-1"></div>
                </div>

                {# Template Definition Area #}
                <div class="mb-4">
                    <label for="templateContent" class="block text-sm font-medium text-gray-700 mb-1">Template JSON Structure:</label>
                    {# --- MODIFICATION START: Updated placeholder example --- #}
                    <textarea id="templateContent" rows="15"
                        placeholder='{\n  "staticExample": "This is a fixed value",\n  "entityType": "agentGroup",\n  "vqName": "{row.Item}", \n  "skillExprKey": "{row.Concatenated Key}",\n  "combinedExample": "AG_{row.ID}_Suffix",\n  "needsReview": true,\n  "newIdExample": "{func.next_id}",\n  "details": {\n    "expression": "{row.Expression}",\n    "ideal": "{row.Ideal Expression}",\n    "statusFromSheet": "{row.Status}"\n  }\n}'
                        class="w-full border border-gray-300 rounded-md p-2 focus:ring-indigo-500 focus:border-indigo-500"></textarea>
                    {# --- MODIFICATION END --- #}

                    {# --- MODIFICATION START: Updated help text --- #}
                    <div class="text-xs text-gray-600 mt-1 space-y-1 help-text">
                        <p>Define the target JSON structure for database updates.</p>
                        <p><strong>Value Types:</strong></p>
                        <ul class="list-disc list-inside ml-2">
                            <li><strong>Static:</strong> Enter plain text, numbers, or booleans (e.g., <code>"active": true</code>).</li>
                            <li><strong>Row Data:</strong> Use <code>{row.ColumnName}</code>. The <code>ColumnName</code> must exactly match a key in the data read from the comparison sheet row.
                                <ul class="list-circle list-inside ml-4">
                                     <li>For VQs, Skills, VAGs sheets: <code>{row.Item}</code>, <code>{row.ID}</code>, <code>{row.Status}</code></li>
                                     <li>For Skill Exprs sheet: <code>{row.Concatenated Key}</code>, <code>{row.Expression}</code>, <code>{row.Ideal Expression}</code>, <code>{row.ID}</code>, <code>{row.Status}</code></li>
                                </ul>
                            </li>
                            <li><strong>Combined:</strong> Mix static text and row data (e.g., <code>"AG_{row.ID}"</code>, <code>"Prefix:{row.Item}"</code>).</li>
                            <li><strong>Function (Next ID):</strong> Use <code>{func.next_id}</code> to insert the next available sequential ID for the current row being processed. The same ID is used for all <code>{func.next_id}</code> placeholders within a single row's template application. The starting ID is based on the max ID found in the loaded Excel file's Metadata sheet.</li>
                        </ul>
                    </div>
                     {# --- MODIFICATION END --- #}
                </div>

                {# Saving Section #}
                <div class="flex items-center space-x-3">
                     <label for="templateName" class="text-sm font-medium text-gray-700">Save as:</label>
                     <input type="text" id="templateName" placeholder="Template Name (e.g., RE_Config_MySetup)" class="flex-grow border border-gray-300 rounded-md p-1.5 text-sm focus:ring-indigo-500 focus:border-indigo-500">
                     <button onclick="saveTemplate()" class="bg-green-600 hover:bg-green-700 text-white px-4 py-1.5 rounded-md text-sm">Save Template</button>
                </div>
                <p class="text-xs text-gray-500 mt-1">Filename will be saved as <code>./config_templates/&lt;Template Name&gt;.json</code>.</p>

            </div>
        </div> {# End Grid #}
    </div> {# End Container #}

    {# --- JavaScript Section --- #}
    <script>
      // Initialize Lucide icons
      lucide.createIcons();

      // --- DOM Element References ---
      const templateListEl = document.getElementById('templateList');
      const templateContentEl = document.getElementById('templateContent');
      const templateNameEl = document.getElementById('templateName');
      const messageAreaEl = document.getElementById('messageArea');
      const refApiUrlEl = document.getElementById('refApiUrl');
      const refApiResultEl = document.getElementById('refApiResult');
      const refApiErrorEl = document.getElementById('refApiError');

      // --- State Variable ---
      let currentEditingTemplate = null; // Track which template filename is loaded

       // --- Utility: Message Display ---
      function showMessage(text, isError = false) {
          messageAreaEl.textContent = text;
          messageAreaEl.className = `mb-4 p-3 rounded-md text-sm ${isError ? 'bg-red-100 text-red-700 border border-red-300' : 'bg-green-100 text-green-700 border border-green-300'}`;
          setTimeout(() => { messageAreaEl.textContent = ''; messageAreaEl.className = 'mb-4'; }, 5000); // Auto-hide
      }

      // --- Template Loading Functions ---
      async function loadTemplateList() {
          // Fetches and displays the list of templates.
          templateListEl.innerHTML = '<li class="text-gray-500 italic">Loading...</li>';
          try {
              const response = await fetch('{{ url_for("templates.list_templates") }}');
              if (!response.ok) {
                  throw new Error(`HTTP error! status: ${response.status}`);
              }
              const templates = await response.json();
              templateListEl.innerHTML = ''; // Clear list
              if (templates.length === 0) {
                  templateListEl.innerHTML = '<li class="text-gray-500 italic">No templates found. Create one!</li>';
              } else {
                  templates.forEach(filename => {
                      const li = document.createElement('li');
                      li.className = 'flex justify-between items-center text-sm border-b pb-1';
                      const baseName = filename.replace('.json', '');
                      const nameSpan = document.createElement('span');
                      nameSpan.textContent = baseName;
                      nameSpan.className = 'cursor-pointer hover:text-indigo-600';
                      nameSpan.onclick = () => loadTemplateContent(filename);
                      const actionsDiv = document.createElement('div');
                      actionsDiv.className = 'space-x-2';
                      const viewButton = document.createElement('button');
                      viewButton.innerHTML = '<i data-lucide="eye" class="w-4 h-4 text-blue-500 hover:text-blue-700"></i>';
                      viewButton.title = 'View/Edit';
                      viewButton.onclick = () => loadTemplateContent(filename);
                      const deleteButton = document.createElement('button');
                      deleteButton.innerHTML = '<i data-lucide="trash-2" class="w-4 h-4 text-red-500 hover:text-red-700"></i>';
                      deleteButton.title = 'Delete';
                      deleteButton.onclick = () => deleteTemplate(filename);
                      actionsDiv.appendChild(viewButton);
                      actionsDiv.appendChild(deleteButton);
                      li.appendChild(nameSpan);
                      li.appendChild(actionsDiv);
                      templateListEl.appendChild(li);
                   });
                  lucide.createIcons(); // Render icons
              }
          } catch (error) {
              console.error('Error loading template list:', error);
              templateListEl.innerHTML = '<li class="text-red-600 italic">Error loading templates.</li>';
              showMessage('Failed to load template list.', true);
          }
      }

      async function loadTemplateContent(filename) {
          // Loads the content of a selected template into the editor.
          currentEditingTemplate = filename;
          const baseName = filename.replace('.json', '');
          templateNameEl.value = baseName;
          templateContentEl.value = 'Loading...';
           messageAreaEl.textContent = ''; // Clear messages
          try {
              const url = `{{ url_for("templates.get_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", encodeURIComponent(filename));
              const response = await fetch(url);
              if (!response.ok) {
                  throw new Error(`HTTP error! status: ${response.status}`);
              }
              const content = await response.json();
              templateContentEl.value = JSON.stringify(content, null, 2); // Pretty print
          } catch (error) {
              console.error(`Error loading template ${filename}:`, error);
              templateContentEl.value = `Error loading template: ${error.message}`;
              showMessage(`Failed to load template '${baseName}'.`, true);
              currentEditingTemplate = null;
          }
      }

      // --- Template Saving Function ---
      async function saveTemplate() {
          // Saves the editor content as a template file.
          const name = templateNameEl.value.trim();
          const content = templateContentEl.value.trim();
           messageAreaEl.textContent = '';
          if (!name) {
              showMessage('Template name cannot be empty.', true);
              templateNameEl.focus();
              return;
          }
          if (/[\\/]/.test(name)) {
              showMessage('Template name contains invalid characters (\\ or /).', true);
              templateNameEl.focus();
              return;
          }
          let jsonData;
          try {
              jsonData = JSON.parse(content); // Validate JSON
          }
          catch (error) {
              showMessage(`Invalid JSON format: ${error.message}`, true);
              templateContentEl.focus();
              return;
          }
          const filename = name + '.json';
          try {
              const response = await fetch('{{ url_for("templates.save_template") }}', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ filename: filename, content: jsonData })
              });
              const result = await response.json();
              if (response.ok) {
                  showMessage(result.message || 'Template saved successfully.');
                  loadTemplateList();
              }
              else {
                  throw new Error(result.error || `HTTP error! status: ${response.status}`);
              }
          } catch (error) {
              console.error('Error saving template:', error);
              showMessage(`Error saving template: ${error.message}`, true);
          }
      }

      // --- Template Deletion Function ---
      async function deleteTemplate(filename) {
          // Deletes a template file after confirmation.
          const baseName = filename.replace('.json', '');
          if (!confirm(`Are you sure you want to delete the template "${baseName}"?`)) {
              return;
          }
           messageAreaEl.textContent = '';
          try {
              const url = `{{ url_for("templates.delete_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", encodeURIComponent(filename));
              const response = await fetch(url, { method: 'DELETE' });
              const result = await response.json();
              if (response.ok) {
                  showMessage(result.message || 'Template deleted successfully.');
                  loadTemplateList();
                  if (currentEditingTemplate === filename) {
                      templateNameEl.value = '';
                      templateContentEl.value = '';
                      currentEditingTemplate = null;
                  }
              } else {
                  throw new Error(result.error || `HTTP error! status: ${response.status}`);
              }
          } catch (error) {
              console.error(`Error deleting template ${filename}:`, error);
              showMessage(`Error deleting template '${baseName}': ${error.message}`, true);
          }
      }

      // --- Reference API Fetcher Function ---
       async function fetchReferenceApi() {
           // Fetches example API data via the backend proxy.
            const url = refApiUrlEl.value.trim();
            refApiResultEl.style.display = 'none';
            refApiResultEl.textContent = '';
            refApiErrorEl.textContent = '';
            if (!url) {
                refApiErrorEl.textContent = 'Please enter an API URL.';
                return;
            }
            refApiResultEl.textContent = 'Fetching...';
            refApiResultEl.style.display = 'block';
            try {
                const response = await fetch('{{ url_for("templates.proxy_api_fetch") }}', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: url })
                });
                const result = await response.json();
                if (response.ok && result.data) {
                    refApiResultEl.textContent = JSON.stringify(result.data, null, 2); // Pretty print
                } else {
                    throw new Error(result.error || `HTTP error! status: ${response.status}`);
                }
            } catch (error) {
                 console.error('Error fetching reference API:', error);
                 refApiResultEl.textContent = '';
                 refApiResultEl.style.display = 'none';
                 refApiErrorEl.textContent = `Error fetching API: ${error.message}`;
                 showMessage(`Failed to fetch reference API: ${error.message}`, true);
            }
       }


      // --- Initial Load ---
      document.addEventListener('DOMContentLoaded', loadTemplateList); // Load template list when page is ready

    </script>
</body>
</html>
"""

# --- Backend Routes ---

@template_bp.route('/templates')
def template_manager_page():
    """Renders the main template management UI page."""
    logger.info("Rendering template manager page.")
    # The page structure is defined in TEMPLATE_MANAGER_HTML.
    # JavaScript will fetch the dynamic data (template list).
    return render_template_string(TEMPLATE_MANAGER_HTML)

@template_bp.route('/templates/list', methods=['GET'])
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

@template_bp.route('/templates/get/<path:filename>', methods=['GET'])
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

@template_bp.route('/templates/save', methods=['POST'])
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
@template_bp.route('/templates/delete/<path:filename>', methods=['DELETE', 'POST'])
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


@template_bp.route('/templates/proxy_api_fetch', methods=['POST'])
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

