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
"""

import os
import json
import logging
import datetime
import requests
from flask import Blueprint, request, jsonify, render_template_string, abort

# --- Constants ---
TEMPLATE_DIR = './config_templates/' # Directory where JSON templates are stored
TEMPLATE_PREFIX = 'RE_Config_'       # Optional prefix for template names (not strictly enforced here)
LOG_FILE_UI = 'ui_viewer.log'        # Assuming shared log file with main app

# --- Logging ---
# Use the root logger configured in the main app (ui_viewer.py)
logger = logging.getLogger()

# --- Blueprint Definition ---
# Create a Blueprint named 'templates'. The main app (ui_viewer.py) will register this.
# No separate template folder is specified; the HTML is embedded in this file.
template_bp = Blueprint('templates', __name__, template_folder=None)

# --- HTML Template for Template Manager ---
# This HTML structure defines the UI for the '/templates' page.
TEMPLATE_MANAGER_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Configuration Templates</title>
    {# Load Tailwind CSS via CDN for styling #}
    <script src="https://cdn.tailwindcss.com?plugins=typography"></script>
    {# Load Lucide icons library #}
    <script src="https://unpkg.com/lucide@latest/dist/umd/lucide.js"></script>
     <style>
        /* Basic styling for elements */
        .json-display { /* Style for displaying JSON content */
            background-color: #f3f4f6; /* gray-100 */
            border: 1px solid #d1d5db; /* gray-300 */
            padding: 1rem;
            border-radius: 0.375rem; /* rounded-md */
            max-height: 300px; /* Limit height and allow scrolling */
            overflow-y: auto;
            white-space: pre-wrap; /* Allow wrapping */
            word-wrap: break-word; /* Break long words */
            font-family: monospace; /* Use monospace font for code */
            font-size: 0.875rem; /* text-sm */
        }
        textarea { /* Style for the template editor textarea */
            font-family: monospace;
            font-size: 0.875rem; /* text-sm */
            min-height: 250px; /* Set a minimum height */
        }
    </style>
</head>
<body class="bg-gray-100 font-sans">
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">
        <h1 class="text-3xl font-bold text-gray-800 mb-4">Configuration Template Manager</h1>

        {# Link back to main comparison viewer page #}
         <div class="mb-6">
            <a href="{{ url_for('index') }}" class="text-indigo-600 hover:text-indigo-800 text-sm">&larr; Back to Comparison Viewer</a>
        </div>

        {# Area for displaying success/error messages from actions #}
        <div id="messageArea" class="mb-4"></div>

        {# Main grid layout: Lister | Editor #}
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">

            {# Column 1: Template Lister #}
            <div class="md:col-span-1 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Existing Templates</h2>
                {# List element populated by JavaScript #}
                <ul id="templateList" class="list-none space-y-2">
                     <li class="text-gray-500 italic">Loading...</li>
                </ul>
            </div>

            {# Column 2: Template Editor/Creator #}
            <div class="md:col-span-2 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Create / Edit Template</h2>

                {# Optional: Reference API Fetcher Section #}
                <div class="mb-4 p-3 border rounded bg-gray-50">
                    <h3 class="text-sm font-medium text-gray-600 mb-2">Reference API Data (Optional)</h3>
                    <div class="flex items-center space-x-2 mb-2">
                        <input type="text" id="refApiUrl" placeholder="Enter API URL (e.g., Agent Group URL)"
                               class="flex-grow border border-gray-300 rounded-md p-1.5 text-sm focus:ring-indigo-500 focus:border-indigo-500">
                        <button onclick="fetchReferenceApi()" class="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1.5 rounded-md text-sm">Fetch Example</button>
                    </div>
                    {# Areas to display fetched data or errors #}
                    <div id="refApiResult" class="json-display text-xs" style="display: none;"></div>
                    <div id="refApiError" class="text-red-600 text-xs mt-1"></div>
                </div>

                {# Template Definition Text Area #}
                <div class="mb-4">
                    <label for="templateContent" class="block text-sm font-medium text-gray-700 mb-1">Template JSON Structure:</label>
                    <textarea id="templateContent" rows="10" placeholder='{\n  "type": "agentGroup",\n  "key": "{row.Concatenated Key}",\n  "unNormalizedExpression": "{row.Expression}",\n  "IdealExpression": "{row.Ideal Expression}"\n}'
                              class="w-full border border-gray-300 rounded-md p-2 focus:ring-indigo-500 focus:border-indigo-500"></textarea>
                    <p class="text-xs text-gray-500 mt-1">Define the target JSON. Use placeholders like <code>{row.ColumnName}</code> to map data from selected rows (case-sensitive, matches keys from Excel reader).</p>
                </div>

                {# Saving Section #}
                <div class="flex items-center space-x-3">
                     <label for="templateName" class="text-sm font-medium text-gray-700">Save as:</label>
                     <input type="text" id="templateName" placeholder="Template Name (e.g., RE_Config_MySetup)"
                            class="flex-grow border border-gray-300 rounded-md p-1.5 text-sm focus:ring-indigo-500 focus:border-indigo-500">
                     <button onclick="saveTemplate()" class="bg-green-600 hover:bg-green-700 text-white px-4 py-1.5 rounded-md text-sm">Save Template</button>
                </div>
                <p class="text-xs text-gray-500 mt-1">Filename will be saved as <code>./config_templates/&lt;Template Name&gt;.json</code>.</p>

            </div>
        </div> {# End Grid #}
    </div> {# End Container #}

    {# --- JavaScript Section --- #}
    <script>
      // Initialize Lucide icons used in the template
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
      let currentEditingTemplate = null; // Track which template filename is currently loaded in the editor

       // --- Utility: Message Display ---
      function showMessage(text, isError = false) {
          // Displays a feedback message (success or error) in the message area.
          messageAreaEl.textContent = text;
          // Apply appropriate styling based on whether it's an error
          messageAreaEl.className = `mb-4 p-3 rounded-md text-sm ${isError ? 'bg-red-100 text-red-700 border border-red-300' : 'bg-green-100 text-green-700 border border-green-300'}`;
          // Auto-hide message after 5 seconds for better UX
          setTimeout(() => { messageAreaEl.textContent = ''; messageAreaEl.className = 'mb-4'; }, 5000);
      }

      // --- Template Loading Functions ---
      async function loadTemplateList() {
          // Fetches the list of template filenames from the backend and populates the UI list.
          templateListEl.innerHTML = '<li class="text-gray-500 italic">Loading...</li>'; // Show loading indicator
          try {
              // Fetch list from the backend endpoint
              const response = await fetch('{{ url_for("templates.list_templates") }}');
              if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
              const templates = await response.json(); // Expecting list of filenames

              templateListEl.innerHTML = ''; // Clear loading/previous list
              if (templates.length === 0) {
                  templateListEl.innerHTML = '<li class="text-gray-500 italic">No templates found. Create one!</li>';
              } else {
                  // Create list items for each template with view/edit and delete buttons
                  templates.forEach(filename => {
                      const li = document.createElement('li');
                      li.className = 'flex justify-between items-center text-sm border-b pb-1';
                      const baseName = filename.replace('.json', ''); // Display name without extension

                      // Template name (clickable to load)
                      const nameSpan = document.createElement('span');
                      nameSpan.textContent = baseName;
                      nameSpan.className = 'cursor-pointer hover:text-indigo-600';
                      nameSpan.onclick = () => loadTemplateContent(filename); // Load content on click

                      // Action buttons container
                      const actionsDiv = document.createElement('div');
                      actionsDiv.className = 'space-x-2';

                      // View/Edit Button
                      const viewButton = document.createElement('button');
                      viewButton.innerHTML = '<i data-lucide="eye" class="w-4 h-4 text-blue-500 hover:text-blue-700"></i>';
                      viewButton.title = 'View/Edit';
                      viewButton.onclick = () => loadTemplateContent(filename); // Load content on click

                      // Delete Button
                      const deleteButton = document.createElement('button');
                      deleteButton.innerHTML = '<i data-lucide="trash-2" class="w-4 h-4 text-red-500 hover:text-red-700"></i>';
                      deleteButton.title = 'Delete';
                      deleteButton.onclick = () => deleteTemplate(filename); // Trigger delete confirmation

                      // Assemble the list item
                      actionsDiv.appendChild(viewButton);
                      actionsDiv.appendChild(deleteButton);
                      li.appendChild(nameSpan);
                      li.appendChild(actionsDiv);
                      templateListEl.appendChild(li);
                  });
                  lucide.createIcons(); // Re-render Lucide icons after adding them
              }
          } catch (error) {
              // Handle errors during template list loading
              console.error('Error loading template list:', error);
              templateListEl.innerHTML = '<li class="text-red-600 italic">Error loading templates.</li>';
              showMessage('Failed to load template list.', true);
          }
      }

      async function loadTemplateContent(filename) {
          // Fetches the content of a specific template file and displays it in the editor.
          currentEditingTemplate = filename; // Track the file being edited
          const baseName = filename.replace('.json', '');
          templateNameEl.value = baseName; // Populate the name input field
          templateContentEl.value = 'Loading...'; // Indicate loading in textarea
           messageAreaEl.textContent = ''; // Clear previous messages
          try {
              // Construct URL safely using placeholder replacement
              const url = `{{ url_for("templates.get_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", encodeURIComponent(filename));
              const response = await fetch(url);
              if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
              const content = await response.json(); // Expecting JSON content
              // Display JSON nicely formatted in the textarea
              templateContentEl.value = JSON.stringify(content, null, 2);
          } catch (error) {
              // Handle errors loading template content
              console.error(`Error loading template ${filename}:`, error);
              templateContentEl.value = `Error loading template: ${error.message}`;
              showMessage(`Failed to load template '${baseName}'.`, true);
              currentEditingTemplate = null; // Reset editing state on error
          }
      }

      // --- Template Saving Function ---
      async function saveTemplate() {
          // Saves the content of the editor to a JSON file on the server.
          const name = templateNameEl.value.trim();
          const content = templateContentEl.value.trim();
           messageAreaEl.textContent = ''; // Clear previous messages

          // Basic filename validation
          if (!name) {
              showMessage('Template name cannot be empty.', true);
              templateNameEl.focus();
              return;
          }
           // Prevent directory traversal or invalid characters
          if (/[\\/]/.test(name)) {
             showMessage('Template name contains invalid characters (\\ or /).', true);
             templateNameEl.focus();
             return;
          }

          // Validate JSON content before sending
          let jsonData;
          try {
              jsonData = JSON.parse(content);
          } catch (error) {
              showMessage(`Invalid JSON format in editor: ${error.message}`, true);
              templateContentEl.focus();
              return;
          }

          // Construct filename
          const filename = name + '.json';

          try {
              // Send POST request to the save endpoint
              const response = await fetch('{{ url_for("templates.save_template") }}', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  // Send filename and the parsed JSON content
                  body: JSON.stringify({ filename: filename, content: jsonData })
              });
              const result = await response.json(); // Expecting {message: ...} or {error: ...}

              if (response.ok) {
                  showMessage(result.message || 'Template saved successfully.');
                  loadTemplateList(); // Refresh the list of templates
                  // Keep content in editor for further edits or clear it:
                  // loadTemplateContent(filename); // Option: reload saved content
              } else {
                  // Handle errors reported by the backend
                  throw new Error(result.error || `HTTP error! status: ${response.status}`);
              }
          } catch (error) {
              // Handle network errors or errors parsing the response
              console.error('Error saving template:', error);
              showMessage(`Error saving template: ${error.message}`, true);
          }
      }

      // --- Template Deletion Function ---
      async function deleteTemplate(filename) {
          // Deletes a template file from the server after confirmation.
          const baseName = filename.replace('.json', '');
          // Confirm deletion with the user
          if (!confirm(`Are you sure you want to delete the template "${baseName}"? This action cannot be undone.`)) {
              return;
          }
           messageAreaEl.textContent = ''; // Clear previous messages

          try {
               // Construct URL safely using placeholder replacement
              const url = `{{ url_for("templates.delete_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", encodeURIComponent(filename));
              // Send DELETE request (or POST if DELETE causes issues)
              const response = await fetch(url, {
                  method: 'DELETE'
              });
              const result = await response.json(); // Expecting {message: ...} or {error: ...}

              if (response.ok) {
                  showMessage(result.message || 'Template deleted successfully.');
                  loadTemplateList(); // Refresh the list
                  // If the deleted template was loaded in the editor, clear the editor
                  if (currentEditingTemplate === filename) {
                      templateNameEl.value = '';
                      templateContentEl.value = '';
                      currentEditingTemplate = null;
                  }
              } else {
                   // Handle errors reported by the backend
                  throw new Error(result.error || `HTTP error! status: ${response.status}`);
              }
          } catch (error) {
              // Handle network errors or errors parsing the response
              console.error(`Error deleting template ${filename}:`, error);
              showMessage(`Error deleting template '${baseName}': ${error.message}`, true);
          }
      }

      // --- Reference API Fetcher Function ---
       async function fetchReferenceApi() {
            // Fetches example data from a user-provided URL via a server-side proxy.
            const url = refApiUrlEl.value.trim();
            // Clear previous results/errors
            refApiResultEl.style.display = 'none';
            refApiResultEl.textContent = '';
            refApiErrorEl.textContent = '';

            if (!url) {
                refApiErrorEl.textContent = 'Please enter an API URL.';
                return;
            }

            // Show fetching indicator
            refApiResultEl.textContent = 'Fetching...';
            refApiResultEl.style.display = 'block';

            try {
                // Call the backend proxy endpoint
                const response = await fetch('{{ url_for("templates.proxy_api_fetch") }}', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    // Send the target URL in the request body
                    body: JSON.stringify({ url: url })
                });
                const result = await response.json(); // Expecting {data: ...} or {error: ...}

                if (response.ok && result.data) {
                    // Display the fetched data pretty-printed
                    refApiResultEl.textContent = JSON.stringify(result.data, null, 2);
                } else {
                     // Handle errors reported by the proxy or the target API
                     throw new Error(result.error || `HTTP error! status: ${response.status}`);
                }
            } catch (error) {
                 // Handle network errors or errors parsing the response
                 console.error('Error fetching reference API:', error);
                 refApiResultEl.textContent = '';
                 refApiResultEl.style.display = 'none';
                 refApiErrorEl.textContent = `Error fetching API: ${error.message}`;
                 // Optionally show error in main message area too
                 // showMessage(`Failed to fetch reference API: ${error.message}`, true);
            }
       }


      // --- Initial Load ---
      // Load the list of templates when the page DOM is ready
      document.addEventListener('DOMContentLoaded', loadTemplateList);

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
        # Prevent path traversal and invalid characters
        if '..' in filename or filename.startswith('/') or os.path.dirname(filename) or any(c in filename for c in r'<>:"|?*'):
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

        # Write the JSON content to the file, pretty-printed
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(content, f, indent=2)

        logger.info(f"{action} template file: {filepath}")
        # Return success message
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

