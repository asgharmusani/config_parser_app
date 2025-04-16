# -*- coding: utf-8 -*-
"""
Flask Blueprint for managing configuration templates (.json files).
Handles listing, viewing, saving, deleting, and fetching reference API data.
"""

import os
import json
import logging
import datetime
import requests
from flask import Blueprint, request, jsonify, render_template_string, abort

# --- Constants ---
TEMPLATE_DIR = './config_templates/'
TEMPLATE_PREFIX = 'RE_Config_'
LOG_FILE_UI = 'ui_viewer.log' # Assuming shared log file

# --- Logging ---
# Use the root logger configured in ui_viewer.py
logger = logging.getLogger()

# --- Blueprint Definition ---
template_bp = Blueprint('templates', __name__, template_folder=None) # No separate template folder

# --- HTML Template for Template Manager (Added fetchReferenceApi JS function) ---
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
        .json-display {
            background-color: #f3f4f6; /* gray-100 */
            border: 1px solid #d1d5db; /* gray-300 */
            padding: 1rem;
            border-radius: 0.375rem; /* rounded-md */
            max-height: 300px;
            overflow-y: auto;
            white-space: pre-wrap; /* Allow wrapping */
            word-wrap: break-word; /* Break long words */
            font-family: monospace;
            font-size: 0.875rem; /* text-sm */
        }
        textarea {
            font-family: monospace;
            font-size: 0.875rem; /* text-sm */
            min-height: 250px;
        }
    </style>
</head>
<body class="bg-gray-100 font-sans">
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">
        <h1 class="text-3xl font-bold text-gray-800 mb-4">Configuration Template Manager</h1>

        {# Link back to main viewer #}
         <div class="mb-6">
            <a href="{{ url_for('index') }}" class="text-indigo-600 hover:text-indigo-800 text-sm">&larr; Back to Comparison Viewer</a>
        </div>

        {# Display messages (success/error) #}
        <div id="messageArea" class="mb-4"></div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">

            {# Column 1: Template Lister #}
            <div class="md:col-span-1 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Existing Templates</h2>
                <ul id="templateList" class="list-none space-y-2">
                    {# Template list items will be added here by JS #}
                     <li class="text-gray-500 italic">Loading...</li>
                </ul>
            </div>

            {# Column 2: Template Editor/Creator #}
            <div class="md:col-span-2 bg-white p-4 shadow rounded-lg">
                <h2 class="text-xl font-semibold mb-3 border-b pb-2">Create / Edit Template</h2>

                {# Optional: Reference API Fetcher #}
                <div class="mb-4 p-3 border rounded bg-gray-50">
                    <h3 class="text-sm font-medium text-gray-600 mb-2">Reference API Data (Optional)</h3>
                    <div class="flex items-center space-x-2 mb-2">
                        <input type="text" id="refApiUrl" placeholder="Enter API URL (e.g., Agent Group URL)"
                               class="flex-grow border border-gray-300 rounded-md p-1.5 text-sm focus:ring-indigo-500 focus:border-indigo-500">
                        {# --- Button calling the function --- #}
                        <button onclick="fetchReferenceApi()" class="bg-blue-500 hover:bg-blue-600 text-white px-3 py-1.5 rounded-md text-sm">Fetch Example</button>
                    </div>
                    <div id="refApiResult" class="json-display text-xs" style="display: none;"></div>
                    <div id="refApiError" class="text-red-600 text-xs mt-1"></div>
                </div>

                {# Template Definition Area #}
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

    <script>
      lucide.createIcons();

      const templateListEl = document.getElementById('templateList');
      const templateContentEl = document.getElementById('templateContent');
      const templateNameEl = document.getElementById('templateName');
      const messageAreaEl = document.getElementById('messageArea');
      const refApiUrlEl = document.getElementById('refApiUrl');
      const refApiResultEl = document.getElementById('refApiResult');
      const refApiErrorEl = document.getElementById('refApiError');

      let currentEditingTemplate = null; // Track which template is being edited

       // --- Message Display ---
      function showMessage(text, isError = false) {
          messageAreaEl.textContent = text;
          messageAreaEl.className = `mb-4 p-3 rounded-md text-sm ${isError ? 'bg-red-100 text-red-700 border border-red-300' : 'bg-green-100 text-green-700 border border-green-300'}`;
          setTimeout(() => { messageAreaEl.textContent = ''; messageAreaEl.className = 'mb-4'; }, 5000);
      }

      // --- Template Loading ---
      async function loadTemplateList() {
          templateListEl.innerHTML = '<li class="text-gray-500 italic">Loading...</li>';
          try {
              const response = await fetch('{{ url_for("templates.list_templates") }}');
              if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
              const templates = await response.json();

              templateListEl.innerHTML = '';
              if (templates.length === 0) {
                  templateListEl.innerHTML = '<li class="text-gray-500 italic">No templates found.</li>';
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
                  lucide.createIcons();
              }
          } catch (error) {
              console.error('Error loading template list:', error);
              templateListEl.innerHTML = '<li class="text-red-600 italic">Error loading templates.</li>';
              showMessage('Failed to load template list.', true);
          }
      }

      async function loadTemplateContent(filename) {
          currentEditingTemplate = filename;
          const baseName = filename.replace('.json', '');
          templateNameEl.value = baseName;
          templateContentEl.value = 'Loading...';
           messageAreaEl.textContent = '';
          try {
              const response = await fetch(`{{ url_for("templates.get_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", filename));
              if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
              const content = await response.json();
              templateContentEl.value = JSON.stringify(content, null, 2); // Pretty print
          } catch (error) {
              console.error(`Error loading template ${filename}:`, error);
              templateContentEl.value = `Error loading template: ${error.message}`;
              showMessage(`Failed to load template '${baseName}'.`, true);
              currentEditingTemplate = null;
          }
      }

      // --- Template Saving ---
      async function saveTemplate() {
          const name = templateNameEl.value.trim();
          const content = templateContentEl.value.trim();
           messageAreaEl.textContent = '';

          if (!name) { showMessage('Template name cannot be empty.', true); templateNameEl.focus(); return; }
          if (/[\\/]/.test(name)) { showMessage('Template name contains invalid characters (\\ or /).', true); templateNameEl.focus(); return; }

          let jsonData;
          try { jsonData = JSON.parse(content); }
          catch (error) { showMessage(`Invalid JSON format: ${error.message}`, true); templateContentEl.focus(); return; }

          const filename = name + '.json';

          try {
              const response = await fetch('{{ url_for("templates.save_template") }}', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ filename: filename, content: jsonData })
              });
              const result = await response.json();
              if (response.ok) { showMessage(result.message || 'Template saved successfully.'); loadTemplateList(); }
              else { throw new Error(result.error || `HTTP error! status: ${response.status}`); }
          } catch (error) { console.error('Error saving template:', error); showMessage(`Error saving template: ${error.message}`, true); }
      }

      // --- Template Deletion ---
      async function deleteTemplate(filename) {
          const baseName = filename.replace('.json', '');
          if (!confirm(`Are you sure you want to delete the template "${baseName}"?`)) return;
           messageAreaEl.textContent = '';

          try {
              const response = await fetch(`{{ url_for("templates.delete_template", filename="PLACEHOLDER") }}`.replace("PLACEHOLDER", filename), {
                  method: 'DELETE'
              });
              const result = await response.json();
              if (response.ok) {
                  showMessage(result.message || 'Template deleted successfully.');
                  loadTemplateList();
                  if (currentEditingTemplate === filename) { templateNameEl.value = ''; templateContentEl.value = ''; currentEditingTemplate = null; }
              } else { throw new Error(result.error || `HTTP error! status: ${response.status}`); }
          } catch (error) { console.error(`Error deleting template ${filename}:`, error); showMessage(`Error deleting template '${baseName}': ${error.message}`, true); }
      }

      // --- Reference API Fetcher ---
      // --- MODIFICATION START: Added the missing function ---
       async function fetchReferenceApi() {
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
                // Use the server-side proxy endpoint
                const response = await fetch('{{ url_for("templates.proxy_api_fetch") }}', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: url })
                });
                const result = await response.json(); // Expecting {data: ...} or {error: ...}

                if (response.ok && result.data) {
                    refApiResultEl.textContent = JSON.stringify(result.data, null, 2); // Pretty print result
                } else {
                     throw new Error(result.error || `HTTP error! status: ${response.status}`);
                }
            } catch (error) {
                 console.error('Error fetching reference API:', error);
                 refApiResultEl.textContent = '';
                 refApiResultEl.style.display = 'none';
                 refApiErrorEl.textContent = `Error fetching API: ${error.message}`;
                 showMessage(`Failed to fetch reference API: ${error.message}`, true); // Also show in main message area
            }
       }
       // --- MODIFICATION END ---


      // --- Initial Load ---
      document.addEventListener('DOMContentLoaded', loadTemplateList);

    </script>
</body>
</html>
"""

# --- Backend Routes ---

@template_bp.route('/templates')
def template_manager_page():
    """Renders the main template management page."""
    return render_template_string(TEMPLATE_MANAGER_HTML)

@template_bp.route('/templates/list', methods=['GET'])
def list_templates():
    """Returns a list of .json template filenames."""
    try:
        if not os.path.exists(TEMPLATE_DIR):
            os.makedirs(TEMPLATE_DIR); logger.info(f"Created template directory: {TEMPLATE_DIR}"); return jsonify([])
        files = [f for f in os.listdir(TEMPLATE_DIR) if f.lower().endswith('.json') and os.path.isfile(os.path.join(TEMPLATE_DIR, f))]
        logger.debug(f"Found template files: {files}")
        return jsonify(sorted(files))
    except Exception as e:
        logger.error(f"Error listing templates in {TEMPLATE_DIR}: {e}", exc_info=True)
        return jsonify({"error": "Failed to list templates"}), 500

@template_bp.route('/templates/get/<path:filename>', methods=['GET'])
def get_template(filename):
    """Returns the JSON content of a specific template file."""
    if '..' in filename or filename.startswith('/'): logger.warning(f"Attempted access to invalid path: {filename}"); abort(400, description="Invalid filename.")
    try:
        filepath = os.path.join(TEMPLATE_DIR, filename)
        if not os.path.exists(filepath) or not os.path.isfile(filepath): logger.warning(f"Template file not found: {filepath}"); abort(404, description="Template not found.")
        with open(filepath, 'r', encoding='utf-8') as f: content = json.load(f)
        logger.debug(f"Read template content from: {filepath}")
        return jsonify(content)
    except json.JSONDecodeError: logger.error(f"Invalid JSON in template file: {filepath}"); return jsonify({"error": "Template file contains invalid JSON."}), 500
    except Exception as e: logger.error(f"Error reading template file {filename}: {e}", exc_info=True); return jsonify({"error": "Failed to read template file."}), 500

@template_bp.route('/templates/save', methods=['POST'])
def save_template():
    """Saves JSON content to a template file."""
    try:
        data = request.get_json();
        if not data or 'filename' not in data or 'content' not in data: abort(400, description="Missing filename or content.")
        filename = data['filename']; content = data['content']
        if not filename.lower().endswith('.json'): abort(400, description="Filename must end with .json")
        if '..' in filename or filename.startswith('/') or os.path.dirname(filename): abort(400, description="Invalid characters or path in filename.")
        if not os.path.exists(TEMPLATE_DIR): os.makedirs(TEMPLATE_DIR)
        filepath = os.path.join(TEMPLATE_DIR, filename); base_name = filename.replace('.json', '')
        is_update = os.path.exists(filepath); action = "Updated" if is_update else "Created"
        with open(filepath, 'w', encoding='utf-8') as f: json.dump(content, f, indent=2)
        logger.info(f"{action} template file: {filepath}")
        return jsonify({"message": f"Template '{base_name}' saved successfully."}), 200 if is_update else 201
    except Exception as e: logger.error(f"Error saving template file: {e}", exc_info=True); return jsonify({"error": "Failed to save template file."}), 500

@template_bp.route('/templates/delete/<path:filename>', methods=['DELETE', 'POST']) # Allow POST for simplicity
def delete_template(filename):
    """Deletes a specific template file."""
    if '..' in filename or filename.startswith('/'): logger.warning(f"Attempted delete with invalid path: {filename}"); abort(400, description="Invalid filename.")
    try:
        filepath = os.path.join(TEMPLATE_DIR, filename); base_name = filename.replace('.json', '')
        if os.path.exists(filepath) and os.path.isfile(filepath):
            os.remove(filepath); logger.info(f"Deleted template file: {filepath}")
            return jsonify({"message": f"Template '{base_name}' deleted successfully."}), 200
        else: logger.warning(f"Attempted to delete non-existent template: {filepath}"); abort(404, description="Template not found.")
    except Exception as e: logger.error(f"Error deleting template file {filename}: {e}", exc_info=True); return jsonify({"error": "Failed to delete template file."}), 500


@template_bp.route('/templates/proxy_api_fetch', methods=['POST'])
def proxy_api_fetch():
    """Server-side proxy to fetch data from a given URL to avoid CORS."""
    try:
        data = request.get_json(); url = data.get('url')
        if not url: return jsonify({"error": "URL is required."}), 400
        logger.info(f"Proxy fetching reference API: {url}")
        response = requests.get(url, timeout=10); response.raise_for_status()
        try: api_data = response.json()
        except requests.exceptions.JSONDecodeError: api_data = response.text
        return jsonify({"data": api_data}), 200
    except requests.exceptions.Timeout: logger.warning(f"Proxy fetch timed out for URL: {url}"); return jsonify({"error": "Request timed out."}), 504
    except requests.exceptions.RequestException as e: logger.error(f"Proxy fetch failed for URL {url}: {e}"); return jsonify({"error": f"Failed to fetch API: {e}"}), 502
    except Exception as e: logger.error(f"Unexpected error during proxy fetch: {e}", exc_info=True); return jsonify({"error": "An internal error occurred."}), 500

