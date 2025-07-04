# -*- coding: utf-8 -*-
"""
Flask Blueprint for defining routes that render the main UI pages.
- Upload/Configuration Page
- Comparison Results Viewer Page
- Data Refresh Trigger
"""

import logging
import math
import os # Added for listing processed files
from flask import (
    Blueprint, render_template, request, redirect, url_for, current_app, flash, session # Added session
)
from typing import Optional, Tuple, List, Dict, Any # Added List, Dict, Any

# --- Constants (Defined locally for this blueprint) ---
# These constants are used for pagination and template rendering logic.
DEFAULT_PAGE_SIZE = 100
PAGE_SIZE_OPTIONS = [100, 200, 500, 1000]
COMPARISON_SUFFIX = " Comparison" # Expected suffix for comparison sheet names in Excel
SKILL_EXPR_SHEET_NAME = "Skill_exprs Comparison" # Specific sheet name for special handling
UPLOAD_FOLDER = './uploads' # Directory where uploaded and processed files are stored


# --- Logging ---
logger = logging.getLogger(__name__) # Use module-specific logger

# --- Blueprint Definition ---
# Create a Blueprint named 'ui'. The main app (app.py) will register this.
# Point to the main 'templates' folder where HTML files reside.
ui_bp = Blueprint('ui', __name__, template_folder='../templates')

# --- UI Routes ---

@ui_bp.route('/upload')
def upload_config_page():
    """
    Renders the initial page ('upload_config.html').
    This page allows users to:
    1. Upload a new original Excel file for processing.
    2. View and select from a list of previously processed Excel files.
    3. View and modify application configuration settings (from config.ini).
    """
    logger.info("Rendering Upload/Configuration page.")
    # Fetch current application configuration settings to display in the form
    # Assumes config is loaded into current_app.config['APP_SETTINGS'] by app.py
    app_config = current_app.config.get('APP_SETTINGS', {})

    # List previously processed files for the user to select from
    processed_files = []
    # Ensure upload folder exists before trying to list files
    if os.path.exists(UPLOAD_FOLDER):
        try:
            processed_files = sorted(
                # List files ending with _processed.xlsx and ensure they are files
                [
                    f for f in os.listdir(UPLOAD_FOLDER)
                    if f.endswith('_processed.xlsx') and os.path.isfile(os.path.join(UPLOAD_FOLDER, f))
                ],
                # Sort by modification time, newest first
                key=lambda f: os.path.getmtime(os.path.join(UPLOAD_FOLDER, f)),
                reverse=True
            )
            logger.debug(f"Found processed files: {processed_files}")
        except Exception as e:
            logger.error(f"Error listing processed files in {UPLOAD_FOLDER}: {e}")
            flash("Error listing previously processed files. Check logs.", "error")
    else:
        logger.warning(f"Upload folder '{UPLOAD_FOLDER}' does not exist. Cannot list processed files.")

    # Pass necessary context for base.html's navigation, even if no data is loaded
    # These are needed because upload_config.html extends base.html which uses these for nav
    available_sheets_for_nav = current_app.config.get('COMPARISON_SHEETS', [])

    # Pass config, file list, and navigation context to the template
    return render_template(
        'upload_config.html',
        config=app_config,
        processed_files=processed_files,
        # Context for base.html's navigation
        available_sheets=available_sheets_for_nav,
        comparison_suffix_for_template=COMPARISON_SUFFIX, # Pass the constant
        # Provide defaults for other nav-related vars that might be expected by base.html
        current_comparison_type=None,
        sort_by=None,
        sort_order=None,
        page_size_str=str(DEFAULT_PAGE_SIZE),
        filename=current_app.config.get('EXCEL_FILENAME') # Pass filename if available
    )


@ui_bp.route('/view/<comparison_type>')
def view_comparison(comparison_type: str):
    """
    Displays a specific comparison type sheet with pagination and sorting.
    Renders the 'results_viewer.html' template.
    Stores the viewed comparison type in the session.

    Args:
        comparison_type: The name of the comparison sheet to display
                         (e.g., "Vqs Comparison").
    """
    logger.info(f"Request to view comparison type: {comparison_type}")

    # --- Get Data and Config from App Cache ---
    filename = current_app.config.get('EXCEL_FILENAME')
    all_data = current_app.config.get('EXCEL_DATA', {})
    available_sheets = current_app.config.get('COMPARISON_SHEETS', [])
    sheet_headers_map = current_app.config.get('SHEET_HEADERS', {})
    error = None # Initialize error variable for this request

    # Check if data is loaded; if not, redirect to the upload page with a message
    if not filename or not all_data or not available_sheets or not sheet_headers_map:
        error_msg = "No comparison data loaded. Please upload/process or load a file first."
        logger.warning(error_msg)
        flash(error_msg, 'warning') # Use Flask flash messaging
        return redirect(url_for('ui.upload_config_page')) # Redirect to upload page

    # Validate requested comparison type
    if comparison_type not in all_data or comparison_type not in sheet_headers_map:
        logger.warning(f"Invalid comparison type requested or headers missing: '{comparison_type}'. Redirecting.")
        flash(f"Invalid comparison type requested: {comparison_type}", 'error')
        if available_sheets:
             return redirect(url_for('ui.view_comparison', comparison_type=available_sheets[0]))
        else:
             return redirect(url_for('ui.upload_config_page'))

    # Store last viewed page in session for the 'Back' link on template manager
    session['last_viewed_comparison'] = comparison_type
    logger.debug(f"Stored last viewed comparison in session: {comparison_type}")

    current_headers = sheet_headers_map.get(comparison_type, [])
    if not current_headers:
         logger.error(f"Headers not found for sheet: {comparison_type}. Cannot render table.")
         flash(f"Could not load headers for '{comparison_type}'.", 'error')
         return redirect(url_for('ui.upload_config_page'))

    # --- Get URL Parameters (Page, Size, Sort) ---
    try:
        page = request.args.get('page', 1, type=int)
        page_size_str = request.args.get('size', str(DEFAULT_PAGE_SIZE), type=str).lower()
        default_sort_col = current_headers[0] # Default sort by first header
        sort_by = request.args.get('sort_by', default_sort_col, type=str)
        sort_order = request.args.get('order', 'asc', type=str).lower()
    except (ValueError, IndexError): # Catch potential errors if headers are empty or params invalid
        # Fallback to defaults if query parameters are invalid type or headers missing
        logging.warning("Invalid query parameter type or missing headers during param parsing, using defaults.")
        page = 1
        page_size_str = str(DEFAULT_PAGE_SIZE)
        default_sort_col = current_headers[0] if current_headers else None
        sort_by = default_sort_col
        sort_order = 'asc'


    # --- Validate and process parameters ---
    if page < 1:
        page = 1 # Ensure page is at least 1
    if sort_order not in ['asc', 'desc']:
        sort_order = 'asc' # Default to ascending

    # Use the actual headers read from the sheet as valid sort columns
    valid_sort_columns = current_headers
    # If requested sort_by is invalid, revert to default (first header)
    if sort_by not in valid_sort_columns:
        sort_by = default_sort_col if default_sort_col else (current_headers[0] if current_headers else None)

    # Determine numeric page size
    show_all = (page_size_str == 'all')
    page_size = DEFAULT_PAGE_SIZE # Default numeric size
    if not show_all:
        try:
            requested_size = int(page_size_str)
            # Use requested size only if it's one of the predefined valid options
            if requested_size in PAGE_SIZE_OPTIONS:
                 page_size = requested_size
            # else: keep the default numeric page_size
        except ValueError:
            # If conversion fails (e.g., size=abc), reset string and use default numeric
            page_size_str = str(DEFAULT_PAGE_SIZE)
            # page_size already holds DEFAULT_PAGE_SIZE

    # --- Get Data and Sort ---
    current_sheet_data = all_data.get(comparison_type, [])
    total_items = len(current_sheet_data)
    sorted_data = current_sheet_data # Default to unsorted if sorting fails or not applicable

    if total_items > 0 and sort_by: # Only sort if there's data and a valid column to sort by
        reverse_sort = (sort_order == 'desc')

        # Sort key function (handles None, tries numeric for ID, defaults to string)
        def sort_key(item_row_dict: Dict[str, Any]) -> Tuple:
            """Generate a sort key for Python's sort, handling None and basic types."""
            value = item_row_dict.get(sort_by) # Get the value for the column we're sorting by

            if value is None:
                # Place None values consistently (e.g., at the end when ascending)
                return (1, float('inf')) if sort_order == 'asc' else (0, float('-inf'))
            try:
                # Try numeric sort for 'ID' column (or similar) if possible
                # Check header name case-insensitively for flexibility
                if sort_by.upper() == 'ID' or sort_by.upper() == 'ID (FROM API)':
                    try:
                        # Group numbers first
                        return (0, float(value))
                    except (ValueError, TypeError):
                         # Treat non-numeric IDs as strings, group after numbers
                        return (1, str(value).lower())
                # Default: Case-insensitive string sort
                return (0, str(value).lower())
            except Exception as e:
                # Fallback for any unexpected error during value processing
                logging.warning(f"Could not process value '{value}' for sorting by '{sort_by}': {e}")
                 # Group these problematic values last
                return (2, str(value).lower())

        # Perform the sort
        try:
            sorted_data = sorted(current_sheet_data, key=sort_key, reverse=reverse_sort)
        except Exception as sort_e:
            # Handle potential errors during sorting (e.g., complex type issues)
            logging.error(f"Error during sorting data for '{comparison_type}': {sort_e}", exc_info=True)
            error = f"Error sorting data by {sort_by}. Displaying unsorted." # Inform user via error var
            # sorted_data remains the original current_sheet_data (unsorted)

    # --- Pagination ---
    page_data = []
    total_pages = 0
    start_index = 0
    end_index = 0

    if show_all:
        # If showing all, set page to 1, one total page, and use all sorted data
        page = 1
        total_pages = 1 if total_items > 0 else 0
        start_index = 0
        end_index = total_items
        page_data = sorted_data
    elif total_items > 0:
        # Calculate total pages needed based on numeric page_size
        total_pages = math.ceil(total_items / page_size)
        # Adjust current page if it exceeds total pages (or is less than 1)
        page = max(1, min(page, total_pages))
        # Calculate start and end index for slicing
        start_index = (page - 1) * page_size
        end_index = start_index + page_size
        # Get the slice of data for the current page
        page_data = sorted_data[start_index:end_index]
    # else: variables remain 0 / empty list if total_items is 0

    # Create pagination info dictionary for the template
    pagination_info = {
        'page': page,
        'total_pages': total_pages,
        'total_items': total_items,
        'has_prev': page > 1 and not show_all,
        'prev_num': page - 1,
        'has_next': page < total_pages and not show_all,
        'next_num': page + 1,
        'start_item': min(start_index + 1, total_items) if total_items > 0 else 0,
        'end_item': min(end_index, total_items)
    }
    logging.debug(f"Pagination for '{comparison_type}': Page {page}/{total_pages}, Size='{page_size_str}', Items {pagination_info['start_item']}-{pagination_info['end_item']} of {total_items}")

    # --- Render Template ---
    return render_template(
        'results_viewer.html',
        title=comparison_type.replace(COMPARISON_SUFFIX, ''),
        page_data=page_data,
        pagination=pagination_info,
        filename=filename,
        available_sheets=available_sheets, # For nav links in base.html
        current_comparison_type=comparison_type,
        current_headers=current_headers,
        sort_by=sort_by,
        sort_order=sort_order,
        page_size_str=page_size_str,
        page_size_options=PAGE_SIZE_OPTIONS,
        comparison_suffix_for_template=COMPARISON_SUFFIX, # For nav links in base.html
        skill_expr_sheet_name=SKILL_EXPR_SHEET_NAME,
        error=error
    )


@ui_bp.route('/refresh')
def refresh_data():
    """
    Clears the cached Excel data and redirects to the upload page.
    """
    logger.info("Refresh request received. Clearing data cache.")
    current_app.config['EXCEL_DATA'] = {}
    current_app.config['EXCEL_FILENAME'] = None
    current_app.config['COMPARISON_SHEETS'] = []
    current_app.config['SHEET_HEADERS'] = {}
    current_app.config['MAX_DN_ID'] = 0
    current_app.config['MAX_AG_ID'] = 0
    session.pop('last_viewed_comparison', None) # Clear last viewed page from session
    flash("Data cache cleared. Please upload an Excel file.", "info")
    return redirect(url_for('ui.upload_config_page'))

