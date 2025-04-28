# -*- coding: utf-8 -*-
"""
Main Flask application file for the Excel Comparator UI.
Initializes the Flask app, loads configuration, registers blueprints,
and runs the development server.
"""

import os
import sys # Import sys for sys.exit
import logging
from flask import Flask, redirect, url_for, flash # Import flash for potential use

# Import configuration loading function
from config import load_config, save_config # Assuming these functions exist in config.py

# Import blueprints
from blueprints.ui_routes import ui_bp
from blueprints.template_routes import template_bp
from blueprints.processing_routes import processing_bp

# --- Constants ---
CONFIG_FILE = 'config.ini'
LOG_FILE = 'ui_viewer.log' # Central log file for the web app part
TEMPLATE_DIR = './config_templates/'
UPLOAD_FOLDER = './uploads' # Ensure upload folder is defined here too if needed by app setup

# --- Logging Setup ---
# Configure logging (similar to previous ui_viewer.py setup)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w'),
        logging.StreamHandler()
    ]
)
for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.StreamHandler):
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(levelname)s: [%(module)s] %(message)s')
        handler.setFormatter(formatter)

logger = logging.getLogger(__name__)

# --- Flask App Creation ---
def create_app():
    """Creates and configures the Flask application."""
    app = Flask(__name__, template_folder='templates', static_folder='static')

    # --- IMPORTANT: Set a Secret Key ---
    # This is required for session management (e.g., flash messages).
    # Use a strong, random secret key in production, possibly from environment variables.
    # For development, a simple key is okay.
    app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key-replace-in-prod')
    if app.secret_key == 'dev-secret-key-replace-in-prod':
        logger.warning("Using default development secret key. Set FLASK_SECRET_KEY environment variable for production.")
    # --- End Secret Key Setup ---


    # Load configuration
    # In a real app, use Flask's config handling more robustly
    # For now, load into a simple dict accessible via current_app later if needed
    # Or pass config explicitly to blueprints/functions
    try:
        app_config = load_config(CONFIG_FILE)
        # Store loaded config in Flask's config object for better access
        app.config['APP_SETTINGS'] = app_config
        # Initialize caches based on config or defaults
        app.config['EXCEL_DATA'] = {}
        app.config['EXCEL_FILENAME'] = None
        app.config['COMPARISON_SHEETS'] = []
        app.config['SHEET_HEADERS'] = {}
        app.config['MAX_DN_ID'] = 0
        app.config['MAX_AG_ID'] = 0
        # Store config file path for saving later
        app.config['CONFIG_FILE_PATH'] = CONFIG_FILE
        logger.info("Application configuration loaded.")
    except Exception as e:
        logger.error(f"FATAL: Could not load configuration from {CONFIG_FILE}. Error: {e}", exc_info=True)
        # Decide how to handle config load failure - exit or run with defaults?
        # For now, we might let it continue but log the error prominently.
        print(f"FATAL ERROR loading configuration: {e}. Check logs. Exiting.")
        sys.exit(1) # Exit if config is critical

    # Ensure template and upload directories exist
    for dir_path in [TEMPLATE_DIR, UPLOAD_FOLDER]:
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
                logger.info(f"Created missing directory: {dir_path}")
            except OSError as e:
                logger.error(f"Could not create directory {dir_path}: {e}")
                print(f"ERROR: Could not create directory {dir_path}. Please create it manually.")
                # Consider exiting if directories are critical
                # sys.exit(1)

    # Register blueprints
    if ui_bp: app.register_blueprint(ui_bp)
    if template_bp: app.register_blueprint(template_bp, url_prefix='/templates') # Add prefix for template routes
    if processing_bp: app.register_blueprint(processing_bp, url_prefix='/api') # Add prefix for processing API routes
    logger.info("Registered Flask Blueprints.")

    # Add a simple root route redirecting to the upload/config page
    @app.route('/')
    def index():
        # Redirect to the main UI page defined in ui_routes
        return redirect(url_for('ui.upload_config_page'))

    return app

# --- Main Execution ---
if __name__ == '__main__':
    logger.info("Starting Excel Comparator Flask Application...")
    app = create_app()
    # Run the Flask development server
    # Use host='0.0.0.0' to make accessible on network, default port 5000
    try:
        print(f"\nApplication running. Open your web browser and go to http://127.0.0.1:5001\nLog file: {LOG_FILE}")
        # Consider using Waitress or Gunicorn for production instead of app.run
        app.run(host='127.0.0.1', port=5001, debug=False, threaded=True)
    except OSError as e:
        if "address already in use" in str(e).lower():
             err_msg = "Port 5001 is already in use. Close other apps or change the port in app.py."
             logger.error(err_msg)
             print(f"ERROR: {err_msg}")
        else:
             logger.error(f"Failed to start Flask server: {e}", exc_info=True)
             print(f"ERROR: Failed to start web server. See {LOG_FILE} for details.")
    except Exception as e:
        logger.error(f"An unexpected error occurred on startup: {e}", exc_info=True)
        print(f"FATAL: An unexpected error occurred. See {LOG_FILE} for details.")

