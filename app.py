# -*- coding: utf-8 -*-
"""
Main Flask application file for the Excel Comparator UI.

This script initializes the Flask application, loads configuration settings
from config.ini (including a configurable logging level), registers the
necessary blueprints, ensures required directories exist, and runs the Flask
development server.
"""

import os
import sys # Import sys for sys.exit
import logging
import configparser # Import configparser for handling config load errors
from flask import Flask, redirect, url_for, flash # Import flash for potential use

# Import configuration loading function from config.py
try:
    from config import load_config, save_config
except ImportError as e:
    print(f"ERROR: Failed to import from config.py: {e}. Ensure config.py exists in the same directory.")
    sys.exit(1)

# Import blueprints from the blueprints package
try:
    from blueprints.ui_routes import ui_bp
    from blueprints.template_routes import template_bp
    from blueprints.processing_routes import processing_bp
    from blueprints.excel_rule_routes import excel_rule_bp
except ImportError as e:
    print(f"ERROR: Failed to import Blueprints. Ensure blueprint files exist in 'blueprints/' directory and __init__.py is present. Details: {e}")
    sys.exit(1)

# --- Constants ---
CONFIG_FILE = 'config.ini'
LOG_FILE = 'ui_viewer.log' # Central log file for the web app part
TEMPLATE_DIR = './config_templates/' # For DB update templates
EXCEL_RULE_TEMPLATE_DIR = './excel_rule_templates/' # For Excel processing rules
UPLOAD_FOLDER = './uploads/' # For uploaded and processed Excel files

# --- Logging Setup ---
# Configure logging to file and console
logging.basicConfig(
    level=logging.INFO, # Set root logger level (can be DEBUG for more detail)
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s:%(lineno)d] - %(message)s', # Added lineno for file
    datefmt='%Y-%m-%d %H:%M:%S', # Define date format for asctime
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w'), # Overwrite log file each run
        logging.StreamHandler(sys.stdout) # Log to console (stdout)
    ]
)
# Use a simpler format for console messages
console_formatter = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(module)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S' # Consistent date format
)
for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.StreamHandler):
        handler.setFormatter(console_formatter)

# --- MODIFICATION: Define module-level logger after basicConfig ---
logger = logging.getLogger(__name__)
# --- END MODIFICATION ---

# --- Flask App Creation ---
def create_app():
    """Creates and configures the Flask application instance."""
    # Get a logger specific to this function/context if needed, or use the module-level one
    # func_logger = logging.getLogger(f"{__name__}.create_app")

    app = Flask(__name__, template_folder='templates', static_folder='static')

    # --- IMPORTANT: Set a Secret Key ---
    app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key-replace-in-prod-very-secret')
    if app.secret_key == 'dev-secret-key-replace-in-prod-very-secret':
        logger.warning("Using default development secret key. SET FLASK_SECRET_KEY environment variable for production!")

    # --- Load Application Configuration ---
    try:
        app_config = load_config(CONFIG_FILE)
        app.config['APP_SETTINGS'] = app_config
        app.config['EXCEL_DATA'] = {}
        app.config['EXCEL_FILENAME'] = None
        app.config['COMPARISON_SHEETS'] = []
        app.config['SHEET_HEADERS'] = {}
        app.config['MAX_DN_ID'] = 0
        app.config['MAX_AG_ID'] = 0
        app.config['LAST_UPLOADED_ORIGINAL_FILE'] = None
        app.config['CONFIG_FILE_PATH'] = CONFIG_FILE
        app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
        logger.info("Application configuration file loaded.") # Use the module-level logger
    except (FileNotFoundError, ValueError, configparser.Error) as e:
        logger.error(f"FATAL: Could not load or validate configuration from {CONFIG_FILE}. Error: {e}", exc_info=True)
        print(f"FATAL ERROR loading configuration: {e}. Check logs ({LOG_FILE}). Exiting.")
        sys.exit(1)
    except Exception as e:
         logger.error(f"FATAL: Unexpected error during configuration loading: {e}", exc_info=True)
         print(f"FATAL ERROR during configuration loading. Check logs ({LOG_FILE}). Exiting.")
         sys.exit(1)

    # --- Reconfigure Logging Based on Loaded Config ---
    configured_log_level_str = app.config.get('APP_SETTINGS', {}).get('log_level_str', 'INFO').upper()
    log_level_map = {
        'DEBUG': logging.DEBUG, 'INFO': logging.INFO, 'WARNING': logging.WARNING,
        'ERROR': logging.ERROR, 'CRITICAL': logging.CRITICAL
    }
    configured_log_level = log_level_map.get(configured_log_level_str, logging.INFO)

    if configured_log_level_str not in log_level_map:
        logger.warning(f"Invalid log_level '{configured_log_level_str}' in config. Defaulting to INFO.")

    root_logger = logging.getLogger() # Get the root logger
    root_logger.setLevel(configured_log_level) # Set its level

    # Update levels of existing handlers
    for h in root_logger.handlers:
        h.setLevel(configured_log_level)

    logger.info(f"Logging reconfigured. Root logger level set to: {logging.getLevelName(root_logger.getEffectiveLevel())}")


    # --- Ensure Required Directories Exist ---
    for dir_path in [TEMPLATE_DIR, EXCEL_RULE_TEMPLATE_DIR, UPLOAD_FOLDER]:
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
                logger.info(f"Created missing directory: {dir_path}")
            except OSError as e:
                logger.error(f"Could not create directory {dir_path}: {e}")
                print(f"ERROR: Could not create directory {dir_path}. Please create it manually.")
                # sys.exit(1) # Consider exiting if directories are critical

    # --- Register Blueprints ---
    if ui_bp:
        app.register_blueprint(ui_bp)
        logger.info("Registered ui_bp Blueprint.")
    else:
        logger.error("ui_bp was not available for registration. UI routes will be unavailable.")

    if template_bp:
        app.register_blueprint(template_bp, url_prefix='/templates')
        logger.info("Registered template_bp Blueprint with prefix /templates.")
    else:
        logger.error("template_bp was not available for registration. DB Update Template routes will be unavailable.")

    if excel_rule_bp:
        app.register_blueprint(excel_rule_bp, url_prefix='/excel-rules')
        logger.info("Registered excel_rule_bp Blueprint with prefix /excel-rules.")
    else:
        logger.error("excel_rule_bp was not available for registration. Excel Rule Template routes will be unavailable.")

    if processing_bp:
        app.register_blueprint(processing_bp, url_prefix='/api')
        logger.info("Registered processing_bp Blueprint with prefix /api.")
    else:
        logger.error("processing_bp was not available for registration. Processing API routes will be unavailable.")

    # --- Define Root Route ---
    @app.route('/')
    def index():
        """Redirects the root URL ('/') to the main upload/config page."""
        return redirect(url_for('ui.upload_config_page'))

    return app

# --- Main Execution Guard ---
if __name__ == '__main__':
    # This logger is now defined and configured before use
    logger.info("Starting Excel Comparator Flask Application...")
    app_instance = create_app()
    try:
        print(f"\nApplication running. Open your web browser and go to http://127.0.0.1:5001\nLog file: {LOG_FILE}")
        app_instance.run(host='127.0.0.1', port=5001, debug=False, threaded=True)
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

