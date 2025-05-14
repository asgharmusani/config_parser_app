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

# Import configuration loading function from config.py
try:
    from config import load_config, save_config
except ImportError as e:
    print(f"ERROR: Failed to import from config.py: {e}. Ensure config.py exists.")
    sys.exit(1)

# Import blueprints from the blueprints package
try:
    from blueprints.ui_routes import ui_bp
    from blueprints.template_routes import template_bp
    from blueprints.processing_routes import processing_bp
    from blueprints.excel_rule_routes import excel_rule_bp # Added import for excel rules
except ImportError as e:
    print(f"ERROR: Failed to import Blueprints. Ensure blueprint files exist in 'blueprints/' directory. Details: {e}")
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
    level=logging.DEBUG, # Set root logger level (can be DEBUG for more detail)
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s:%(lineno)d] - %(message)s', # Added lineno for file
    datefmt='%Y-%m-%d %H:%M:%S', # Define date format for asctime
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w'), # Overwrite log file each run
        logging.StreamHandler(sys.stdout) # Log to console (stdout)
    ]
)
# Adjust console handler level and format
for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.StreamHandler):
        handler.setLevel(logging.DEBUG)
        # --- MODIFICATION START: Add asctime to console formatter ---
        console_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(module)s] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S' # Consistent date format
        )
        # --- MODIFICATION END ---
        handler.setFormatter(console_formatter)

logger = logging.getLogger(__name__)

# --- Flask App Creation ---
def create_app():
    """Creates and configures the Flask application instance."""
    app = Flask(__name__, template_folder='templates', static_folder='static')

    # --- IMPORTANT: Set a Secret Key ---
    # This is required for session management (e.g., flash messages).
    # Use a strong, random secret key in production, possibly from environment variables.
    app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key-replace-in-prod') # Replace default in production
    if app.secret_key == 'dev-secret-key-replace-in-prod':
        logger.warning("Using default development secret key. SET FLASK_SECRET_KEY environment variable for production.")
    # --- End Secret Key Setup ---


    # --- Load Application Configuration ---
    try:
        app_config = load_config(CONFIG_FILE)
        # Store loaded config in Flask's config object for global access via current_app
        app.config['APP_SETTINGS'] = app_config
        # Initialize caches and state variables in Flask's config
        app.config['EXCEL_DATA'] = {}
        app.config['EXCEL_FILENAME'] = None
        app.config['COMPARISON_SHEETS'] = []
        app.config['SHEET_HEADERS'] = {}
        app.config['MAX_DN_ID'] = 0
        app.config['MAX_AG_ID'] = 0
        app.config['LAST_UPLOADED_ORIGINAL_FILE'] = None # Track uploaded file path
        app.config['CONFIG_FILE_PATH'] = CONFIG_FILE # Store path for saving
        app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER # Make upload folder path available
        logger.info("Application configuration loaded and initial state set.")
    except (FileNotFoundError, ValueError, configparser.Error) as e:
        logger.error(f"FATAL: Could not load or validate configuration from {CONFIG_FILE}. Error: {e}", exc_info=True)
        print(f"FATAL ERROR loading configuration: {e}. Check logs ({LOG_FILE}). Exiting.")
        sys.exit(1) # Exit if config is critical and fails to load
    except Exception as e:
         logger.error(f"FATAL: Unexpected error during configuration loading: {e}", exc_info=True)
         print(f"FATAL ERROR during configuration loading. Check logs ({LOG_FILE}). Exiting.")
         sys.exit(1)


    # --- Ensure Required Directories Exist ---
    for dir_path in [TEMPLATE_DIR, EXCEL_RULE_TEMPLATE_DIR, UPLOAD_FOLDER]:
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
                logger.info(f"Created missing directory: {dir_path}")
            except OSError as e:
                logger.error(f"Could not create directory {dir_path}: {e}")
                print(f"ERROR: Could not create directory {dir_path}. Please create it manually.")
                # Consider exiting if directories are critical
                # sys.exit(1)

    # --- Register Blueprints ---
    # Registering blueprints organizes routes into modules
    if ui_bp:
        app.register_blueprint(ui_bp)
        logger.info("Registered ui_bp Blueprint.")
    else:
        logger.error("ui_bp not available for registration.")

    if template_bp:
        # Prefix template manager routes with /templates
        app.register_blueprint(template_bp, url_prefix='/templates')
        logger.info("Registered template_bp Blueprint with prefix /templates.")
    else:
        logger.error("template_bp not available for registration.")

    if excel_rule_bp:
        # Prefix excel rule manager routes with /excel-rules
        app.register_blueprint(excel_rule_bp, url_prefix='/excel-rules')
        logger.info("Registered excel_rule_bp Blueprint with prefix /excel-rules.")
    else:
        logger.error("excel_rule_bp not available for registration.")


    if processing_bp:
         # Prefix processing API routes with /api
        app.register_blueprint(processing_bp, url_prefix='/api')
        logger.info("Registered processing_bp Blueprint with prefix /api.")
    else:
        logger.error("processing_bp not available for registration.")

    # --- Define Root Route ---
    @app.route('/')
    def index():
        """Redirects the root URL ('/') to the main upload/config page."""
        # Use url_for to generate the URL for the target endpoint within the 'ui' blueprint
        return redirect(url_for('ui.upload_config_page'))

    return app

# --- Main Execution Guard ---
if __name__ == '__main__':
    # This block runs only when the script is executed directly (e.g., python app.py)
    logger.info("Starting Excel Comparator Flask Application...")
    # Create the Flask app instance
    app = create_app()
    # Run the Flask development server
    # Use host='0.0.0.0' to make accessible on network, default port 5001
    try:
        print(f"\nApplication running. Open your web browser and go to http://127.0.0.1:5001\nLog file: {LOG_FILE}")
        # Consider using Waitress or Gunicorn for production instead of app.run
        # debug=False is recommended for stability
        # threaded=True allows handling multiple requests concurrently
        app.run(host='127.0.0.1', port=5001, debug=False, threaded=True)
    except OSError as e:
        # Handle common error: port already in use
        if "address already in use" in str(e).lower():
             err_msg = "Port 5001 is already in use. Close other apps or change the port in app.py."
             logger.error(err_msg)
             print(f"ERROR: {err_msg}")
        else:
             # Handle other OS errors during server start
             logger.error(f"Failed to start Flask server: {e}", exc_info=True)
             print(f"ERROR: Failed to start web server. See {LOG_FILE} for details.")
    except Exception as e:
        # Catch any other unexpected errors during startup
        logger.error(f"An unexpected error occurred on startup: {e}", exc_info=True)
        print(f"FATAL: An unexpected error occurred. See {LOG_FILE} for details.")

