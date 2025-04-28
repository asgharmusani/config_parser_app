# -*- coding: utf-8 -*-
"""
Handles loading and saving application configuration from/to an INI file.
"""

import configparser
import logging
import os
from typing import Dict, Any
# --- MODIFICATION: Import openpyxl utils if needed for validation ---
from openpyxl.utils import cell as openpyxl_cell_utils

logger = logging.getLogger(__name__) # Use module-specific logger

# --- MODIFICATION: Removed 'source_file' from Files section ---
EXPECTED_CONFIG = {
    'Files': [], # No longer expecting source_file here
    'API': ['dn_url', 'agent_group_url', 'timeout'],
    'SheetLayout': ['ideal_agent_header_text', 'ideal_agent_fallback_cell', 'vag_extraction_sheet']
}
# --- END MODIFICATION ---

# Define default values for optional settings or if file is missing
DEFAULT_CONFIG = {
    'API': {'timeout': '15'}, # Timeout as string initially
    # Add other defaults if needed
}

def load_config(config_path: str) -> Dict[str, Any]:
    """
    Loads configuration from the specified INI file.
    Uses defaults for missing optional values.

    Args:
        config_path: Path to the config.ini file.

    Returns:
        A dictionary containing the configuration settings.

    Raises:
        FileNotFoundError: If the config file doesn't exist.
        ValueError: For missing required sections/options or type conversion errors.
    """
    logger.info(f"Attempting to load configuration from: {config_path}")
    config = configparser.ConfigParser(interpolation=None)

    if not os.path.exists(config_path):
        logger.error(f"Configuration file '{config_path}' not found.")
        raise FileNotFoundError(f"Configuration file '{config_path}' not found.")

    try:
        config.read(config_path, encoding='utf-8')
    except configparser.Error as e:
        logger.error(f"Error parsing configuration file '{config_path}': {e}")
        raise ValueError(f"Error parsing configuration file: {e}")

    settings = {}

    # Validate and extract settings based on EXPECTED_CONFIG
    for section, keys in EXPECTED_CONFIG.items():
        # --- MODIFICATION: Handle potentially empty Files section ---
        if not config.has_section(section) and section != 'Files': # Allow Files section to be missing now
        # --- END MODIFICATION ---
            # Use default section if available, otherwise raise error for required sections
            if section in DEFAULT_CONFIG:
                 logger.warning(f"Config section '[{section}]' not found, using defaults.")
                 config[section] = DEFAULT_CONFIG[section] # Add default section
            else:
                 # Raise error only if section is truly required and missing
                 if keys: # Only raise if keys were expected
                    raise ValueError(f"Missing required section '[{section}]' in configuration file.")
                 else: # If section has no required keys (like Files now), it's okay if missing
                    logger.debug(f"Section '[{section}]' not found, but no keys required.")
                    settings[section] = {} # Ensure section exists in settings dict
                    continue # Skip key processing for this section

        # Ensure section exists in settings dict even if read from file
        if section not in settings:
            settings[section] = {}

        for key in keys:
            if config.has_option(section, key):
                # Perform type conversion as needed
                if key == 'timeout':
                    try:
                        settings['api_timeout'] = config.getint(section, key) # Store as int with specific key
                    except ValueError:
                        raise ValueError(f"Invalid integer value for '{key}' in section '[{section}]'.")
                else:
                    settings[key] = config.get(section, key) # Store as string
            elif key in DEFAULT_CONFIG.get(section, {}):
                # Use default value if option is missing but has a default
                default_value = DEFAULT_CONFIG[section][key]
                # Store with the correct internal key ('api_timeout')
                if key == 'timeout':
                     try:
                         settings['api_timeout'] = int(default_value)
                     except ValueError:
                          raise ValueError(f"Invalid default integer value for '{key}'.")
                else:
                    settings[key] = default_value

                logger.debug(f"Optional setting '{key}' not found in '[{section}]', using default: {default_value}")
            else:
                # Option is required but missing
                raise ValueError(f"Missing required option '{key}' in section '[{section}]'.")

    # --- Post-load validation (moved from load_configuration caller) ---
    # Validate fallback cell format (e.g., "C2")
    fallback_cell_key = 'ideal_agent_fallback_cell'
    if fallback_cell_key in settings:
        try:
            openpyxl_cell_utils.coordinate_to_tuple(settings[fallback_cell_key])
        except openpyxl_cell_utils.IllegalCharacterError:
            msg = f"Invalid format for '{fallback_cell_key}' in config: {settings[fallback_cell_key]}"
            logging.error(msg)
            raise ValueError(msg)
    else:
        # This case should be caught earlier by required check, but defensive check
        raise ValueError(f"Missing required option '{fallback_cell_key}' in section '[SheetLayout]'.")


    logger.info("Configuration loaded successfully.")
    return settings


def save_config(config_path: str, settings: Dict[str, Any]):
    """
    Saves the provided settings dictionary to the INI configuration file.
    Organizes settings into sections based on EXPECTED_CONFIG.

    Args:
        config_path: Path to the config.ini file.
        settings: Dictionary containing the configuration settings to save.
                  Keys should match the keys used internally (e.g., 'api_timeout').
    """
    logger.info(f"Attempting to save configuration to: {config_path}")
    config = configparser.ConfigParser(interpolation=None)

    # Reconstruct config structure from settings dict
    # --- MODIFICATION: Removed source_file ---
    # config['Files'] = {}
    # if 'source_file' in settings: # This key should no longer be in settings
    #     config['Files']['source_file'] = settings['source_file']
    # --- END MODIFICATION ---

    config['API'] = {}
    if 'dn_url' in settings:
        config['API']['dn_url'] = settings['dn_url']
    if 'agent_group_url' in settings:
        config['API']['agent_group_url'] = settings['agent_group_url']
    if 'api_timeout' in settings:
        config['API']['timeout'] = str(settings['api_timeout']) # Save as string

    config['SheetLayout'] = {}
    if 'ideal_agent_header_text' in settings:
        config['SheetLayout']['ideal_agent_header_text'] = settings['ideal_agent_header_text']
    if 'ideal_agent_fallback_cell' in settings:
        config['SheetLayout']['ideal_agent_fallback_cell'] = settings['ideal_agent_fallback_cell']
    if 'vag_extraction_sheet' in settings:
        config['SheetLayout']['vag_extraction_sheet'] = settings['vag_extraction_sheet']

    # Ensure parent directory exists before writing
    config_dir = os.path.dirname(config_path)
    if config_dir and not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
            logger.info(f"Created directory for config file: {config_dir}")
        except OSError as e:
            logger.error(f"Could not create directory for config file '{config_path}': {e}")
            raise IOError(f"Failed to create directory for config file: {e}")

    try:
        with open(config_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        logger.info("Configuration saved successfully.")
    except IOError as e:
        logger.error(f"Error writing configuration file '{config_path}': {e}", exc_info=True)
        # Re-raise or handle as appropriate for the application
        raise IOError(f"Failed to save configuration: {e}")
    except Exception as e:
        logger.error(f"Unexpected error saving configuration: {e}", exc_info=True)
        raise

