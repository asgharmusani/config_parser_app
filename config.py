# -*- coding: utf-8 -*-
"""
Handles loading and saving application configuration from/to an INI file.
Includes handling for a configurable logging level.
"""

import configparser
import logging # Import logging module to use its constants
import os
from typing import Dict, Any

# Import openpyxl utils for cell coordinate validation, if still needed for other parts
# from openpyxl.utils import cell as openpyxl_cell_utils # Not directly used in this version

logger = logging.getLogger(__name__) # Use module-specific logger

# Define the expected structure of the config.ini file for validation.
EXPECTED_CONFIG = {
    'Files': [], # No longer expecting source_file here
    'API': ['dn_url', 'agent_group_url', 'timeout'],
    'SheetLayout': ['ideal_agent_header_text', 'ideal_agent_fallback_cell', 'vag_extraction_sheet'],
    'Logging': ['level'] # Added Logging section and level key
}

# Define default values for optional settings or if file is missing.
DEFAULT_CONFIG = {
    'API': {'timeout': '15'}, # Timeout as string initially, converted later
    'Logging': {'level': 'INFO'} # Default logging level
}

# Mapping from log level strings to logging module constants
LOG_LEVEL_MAP = {
    'DEBUG': logging.DEBUG,
    'INFO': logging.INFO,
    'WARNING': logging.WARNING,
    'ERROR': logging.ERROR,
    'CRITICAL': logging.CRITICAL
}
# Reverse mapping for saving
LOG_LEVEL_TO_STRING_MAP = {v: k for k, v in LOG_LEVEL_MAP.items()}


def load_config(config_path: str) -> Dict[str, Any]:
    """
    Loads configuration from the specified INI file.
    Uses defaults for missing optional values. Validates required sections/options.
    Converts logging level string to logging constant.

    Args:
        config_path: Path to the config.ini file.

    Returns:
        A dictionary containing the configuration settings, using internal keys
        (e.g., 'api_timeout', 'log_level_value').

    Raises:
        FileNotFoundError: If the config file doesn't exist.
        ValueError: For missing required sections/options or type conversion errors.
    """
    logger.info(f"Attempting to load configuration from: {config_path}")
    config = configparser.ConfigParser(interpolation=None) # Disable % interpolation

    # Check if the configuration file exists
    if not os.path.exists(config_path):
        logger.error(f"Configuration file '{config_path}' not found.")
        raise FileNotFoundError(f"Configuration file '{config_path}' not found.")

    # Read the configuration file
    try:
        config.read(config_path, encoding='utf-8')
    except configparser.Error as e:
        logger.error(f"Error parsing configuration file '{config_path}': {e}")
        raise ValueError(f"Error parsing configuration file: {e}")

    # Dictionary to store the loaded settings
    settings = {}

    # Validate and extract settings based on the EXPECTED_CONFIG structure
    for section, keys in EXPECTED_CONFIG.items():
        # Check if the section exists in the file
        if not config.has_section(section):
            # If section is optional (has defaults or no required keys), use defaults or skip
            if section in DEFAULT_CONFIG:
                 logger.warning(f"Config section '[{section}]' not found, using defaults for this section.")
                 # Add default section to config object for consistent processing below
                 config[section] = DEFAULT_CONFIG[section]
            # If section is required (has expected keys) but missing, raise error
            elif keys: # Only raise if keys were expected for this section
                 msg = f"Missing required section '[{section}]' in configuration file."
                 logger.error(msg)
                 raise ValueError(msg)
            # If section has no required keys (like 'Files' now), it's okay if missing
            else:
                logger.debug(f"Section '[{section}]' not found, but no keys required. Skipping.")
                settings[section] = {} # Ensure section exists in settings dict if it was optional and empty
                continue # Skip key processing for this section

        # Ensure section exists in settings dict even if read from file (or added from defaults)
        if section not in settings:
            settings[section] = {}

        # Process each expected key within the section
        for key in keys:
            if config.has_option(section, key):
                # Key exists in file, get its value
                value_str = config.get(section, key)
                # Perform type conversion and specific key mapping
                if key == 'timeout' and section == 'API':
                    try:
                        # Store timeout as integer under the key 'api_timeout'
                        settings['api_timeout'] = int(value_str)
                    except ValueError:
                        msg = f"Invalid integer value for '{key}' in section '[{section}]': '{value_str}'."
                        logger.error(msg)
                        raise ValueError(msg)
                elif key == 'level' and section == 'Logging':
                    # Convert log level string to logging constant
                    log_level_str = value_str.upper()
                    settings['log_level_value'] = LOG_LEVEL_MAP.get(log_level_str, logging.INFO)
                    if log_level_str not in LOG_LEVEL_MAP:
                        logger.warning(f"Invalid logging level '{value_str}' in config. Defaulting to INFO.")
                    settings['log_level_str'] = LOG_LEVEL_TO_STRING_MAP.get(settings['log_level_value'], 'INFO') # Store string too
                else:
                    # Store other keys as strings using the original key name
                    settings[key] = value_str
            elif key in DEFAULT_CONFIG.get(section, {}):
                # Key is missing in file, but has a default value defined
                default_value_str = DEFAULT_CONFIG[section][key]
                # Store default value using the correct internal key and type
                if key == 'timeout' and section == 'API':
                     try:
                         settings['api_timeout'] = int(default_value_str)
                     except ValueError:
                          msg = f"Invalid default integer value for '{key}': '{default_value_str}'."
                          logger.error(msg)
                          raise ValueError(msg)
                elif key == 'level' and section == 'Logging':
                    log_level_str = default_value_str.upper()
                    settings['log_level_value'] = LOG_LEVEL_MAP.get(log_level_str, logging.INFO)
                    settings['log_level_str'] = LOG_LEVEL_TO_STRING_MAP.get(settings['log_level_value'], 'INFO')
                else:
                    settings[key] = default_value_str
                logger.debug(f"Optional setting '{key}' not found in '[{section}]', using default: {default_value_str}")
            else:
                # Option is required (in EXPECTED_CONFIG) but missing in file and has no default
                msg = f"Missing required option '{key}' in section '[{section}]'."
                logger.error(msg)
                raise ValueError(msg)

    # --- Post-load validation for specific settings ---
    # Validate fallback cell format (e.g., "C2")
    fallback_cell_key = 'ideal_agent_fallback_cell'
    if fallback_cell_key in settings: # Check if key exists in settings (it should if required)
        try:
            # This import is only needed here if we are validating cell coordinates
            from openpyxl.utils import cell as openpyxl_cell_utils
            openpyxl_cell_utils.coordinate_to_tuple(settings[fallback_cell_key])
        except openpyxl_cell_utils.IllegalCharacterError:
            msg = f"Invalid format for '{fallback_cell_key}' in config: {settings[fallback_cell_key]}"
            logging.error(msg)
            raise ValueError(msg)
        except ImportError:
            logger.warning("openpyxl.utils.cell could not be imported. Skipping ideal_agent_fallback_cell validation.")

    else:
        # This case should be caught earlier by required check, but defensive check
        # if 'SheetLayout' is a required section and 'ideal_agent_fallback_cell' is a required key
        if 'SheetLayout' in EXPECTED_CONFIG and fallback_cell_key in EXPECTED_CONFIG['SheetLayout']:
            msg = f"Missing required option '{fallback_cell_key}' in section '[SheetLayout]'."
            logger.error(msg)
            raise ValueError(msg)


    logger.info("Configuration loaded successfully.")
    return settings


def save_config(config_path: str, settings: Dict[str, Any]):
    """
    Saves the provided settings dictionary to the INI configuration file.
    Organizes settings into sections based on their internal keys.
    Converts logging constant back to string for saving.

    Args:
        config_path: Path to the config.ini file.
        settings: Dictionary containing the configuration settings to save.
                  Keys should match the keys used internally (e.g., 'api_timeout', 'log_level_str').
    """
    logger.info(f"Attempting to save configuration to: {config_path}")
    config = configparser.ConfigParser(interpolation=None)

    # Reconstruct config structure from the flat 'settings' dict
    # Map internal keys back to their INI sections and keys

    # API Section
    config['API'] = {}
    if 'dn_url' in settings:
        config['API']['dn_url'] = settings['dn_url']
    if 'agent_group_url' in settings:
        config['API']['agent_group_url'] = settings['agent_group_url']
    if 'api_timeout' in settings:
        config['API']['timeout'] = str(settings['api_timeout']) # Save timeout as string

    # SheetLayout Section
    config['SheetLayout'] = {}
    if 'ideal_agent_header_text' in settings:
        config['SheetLayout']['ideal_agent_header_text'] = settings['ideal_agent_header_text']
    if 'ideal_agent_fallback_cell' in settings:
        config['SheetLayout']['ideal_agent_fallback_cell'] = settings['ideal_agent_fallback_cell']
    if 'vag_extraction_sheet' in settings:
        config['SheetLayout']['vag_extraction_sheet'] = settings['vag_extraction_sheet']

    # Logging Section
    config['Logging'] = {}
    if 'log_level_str' in settings: # Use the string representation for saving
        config['Logging']['level'] = settings['log_level_str']
    elif 'log_level_value' in settings: # Fallback if only value is present
        config['Logging']['level'] = LOG_LEVEL_TO_STRING_MAP.get(settings['log_level_value'], 'INFO')
    else: # Default if not in settings
        config['Logging']['level'] = 'INFO'


    # Files Section (Currently empty, but create section header for consistency)
    config['Files'] = {}
    # if 'some_other_file_setting' in settings:
    #     config['Files']['some_other_file_setting'] = settings['some_other_file_setting']


    # Ensure parent directory exists before writing
    config_dir = os.path.dirname(config_path)
    # Check if config_dir is not empty (i.e., not the current directory)
    if config_dir and not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
            logger.info(f"Created directory for config file: {config_dir}")
        except OSError as e:
            logger.error(f"Could not create directory for config file '{config_path}': {e}")
            raise IOError(f"Failed to create directory for config file: {e}")

    # Write the configuration to the file
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

