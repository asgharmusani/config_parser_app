# -*- coding: utf-8 -*-
"""
Handles loading and saving application configuration from/to an INI file.
Focuses on API timeout, basic sheet layout hints (as fallbacks),
and a configurable logging level.
API URLs are now expected to be defined within Excel Rule Templates.
"""

import configparser
import logging # Import logging module to use its constants
import os
from typing import Dict, Any

# Import openpyxl utils for cell coordinate validation, if still needed for other parts
# from openpyxl.utils import cell as openpyxl_cell_utils # Not directly used in this version

logger = logging.getLogger(__name__) # Use module-specific logger

# Define the expected structure of the config.ini file for validation.
# 'API' section now only expects 'timeout'.
# 'SheetLayout' keys are kept as they might be used as fallbacks or by legacy parts.
EXPECTED_CONFIG = {
    # 'Files': [], # This section can be removed if no file-related global configs
    'API': ['timeout'], # Only timeout is expected here now
    'SheetLayout': ['ideal_agent_header_text', 'ideal_agent_fallback_cell', 'vag_extraction_sheet'],
    'Logging': ['level']
}

# Define default values for optional settings or if file is missing.
DEFAULT_CONFIG = {
    'API': {'timeout': '15'}, # Timeout as string initially, converted later
    'SheetLayout': { # Provide defaults for SheetLayout as well
        'ideal_agent_header_text': 'Ideal Agent',
        'ideal_agent_fallback_cell': 'C2',
        'vag_extraction_sheet': 'Default Targeting- Group'
    },
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
    Converts logging level string to logging constant and timeout to int.

    Args:
        config_path: Path to the config.ini file.

    Returns:
        A dictionary containing the configuration settings, using internal keys
        (e.g., 'api_timeout', 'log_level_value').

    Raises:
        FileNotFoundError: If the config file doesn't exist and cannot be created with defaults.
        ValueError: For missing required sections/options or type conversion errors.
    """
    logger.info(f"Attempting to load configuration from: {config_path}")
    config = configparser.ConfigParser(interpolation=None) # Disable % interpolation

    # Check if the configuration file exists
    if not os.path.exists(config_path):
        logger.warning(f"Configuration file '{config_path}' not found. Attempting to create with defaults.")
        # Attempt to create a default config file if it doesn't exist
        try:
            default_config_obj = configparser.ConfigParser(interpolation=None)
            # Populate with sections and keys from DEFAULT_CONFIG
            for section, section_keys_values in DEFAULT_CONFIG.items(): # Corrected variable name
                default_config_obj[section] = section_keys_values
            with open(config_path, 'w', encoding='utf-8') as default_configfile:
                default_config_obj.write(default_configfile)
            logger.info(f"Created default configuration file at '{config_path}'. Please review it.")
            # Now read the newly created default config
            config.read(config_path, encoding='utf-8')
        except Exception as e_create:
            logger.error(f"Could not create default configuration file at '{config_path}': {e_create}")
            raise FileNotFoundError(f"Configuration file '{config_path}' not found and could not be created.")


    # Read the configuration file (might be the newly created default one)
    try:
        # Ensure config object is fresh if it was just created by reading it again
        # This check might be redundant if creation logic above ensures config is populated
        if not config.sections():
            config.read(config_path, encoding='utf-8')
    except configparser.Error as e:
        logger.error(f"Error parsing configuration file '{config_path}': {e}")
        raise ValueError(f"Error parsing configuration file: {e}")

    # Dictionary to store the loaded settings
    settings = {}

    # Validate and extract settings based on the EXPECTED_CONFIG structure
    for section, keys in EXPECTED_CONFIG.items():
        # Check if the section exists in the file or defaults
        if not config.has_section(section):
            if section in DEFAULT_CONFIG:
                 logger.warning(f"Config section '[{section}]' not found, using defaults for this section.")
                 # Add default section to config object for consistent processing below
                 config[section] = DEFAULT_CONFIG[section]
            elif keys: # Only raise if keys were expected for this section
                 msg = f"Missing required section '[{section}]' in configuration file and no defaults provided."
                 logger.error(msg)
                 raise ValueError(msg)
            else: # Section has no required keys (like 'Files' now), it's okay if missing
                logger.debug(f"Section '[{section}]' not found, but no keys required. Skipping.")
                # settings[section] = {} # No need to create empty section in settings dict
                continue # Skip key processing for this section

        # Process each expected key within the section
        for key in keys:
            # internal_key_name = key # Default internal key name # Not needed with flat settings dict
            if config.has_option(section, key):
                value_str = config.get(section, key)
            elif key in DEFAULT_CONFIG.get(section, {}):
                value_str = DEFAULT_CONFIG[section][key]
                logger.debug(f"Setting '{key}' in '[{section}]' not found, using default: {value_str}")
            else:
                msg = f"Missing required option '{key}' in section '[{section}]' and no default provided."
                logger.error(msg)
                raise ValueError(msg)

            # Perform type conversion and specific key mapping
            if key == 'timeout' and section == 'API':
                try:
                    # Store timeout as integer under the key 'api_timeout'
                    settings['api_timeout'] = int(value_str)
                except ValueError:
                    msg = f"Invalid integer value for '{key}' in section '[{section}]': '{value_str}'. Using default 15."
                    logger.warning(msg)
                    settings['api_timeout'] = 15 # Fallback default for timeout
            elif key == 'level' and section == 'Logging':
                # Convert log level string to logging constant
                log_level_str_upper = value_str.upper()
                settings['log_level_value'] = LOG_LEVEL_MAP.get(log_level_str_upper, logging.INFO)
                # Store the string representation as well, for saving back and UI display
                settings['log_level_str'] = LOG_LEVEL_TO_STRING_MAP.get(settings['log_level_value'], 'INFO')
                if log_level_str_upper not in LOG_LEVEL_MAP:
                    logger.warning(f"Invalid logging level '{value_str}' in config. Defaulting to INFO.")
            else:
                # Store other keys as strings using the original key name
                settings[key] = value_str


    # --- Post-load validation for specific settings ---
    # Example: Validate fallback cell format (e.g., "C2")
    fallback_cell_key = 'ideal_agent_fallback_cell'
    if fallback_cell_key in settings: # Check if key exists in settings (it should if required or has default)
        try:
            # Import only when needed, to avoid circular dependencies if utils.py imports config.py
            from openpyxl.utils import cell as openpyxl_cell_utils_validator
            openpyxl_cell_utils_validator.coordinate_to_tuple(settings[fallback_cell_key])
        except openpyxl_cell_utils_validator.IllegalCharacterError:
            msg = f"Invalid format for '{fallback_cell_key}' in config: {settings[fallback_cell_key]}"
            logging.warning(f"{msg} This setting might not work as expected.") # Downgrade to warning
        except ImportError:
            logger.warning("openpyxl.utils.cell could not be imported. Skipping ideal_agent_fallback_cell validation.")
    # else: # If not in settings, it means it's not required by EXPECTED_CONFIG or used default

    logger.info("Configuration loaded successfully.")
    return settings


def save_config(config_path: str, settings: Dict[str, Any]):
    """
    Saves the provided settings dictionary to the INI configuration file.
    Organizes settings into sections based on their internal keys.
    Converts logging constant back to string for saving.
    API URLs (dn_url, agent_group_url) are no longer saved here.

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
    if 'api_timeout' in settings: # Use internal key 'api_timeout'
        config['API']['timeout'] = str(settings['api_timeout']) # Save as string
    # dn_url and agent_group_url are removed from this section

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
        config['Logging']['level'] = settings['log_level_str'].upper() # Ensure uppercase
    elif 'log_level_value' in settings: # Fallback if only value is present
        # Convert logging constant back to string
        config['Logging']['level'] = LOG_LEVEL_TO_STRING_MAP.get(settings['log_level_value'], 'INFO')
    else: # Default if not in settings
        config['Logging']['level'] = 'INFO'


    # Files Section (Currently empty, but create section header for consistency)
    # This section is no longer used for 'source_file' by the main app flow.
    if not config.has_section('Files'): # Create section if it doesn't exist
        config['Files'] = {}
    # If you had other file-related settings, they would go here:
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

