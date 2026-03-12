"""
Configuration management for pdf2word.
Handles saving and loading the ConvertAPI key to a local config file.
"""

import json
import logging
import os

logger = logging.getLogger(__name__)

CONFIG_FILE = os.path.expanduser("~/.pdf2word_config.json")

def load_api_key() -> str | None:
    """Load the ConvertAPI key from the global configuration file."""
    if not os.path.exists(CONFIG_FILE):
        return None
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
            return config.get("convertapi_secret")
    except Exception as e:
        logger.warning("Failed to read config file %s: %s", CONFIG_FILE, e)
        return None

def save_api_key(api_key: str):
    """Save the ConvertAPI key to the global configuration file."""
    config = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception:
            pass # Ignore errors, we'll overwrite it
            
    config["convertapi_secret"] = api_key
    
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
        logger.info("API key successfully saved to %s", CONFIG_FILE)
    except Exception as e:
        logger.error("Failed to save API key to %s: %s", CONFIG_FILE, e)
        raise RuntimeError(f"Could not save configuration: {e}")

def remove_api_key():
    """Remove the ConvertAPI key from the global configuration file."""
    if not os.path.exists(CONFIG_FILE):
        return
        
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
            
        if "convertapi_secret" in config:
            del config["convertapi_secret"]
            
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
        logger.info("API key successfully removed from %s", CONFIG_FILE)
    except Exception as e:
        logger.error("Failed to remove API key from %s: %s", CONFIG_FILE, e)
        raise RuntimeError(f"Could not remove configuration: {e}")
