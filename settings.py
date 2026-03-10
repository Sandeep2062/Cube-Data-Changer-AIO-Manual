"""
Cross-platform settings store using a local JSON file.
Replaces the Windows-only registry approach.
"""

import json
import os

_SETTINGS_DIR = os.path.join(os.path.expanduser("~"), ".cube_data_aio")
_SETTINGS_FILE = os.path.join(_SETTINGS_DIR, "settings.json")


def _ensure_dir():
    os.makedirs(_SETTINGS_DIR, exist_ok=True)


def load():
    """Return the full settings dict (empty dict if not found)."""
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save(data: dict):
    """Persist the full settings dict."""
    _ensure_dir()
    with open(_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get(key, default=None):
    return load().get(key, default)


def put(key, value):
    d = load()
    d[key] = value
    save(d)
