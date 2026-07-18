from __future__ import annotations

import configparser
import json
import os

from . import generator

APP_CONFIG_FILE = generator.resource_path("config.ini")
DEFAULT_EXPORT_PATH = os.path.expanduser("~")


def split_csv(value: str) -> list[str]:
    """Split a comma-separated configuration value into non-empty items.

    Args:
        value: Raw comma-separated configuration text.

    Returns:
        Stripped, non-empty configuration items.
    """
    return [item.strip() for item in value.split(",") if item.strip()]


def load_app_config(path: str = APP_CONFIG_FILE) -> configparser.ConfigParser:
    """Load the application INI file while preserving option name casing.

    Args:
        path: Path to the INI configuration file.

    Returns:
        The populated configuration parser.

    Raises:
        FileNotFoundError: If the configuration file cannot be read.
    """
    parser = configparser.ConfigParser()
    parser.optionxform = str
    if not parser.read(path, encoding="utf-8"):
        raise FileNotFoundError(f"找不到設定檔：{path}")
    return parser


APP_CONFIG = load_app_config()
DEBUG = APP_CONFIG.getboolean("app", "debug", fallback=False)

ETACOM_TEMPLATE_FILE = generator.resource_path(APP_CONFIG["templates"]["etacom"])
BUSWAY_TEMPLATE_FILE = generator.resource_path(APP_CONFIG["templates"]["busway"])
YUASA_TEMPLATE_FILE = generator.resource_path(APP_CONFIG["templates"]["yuasa"])
UIC_TEMPLATE_FILE = generator.resource_path(APP_CONFIG["templates"]["uic"])

ETACOM_PRODUCT_NAME = split_csv(APP_CONFIG["products"]["etacom"])
BUSWAY_PRODUCT_NAME = split_csv(APP_CONFIG["products"]["busway"])
UIC_PRODUCT_NAME = split_csv(APP_CONFIG["products"]["uic"])

ETACOM_QTY_PRODUCTS = set(split_csv(APP_CONFIG["quantity_products"].get("etacom", "")))

COUPLE = dict(APP_CONFIG.items("couples"))

APP_SUPPORT_DIR = os.path.expanduser(APP_CONFIG["app"]["support_dir"])
CONFIG_FILE = os.path.join(APP_SUPPORT_DIR, APP_CONFIG["app"]["export_config_file"])
PRODUCT_SPECS_FILE = generator.resource_path(APP_CONFIG["app"]["product_specs_file"])
COMPANIES = tuple(split_csv(APP_CONFIG["app"]["companies"]))

export_paths = {company: DEFAULT_EXPORT_PATH for company in COMPANIES}


def create_export_config_file():
    """Create the per-user export-path configuration with default values.

    Returns:
        None. The function writes the initial JSON configuration to disk.
    """
    if DEBUG:
        print(
            f"Creating export config file at {CONFIG_FILE} "
            f"with default paths: {export_paths}"
        )
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as config_file:
        json.dump(obj=export_paths, fp=config_file, ensure_ascii=False)


def load_last_path(company: str):
    """Return the last valid export directory saved for a company.

    Args:
        company: Company identifier used in the saved configuration key.

    Returns:
        The saved directory, or the default export directory when unavailable.
    """
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as config_file:
            data = json.load(config_file)
            saved_path = data.get(f"{company}_export_path", DEFAULT_EXPORT_PATH)
            if saved_path and os.path.isdir(saved_path):
                return saved_path
    except Exception:
        return DEFAULT_EXPORT_PATH
    return DEFAULT_EXPORT_PATH


def save_all_paths(paths: dict):
    """Persist all valid company export directories to user configuration.

    Args:
        paths: Mapping of company identifiers to candidate export directories.

    Returns:
        None. Valid directories are merged into the JSON configuration.
    """
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    data = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as config_file:
                data = json.load(config_file)
        except Exception:
            pass
    for company, path in paths.items():
        if path and os.path.isdir(path):
            data[f"{company}_export_path"] = path
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as config_file:
            json.dump(data, config_file, ensure_ascii=False)
        if DEBUG:
            print(f"Saved export config paths: {paths}")
    except Exception as exc:
        if DEBUG:
            print(f"Failed to save export config paths: {exc}")


def load_product_specs(path: str = PRODUCT_SPECS_FILE) -> dict:
    """Load product specifications from a JSON file.

    Args:
        path: Path to the product specification JSON file.

    Returns:
        Nested product specification data keyed by company and product.
    """
    with open(path, "r", encoding="utf-8") as specs_file:
        return json.load(specs_file)


def load_export_paths():
    """Load and cache the last export directory for every company.

    Returns:
        The shared mapping of company identifiers to export directories.
    """
    for company in COMPANIES:
        export_paths[company] = load_last_path(company)
    return export_paths
