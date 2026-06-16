from __future__ import annotations

import configparser
import json
import os

import generator

APP_CONFIG_FILE = generator.resource_path("config.ini")
DEFAULT_EXPORT_PATH = os.path.expanduser("~")


def split_csv(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def load_app_config(path: str = APP_CONFIG_FILE) -> configparser.ConfigParser:
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
UIC_QTY_PRODUCTS = set(split_csv(APP_CONFIG["quantity_products"].get("uic", "")))

COUPLE = dict(APP_CONFIG.items("couples"))

APP_SUPPORT_DIR = os.path.expanduser(APP_CONFIG["app"]["support_dir"])
CONFIG_FILE = os.path.join(APP_SUPPORT_DIR, APP_CONFIG["app"]["export_config_file"])
PRODUCT_SPECS_FILE = generator.resource_path(APP_CONFIG["app"]["product_specs_file"])
COMPANIES = tuple(split_csv(APP_CONFIG["app"]["companies"]))

export_paths = {company: DEFAULT_EXPORT_PATH for company in COMPANIES}


def create_export_config_file():
    if DEBUG:
        print(f"Creating export config file at {CONFIG_FILE} with default paths: {export_paths}")
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(obj=export_paths, fp=f, ensure_ascii=False)


def load_last_path(company: str):
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            saved_path = data.get(f"{company}_export_path", DEFAULT_EXPORT_PATH)
            if saved_path and os.path.isdir(saved_path):
                return saved_path
    except Exception:
        return DEFAULT_EXPORT_PATH
    return DEFAULT_EXPORT_PATH


def save_all_paths(paths: dict):
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    data = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    for company, path in paths.items():
        if path and os.path.isdir(path):
            data[f"{company}_export_path"] = path
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        if DEBUG:
            print(f"Saved export config paths: {paths}")
    except Exception as e:
        if DEBUG:
            print(f"Failed to save export config paths: {str(e)}")


def load_product_specs(path: str = PRODUCT_SPECS_FILE) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_export_paths():
    for company in COMPANIES:
        export_paths[company] = load_last_path(company)
    return export_paths
