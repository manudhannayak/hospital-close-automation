"""Month-end close reporting automation."""
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = ROOT / "config"
EXTRACTS = ROOT / "data" / "extracts"
LEGACY = ROOT / "data" / "legacy"
OUTPUT = ROOT / "output"
SITE = ROOT / "docs"


def load_settings() -> dict:
    with open(CONFIG_DIR / "settings.yaml") as fh:
        return yaml.safe_load(fh)
