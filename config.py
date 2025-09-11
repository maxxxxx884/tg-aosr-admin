# config.py
import json
from pathlib import Path
from typing import Any, Dict

CONFIG_PATH = Path(__file__).with_name('config.json')

DEFAULT: Dict[str, Any] = {
    "tasks": [
        {
            "enabled": True,
            "module": "scripts.search_in_file",  # какой модуль запускать
            "function": "run",                   # какая функция в модуле
            "params": {                          # параметры для неё
                "file": "sample.txt",
                "patterns": ["ERROR", "CRITICAL"]
            }
        }
    ]
}

def _create_default_file() -> None:
    CONFIG_PATH.write_text(json.dumps(DEFAULT, ensure_ascii=False, indent=4),
                           encoding='utf-8')

if not CONFIG_PATH.exists():
    _create_default_file()

with CONFIG_PATH.open(encoding='utf-8') as f:
    data: Dict[str, Any] = json.load(f)

def save() -> None:
    CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=4),
                           encoding='utf-8')
