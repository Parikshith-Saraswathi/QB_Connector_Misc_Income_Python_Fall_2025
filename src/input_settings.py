from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import json
import sys
from typing import Any


@dataclass(slots=True)
class InputSettings:
    bank_account: str = ""

    @classmethod
    def load(cls, path: Path | str = "src/input_settings.json") -> "InputSettings":
        # Check if running as a PyInstaller bundle
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # If frozen, look inside the temp folder (_MEIPASS)
            p = Path(sys._MEIPASS) / path
        else:
            # If normal script, look in current working directory
            p = Path(path)

        if not p.exists():
            raise FileNotFoundError(f"Input settings file not found: {p}")
        
        data: dict[str, Any]
        with p.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        
        # Normalize keys
        bank_account = data.get("bank_account") or data.get("bankAccount") or ""
        return cls(bank_account=bank_account)


__all__ = ["InputSettings"]