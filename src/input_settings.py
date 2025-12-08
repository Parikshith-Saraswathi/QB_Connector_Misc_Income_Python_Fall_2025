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
    def load(cls, path: Path) -> "InputSettings":
        p = Path(path)
        # If running as a PyInstaller bundle and path is relative, prepend _MEIPASS
        if (
            getattr(sys, "frozen", False)
            and hasattr(sys, "_MEIPASS")
            and not p.is_absolute()
        ):
            p = Path(sys._MEIPASS) / p

        if not p.exists():
            raise FileNotFoundError(f"Input settings file not found: {p}")

        data: dict[str, Any]
        with p.open("r", encoding="utf-8") as fh:
            data = json.load(fh)

        # Normalize keys
        bank_account = data.get("bank_account") or data.get("bankAccount") or ""
        return cls(bank_account=bank_account)


__all__ = ["InputSettings"]
