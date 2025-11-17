from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import json
from typing import Any


@dataclass(slots=True)
class InputSettings:
    bank_account: str = ""

    @classmethod
    def load(cls, path: Path | str = "src/input_settings.json") -> "InputSettings":
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
