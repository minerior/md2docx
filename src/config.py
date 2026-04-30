"""加载 gongwen.yaml，提供全局配置对象."""

from pathlib import Path
from typing import Any

import yaml


class GongwenConfig:
    """单例配置对象."""

    _instance = None

    def __new__(cls, path: str = "gongwen.yaml"):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, path: str = "gongwen.yaml"):
        if self._initialized:
            return
        self._path = Path(path)
        with self._path.open("r", encoding="utf-8") as f:
            self._data = yaml.safe_load(f)
        self._initialized = True

    def get(self, *keys: str, default: Any = None) -> Any:
        """链式取值，如 config.get("cover", "enabled")."""
        data = self._data
        for key in keys:
            if not isinstance(data, dict):
                return default
            data = data.get(key, default)
            if data is None:
                return default
        return data

    def get_cover_config(self) -> dict:
        return self.get("cover", default={})

    def get_type_mapping(self) -> dict[str, str]:
        return self.get("type_mapping", default={})
