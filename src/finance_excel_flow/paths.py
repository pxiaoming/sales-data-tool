"""应用目录与打包资源路径解析。"""

from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def get_app_root() -> Path:
    """exe 所在目录（打包）或项目根目录（开发）。"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def get_bundle_root() -> Path:
    """内置配置文件、模板等资源所在目录。"""
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS"))
    return get_app_root()


def resolve_bundle_path(relative: str | Path) -> Path:
    path = Path(relative)
    if path.is_absolute():
        return path
    return get_bundle_root() / path
