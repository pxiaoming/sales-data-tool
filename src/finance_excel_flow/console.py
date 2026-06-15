"""Windows 控制台 UTF-8 编码配置。"""

from __future__ import annotations

import sys


def configure_console_encoding() -> None:
    if sys.platform != "win32":
        return

    import os

    os.environ.setdefault("PYTHONUTF8", "1")

    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        kernel32.SetConsoleOutputCP(65001)
        kernel32.SetConsoleCP(65001)
    except (AttributeError, OSError):
        pass

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
