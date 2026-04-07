"""
Path handling utilities — colons, special characters, Desktop copy workaround.

Key rule: colons in macOS paths break xlwings. The workaround is to copy
the file to ~/Desktop before opening with xlwings, then copy back.
"""

from __future__ import annotations

import shutil
import re
from pathlib import Path
from typing import Optional


class PathHandler:
    """Safe path operations for Excel files on macOS."""

    UNSAFE_CHARS_PATTERN = re.compile(r"[:]")

    def __init__(self, desktop_path: Optional[Path] = None) -> None:
        self.desktop_path = desktop_path or Path.home() / "Desktop"

    def has_unsafe_chars(self, path: Path) -> bool:
        """Check if a path contains characters that break xlwings (colons)."""
        return bool(self.UNSAFE_CHARS_PATTERN.search(str(path)))

    def safe_copy_for_xlwings(self, src: Path) -> Path:
        """
        Copy a file to ~/Desktop if its path contains colons.
        Returns the safe path (may be the original if no copy needed).
        """
        if not self.has_unsafe_chars(src):
            return src

        safe_name = self.sanitize_filename(src.name)
        dest = self.desktop_path / safe_name

        if dest.exists():
            dest = self._unique_path(dest)

        shutil.copy2(str(src), str(dest))
        return dest

    def copy_back_from_desktop(self, desktop_file: Path, original: Path) -> None:
        """Copy the modified file back from Desktop to the original location."""
        if desktop_file != original and desktop_file.exists():
            shutil.copy2(str(desktop_file), str(original))

    def cleanup_desktop_copy(self, desktop_file: Path, original: Path) -> None:
        """Remove the temporary Desktop copy if it differs from the original."""
        if desktop_file != original and desktop_file.exists():
            desktop_file.unlink()

    @staticmethod
    def sanitize_filename(name: str) -> str:
        """Remove or replace characters problematic for xlwings paths."""
        return name.replace(":", "_").replace("/", "_")

    @staticmethod
    def _unique_path(path: Path) -> Path:
        """Add a numeric suffix to avoid overwriting existing files."""
        stem = path.stem
        suffix = path.suffix
        parent = path.parent
        counter = 1
        while path.exists():
            path = parent / f"{stem}_{counter}{suffix}"
            counter += 1
        return path

    @staticmethod
    def to_posix(path: Path) -> str:
        """Convert a Path to a POSIX string suitable for AppleScript."""
        return str(path.resolve())

    @staticmethod
    def to_hfs(posix_path: str) -> str:
        """
        Convert a POSIX path to HFS (colon-separated) format for AppleScript.
        e.g., /Users/me/Desktop/file.xlsx → Macintosh HD:Users:me:Desktop:file.xlsx
        """
        parts = Path(posix_path).resolve().parts
        # parts[0] is '/', skip it; prepend volume name
        return "Macintosh HD:" + ":".join(parts[1:])

    @staticmethod
    def ensure_xlsx_extension(path: Path) -> Path:
        """Ensure the path has an .xlsx extension."""
        if path.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
            return path.with_suffix(".xlsx")
        return path
