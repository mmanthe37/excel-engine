"""
Path handling utilities — colons, spaces, special characters, Desktop copy workaround.

Key rule: colons AND spaces in macOS paths break xlwings (Apple Events).
The workaround is to copy the file to ~/Desktop before opening with xlwings,
then copy back.
"""

from __future__ import annotations

import logging
import shutil
import re
import zipfile
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Minimum set of files required in a valid OOXML (.xlsx) archive
REQUIRED_OOXML_FILES = frozenset({
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
})


class PathHandler:
    """Safe path operations for Excel files on macOS."""

    # Colons break xlwings Apple Events; spaces cause OSERROR -50
    UNSAFE_CHARS_PATTERN = re.compile(r"[: ]")

    def __init__(self, desktop_path: Optional[Path] = None) -> None:
        self.desktop_path = desktop_path or Path.home() / "Desktop"

    def has_unsafe_chars(self, path: Path) -> bool:
        """Check if a path contains characters that break xlwings."""
        return bool(self.UNSAFE_CHARS_PATTERN.search(str(path)))

    def safe_copy_for_xlwings(self, src: Path) -> Path:
        """
        Copy a file to ~/Desktop if its path contains unsafe characters
        (colons or spaces). Returns the safe path (may be the original
        if no copy needed).
        """
        if not self.has_unsafe_chars(src):
            return src

        safe_name = self.sanitize_filename(src.name)
        dest = self.desktop_path / safe_name

        if dest.exists():
            dest = self._unique_path(dest)

        shutil.copy2(str(src), str(dest))
        logger.info("Copied to safe xlwings path: %s → %s", src.name, dest)
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
        return name.replace(":", "_").replace("/", "_").replace(" ", "_")

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

    # ── OOXML Validation ──

    @staticmethod
    def validate_ooxml(path: Path) -> tuple[bool, str]:
        """
        Validate that a file is a well-formed OOXML (.xlsx) archive.

        Returns (ok, message). If ok is False, the file should not be opened
        by openpyxl or used as a data source.
        """
        if not path.exists():
            return False, f"File does not exist: {path}"

        try:
            with zipfile.ZipFile(str(path)) as zf:
                names = set(zf.namelist())
                missing = REQUIRED_OOXML_FILES - names
                if missing:
                    return False, (
                        f"Invalid OOXML — missing required files: "
                        f"{', '.join(sorted(missing))}"
                    )
            return True, "OK"
        except zipfile.BadZipFile as e:
            return False, f"Not a valid ZIP archive: {e}"
        except Exception as e:
            return False, f"Cannot read file: {e}"

    @staticmethod
    def create_backup(path: Path) -> Path:
        """
        Create a uniquely-named backup of a workbook before modification.

        Returns the backup path. Never overwrites an existing backup.
        """
        backup = path.with_suffix(path.suffix + ".bak")
        if backup.exists():
            stem = backup.stem  # e.g., "file.xlsx"
            suffix = ".bak"
            parent = backup.parent
            counter = 1
            while backup.exists():
                backup = parent / f"{stem}.{counter}{suffix}"
                counter += 1
        shutil.copy2(str(path), str(backup))
        logger.info("Created backup: %s", backup.name)
        return backup
