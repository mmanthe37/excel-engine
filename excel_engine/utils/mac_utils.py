"""
macOS-specific utilities — AppleScript execution, process management,
clipboard operations, Retina coordinate handling.
"""

from __future__ import annotations

import subprocess
import time
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class MacUtils:
    """macOS system utilities for the Excel Engine."""

    # ── AppleScript Execution ──

    @staticmethod
    def run_applescript(script: str, timeout: float = 30.0) -> str:
        """
        Execute an AppleScript string via osascript.
        Returns stdout on success, raises on failure.
        """
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=timeout,
            )
            if result.returncode != 0:
                stderr = result.stderr.strip()
                logger.error("AppleScript error: %s", stderr)
                raise RuntimeError(f"AppleScript failed: {stderr}")
            return result.stdout.strip()
        except subprocess.TimeoutExpired:
            raise TimeoutError(
                f"AppleScript timed out after {timeout}s: {script[:100]}..."
            )

    @staticmethod
    def run_applescript_file(script_path: Path, timeout: float = 30.0) -> str:
        """Execute an AppleScript file via osascript."""
        result = subprocess.run(
            ["osascript", str(script_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if result.returncode != 0:
            raise RuntimeError(f"AppleScript file failed: {result.stderr.strip()}")
        return result.stdout.strip()

    # ── Clipboard ──

    @staticmethod
    def clipboard_copy(text: str) -> None:
        """Copy text to macOS clipboard via pbcopy."""
        subprocess.run(
            ["pbcopy"],
            input=text.encode("utf-8"),
            check=True,
        )

    @staticmethod
    def clipboard_paste() -> str:
        """Read text from macOS clipboard via pbpaste."""
        result = subprocess.run(
            ["pbpaste"],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout

    # ── Process Management ──

    @staticmethod
    def is_excel_running() -> bool:
        """Check if Microsoft Excel is currently running."""
        result = subprocess.run(
            ["pgrep", "-x", "Microsoft Excel"],
            capture_output=True,
        )
        return result.returncode == 0

    @staticmethod
    def launch_excel() -> None:
        """Launch Microsoft Excel if not already running."""
        if not MacUtils.is_excel_running():
            subprocess.run(
                ["open", "-a", "Microsoft Excel"],
                check=True,
            )
            time.sleep(3)  # wait for Excel to initialize

    @staticmethod
    def open_workbook_in_excel(path: Path) -> None:
        """Open a specific workbook in Excel."""
        subprocess.run(
            ["open", "-a", "Microsoft Excel", str(path)],
            check=True,
        )
        time.sleep(2)

    @staticmethod
    def activate_excel() -> None:
        """Bring Microsoft Excel to the foreground."""
        MacUtils.run_applescript(
            'tell application "Microsoft Excel" to activate'
        )

    @staticmethod
    def quit_excel(saving: bool = True) -> None:
        """Quit Microsoft Excel."""
        save_str = "saving yes" if saving else "saving no"
        MacUtils.run_applescript(
            f'tell application "Microsoft Excel" to quit {save_str}'
        )

    # ── Retina Display ──

    @staticmethod
    def retina_to_logical(x: int, y: int) -> tuple[int, int]:
        """Convert Retina physical coordinates to logical (divide by 2)."""
        return x // 2, y // 2

    @staticmethod
    def logical_to_retina(x: int, y: int) -> tuple[int, int]:
        """Convert logical coordinates to Retina physical (multiply by 2)."""
        return x * 2, y * 2

    # ── Screen Info ──

    @staticmethod
    def get_screen_size() -> tuple[int, int]:
        """Get screen size via system_profiler (logical pixels)."""
        try:
            script = (
                'tell application "Finder" to get bounds of window of desktop'
            )
            result = MacUtils.run_applescript(script)
            parts = result.split(", ")
            if len(parts) == 4:
                return int(parts[2]), int(parts[3])
        except Exception:
            pass
        return 1920, 1080  # fallback

    # ── Wait Helpers ──

    @staticmethod
    def wait_for_excel_ready(timeout: float = 15.0) -> bool:
        """Wait until Excel is running and responsive."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                MacUtils.run_applescript(
                    'tell application "Microsoft Excel" to get name of workbook 1',
                    timeout=5.0,
                )
                return True
            except Exception:
                time.sleep(1)
        return False

    @staticmethod
    def wait_for_file(path: Path, timeout: float = 10.0) -> bool:
        """Wait for a file to exist on disk."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            if path.exists():
                return True
            time.sleep(0.5)
        return False
