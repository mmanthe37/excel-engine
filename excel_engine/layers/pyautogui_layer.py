"""
Layer 6: PyAutoGUI — Last-resort desktop control.

Used when all other layers fail. Provides screen coordinate clicks,
keyboard shortcuts, and basic GUI interaction.

KEY RULES:
  - Retina display: divide physical coordinates by 2
  - Always add delays between actions
  - Use FAILSAFE (move mouse to corner to abort)
  - This is the last resort — prefer higher layers
"""

from __future__ import annotations

import logging
import time
from typing import Optional

logger = logging.getLogger(__name__)

try:
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.5
except ImportError:
    pyautogui = None  # type: ignore[assignment]
    logger.warning("pyautogui not installed — Layer 6 unavailable")


class PyAutoGUILayer:
    """Layer 6 — last-resort desktop control via PyAutoGUI."""

    def __init__(self, retina: bool = True, pause: float = 0.5) -> None:
        self.retina = retina
        self.pause = pause
        if pyautogui:
            pyautogui.PAUSE = pause

    @property
    def available(self) -> bool:
        return pyautogui is not None

    def _require(self) -> None:
        if not self.available:
            raise RuntimeError("pyautogui is not installed")

    def _adjust_coords(self, x: int, y: int) -> tuple[int, int]:
        """Adjust coordinates for Retina display (divide by 2)."""
        if self.retina:
            return x // 2, y // 2
        return x, y

    # ── Mouse Operations ──

    def click(self, x: int, y: int, clicks: int = 1) -> None:
        """Click at screen coordinates (physical pixels, auto-adjusted for Retina)."""
        self._require()
        ax, ay = self._adjust_coords(x, y)
        pyautogui.click(ax, ay, clicks=clicks)
        logger.debug("Click at (%d, %d) [adjusted: (%d, %d)]", x, y, ax, ay)

    def double_click(self, x: int, y: int) -> None:
        """Double-click at screen coordinates."""
        self.click(x, y, clicks=2)

    def right_click(self, x: int, y: int) -> None:
        """Right-click at screen coordinates."""
        self._require()
        ax, ay = self._adjust_coords(x, y)
        pyautogui.rightClick(ax, ay)

    def move_to(self, x: int, y: int, duration: float = 0.25) -> None:
        """Move mouse to coordinates."""
        self._require()
        ax, ay = self._adjust_coords(x, y)
        pyautogui.moveTo(ax, ay, duration=duration)

    def drag_to(
        self, x: int, y: int, duration: float = 0.5, button: str = "left"
    ) -> None:
        """Drag from current position to coordinates."""
        self._require()
        ax, ay = self._adjust_coords(x, y)
        pyautogui.dragTo(ax, ay, duration=duration, button=button)

    # ── Keyboard Operations ──

    def hotkey(self, *keys: str) -> None:
        """Press a hotkey combination (e.g., 'command', 'c')."""
        self._require()
        pyautogui.hotkey(*keys)
        logger.debug("Hotkey: %s", "+".join(keys))

    def type_text(self, text: str, interval: float = 0.02) -> None:
        """Type text character by character."""
        self._require()
        pyautogui.typewrite(text, interval=interval)

    def press(self, key: str) -> None:
        """Press a single key."""
        self._require()
        pyautogui.press(key)

    def key_down(self, key: str) -> None:
        """Hold a key down."""
        self._require()
        pyautogui.keyDown(key)

    def key_up(self, key: str) -> None:
        """Release a held key."""
        self._require()
        pyautogui.keyUp(key)

    # ── Screen ──

    def screenshot(self, region: Optional[tuple[int, int, int, int]] = None):
        """Take a screenshot. Returns a PIL Image."""
        self._require()
        return pyautogui.screenshot(region=region)

    def locate_on_screen(self, image_path: str, confidence: float = 0.8):
        """Find an image on screen. Returns (left, top, width, height) or None."""
        self._require()
        try:
            return pyautogui.locateOnScreen(image_path, confidence=confidence)
        except Exception:
            return None

    def click_image(self, image_path: str, confidence: float = 0.8) -> bool:
        """Find and click an image on screen. Returns True if found."""
        location = self.locate_on_screen(image_path, confidence)
        if location:
            center = pyautogui.center(location)
            pyautogui.click(center)
            return True
        return False

    # ── Compound Actions ──

    def cmd_c(self) -> None:
        """Copy (Cmd+C)."""
        self.hotkey("command", "c")

    def cmd_v(self) -> None:
        """Paste (Cmd+V)."""
        self.hotkey("command", "v")

    def cmd_z(self) -> None:
        """Undo (Cmd+Z)."""
        self.hotkey("command", "z")

    def cmd_s(self) -> None:
        """Save (Cmd+S)."""
        self.hotkey("command", "s")

    def tab(self, count: int = 1) -> None:
        """Press Tab n times."""
        for _ in range(count):
            self.press("tab")
            time.sleep(0.1)

    def enter(self) -> None:
        """Press Enter."""
        self.press("enter")

    def escape(self) -> None:
        """Press Escape."""
        self.press("escape")
