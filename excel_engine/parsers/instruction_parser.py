"""
Instruction Parser — Parse .docx, .rtfd, .pdf, .txt instruction files.

Extracts raw text from instruction documents for task extraction.
"""

from __future__ import annotations

import logging
import re
import subprocess
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class InstructionParser:
    """Parse instruction files of various formats into plain text."""

    SUPPORTED_EXTENSIONS = {".docx", ".rtfd", ".pdf", ".txt", ".rtf"}

    def parse(self, path: Path) -> str:
        """
        Parse an instruction file and return its text content.
        Dispatches to the appropriate parser based on extension.
        """
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Instruction file not found: {path}")

        ext = path.suffix.lower()
        if path.is_dir() and str(path).endswith(".rtfd"):
            ext = ".rtfd"

        if ext not in self.SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported format '{ext}'. Supported: {self.SUPPORTED_EXTENSIONS}"
            )

        parser_map = {
            ".docx": self._parse_docx,
            ".rtfd": self._parse_rtfd,
            ".rtf": self._parse_rtf,
            ".pdf": self._parse_pdf,
            ".txt": self._parse_txt,
        }

        text = parser_map[ext](path)
        text = self._clean_text(text)
        logger.info("Parsed %s: %d characters", path.name, len(text))
        return text

    def parse_multiple(self, paths: list[Path]) -> str:
        """Parse multiple instruction files and concatenate their text."""
        texts = []
        for p in paths:
            try:
                texts.append(self.parse(p))
            except Exception as e:
                logger.warning("Failed to parse %s: %s", p, e)
        return "\n\n".join(texts)

    # ── Format-specific parsers ──

    def _parse_docx(self, path: Path) -> str:
        """Parse a .docx file using python-docx or textutil fallback."""
        try:
            import docx
            doc = docx.Document(str(path))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            # Also extract text from tables
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if cells:
                        paragraphs.append(" | ".join(cells))
            return "\n".join(paragraphs)
        except ImportError:
            logger.info("python-docx not available, falling back to textutil")
            return self._textutil_convert(path)

    def _parse_rtfd(self, path: Path) -> str:
        """
        Parse an .rtfd bundle (macOS rich text with attachments).
        Uses textutil to convert to plain text.
        """
        return self._textutil_convert(path)

    def _parse_rtf(self, path: Path) -> str:
        """Parse an .rtf file using textutil."""
        return self._textutil_convert(path)

    def _parse_pdf(self, path: Path) -> str:
        """Parse a PDF file using pdfplumber, PyPDF2, or macOS textutil."""
        # Try pdfplumber first (best extraction)
        try:
            import pdfplumber
            texts = []
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        texts.append(text)
            if texts:
                return "\n\n".join(texts)
        except ImportError:
            pass

        # Try PyPDF2
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(str(path))
            texts = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    texts.append(text)
            if texts:
                return "\n\n".join(texts)
        except ImportError:
            pass

        # Fallback: macOS mdimport / textutil
        try:
            result = subprocess.run(
                ["mdimport", "-d2", str(path)],
                capture_output=True, text=True, timeout=30,
            )
            if result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass

        raise RuntimeError(
            f"Cannot parse PDF '{path.name}': install pdfplumber or PyPDF2"
        )

    def _parse_txt(self, path: Path) -> str:
        """Parse a plain text file."""
        return path.read_text(encoding="utf-8", errors="replace")

    # ── Helpers ──

    @staticmethod
    def _textutil_convert(path: Path) -> str:
        """
        Convert a document to plain text using macOS textutil.
        Works for .docx, .rtf, .rtfd, .doc, .html, .webarchive.
        """
        result = subprocess.run(
            ["textutil", "-convert", "txt", "-stdout", str(path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if result.returncode != 0:
            raise RuntimeError(f"textutil failed: {result.stderr.strip()}")
        return result.stdout

    @staticmethod
    def _clean_text(text: str) -> str:
        """Clean up extracted text — normalize whitespace, remove artifacts."""
        # Collapse multiple blank lines
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Remove form feed and other control chars (keep \n, \t)
        text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
        return text.strip()
