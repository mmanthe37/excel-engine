"""
Tests for the CLI — uses subprocess to invoke excel_engine.cli.main().
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_engine.cli import main, build_parser


# ── Helpers ──

def _run_cli(*args: str, check: bool = False) -> subprocess.CompletedProcess:
    """Run the CLI as a subprocess."""
    return subprocess.run(
        [sys.executable, "-m", "excel_engine.cli", *args],
        capture_output=True, text=True, timeout=30,
    )


@pytest.fixture
def instructions_file(tmp_path):
    """Create a simple instruction file."""
    p = tmp_path / "instructions.txt"
    p.write_text("Enter 100 in cell A1\nEnter 200 in cell B1\nFormat B2 as bold\n")
    return p


@pytest.fixture
def cli_workbook(tmp_path):
    """Create a minimal workbook for CLI tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Test"
    path = tmp_path / "cli_test.xlsx"
    wb.save(path)
    wb.close()
    return path


# ══════════════════════════════════════════════════════════════════
# Subcommand tests
# ══════════════════════════════════════════════════════════════════

class TestCLIInfo:
    def test_info_command(self):
        result = _run_cli("info")
        assert result.returncode == 0
        assert "Excel Engine" in result.stdout
        assert "Automation Layers" in result.stdout

    def test_info_shows_task_types(self):
        result = _run_cli("info")
        assert "formula" in result.stdout
        assert "Task Types" in result.stdout


class TestCLICheckEnv:
    def test_check_env(self):
        result = _run_cli("check-env")
        # Should always complete (may return 1 if deps missing, but shouldn't crash)
        assert result.returncode in (0, 1)
        assert "Environment Check" in result.stdout
        assert "Python" in result.stdout
        assert "openpyxl" in result.stdout


class TestCLIParse:
    def test_parse_basic(self, instructions_file):
        result = _run_cli("parse", str(instructions_file))
        assert result.returncode == 0
        assert "Parsed" in result.stdout

    def test_parse_json_output(self, instructions_file, tmp_path):
        out = tmp_path / "parsed.json"
        result = _run_cli("parse", str(instructions_file), "--output", str(out))
        assert result.returncode == 0
        assert out.exists()
        data = json.loads(out.read_text())
        assert "task_count" in data

    def test_parse_missing_file(self):
        result = _run_cli("parse", "/nonexistent/instructions.txt")
        assert result.returncode != 0


class TestCLIRun:
    def test_run_dry_run(self, cli_workbook, instructions_file):
        result = _run_cli("run", str(cli_workbook), str(instructions_file), "--dry-run")
        assert result.returncode == 0
        assert "DRY RUN" in result.stdout

    def test_run_missing_workbook(self, instructions_file):
        result = _run_cli("run", "/nonexistent/workbook.xlsx", str(instructions_file))
        assert result.returncode != 0

    def test_run_missing_instructions(self, cli_workbook):
        result = _run_cli("run", str(cli_workbook), "/nonexistent/instructions.txt")
        assert result.returncode != 0

    def test_run_phase1(self, cli_workbook, instructions_file):
        result = _run_cli("run", str(cli_workbook), str(instructions_file),
                          "--dry-run", "--phase", "1")
        assert result.returncode == 0

    def test_run_phase2(self, cli_workbook, instructions_file):
        result = _run_cli("run", str(cli_workbook), str(instructions_file),
                          "--dry-run", "--phase", "2")
        assert result.returncode == 0

    def test_run_dry_run_output(self, cli_workbook, instructions_file, tmp_path):
        out = tmp_path / "dry.json"
        result = _run_cli("run", str(cli_workbook), str(instructions_file),
                          "--dry-run", "--output", str(out))
        assert result.returncode == 0
        assert out.exists()
        data = json.loads(out.read_text())
        assert data["mode"] == "dry_run"

    def test_run_watch_flag(self, cli_workbook, instructions_file):
        result = _run_cli(
            "run", str(cli_workbook), str(instructions_file),
            "--dry-run", "--watch",
        )
        assert result.returncode == 0


class TestCLIVerify:
    def test_verify_basic(self, cli_workbook):
        result = _run_cli("verify", str(cli_workbook))
        assert result.returncode == 0
        assert "Sheet1" in result.stdout

    def test_verify_missing_file(self):
        result = _run_cli("verify", "/nonexistent/workbook.xlsx")
        assert result.returncode != 0


class TestCLIVersion:
    def test_version(self):
        result = _run_cli("--version")
        assert result.returncode == 0
        assert "excel-engine" in result.stdout or "1." in result.stdout


class TestCLINoCommand:
    def test_no_command_shows_help(self):
        result = _run_cli()
        assert result.returncode == 0
        # Should print help text
        assert "usage" in result.stdout.lower() or "Excel Engine" in result.stdout


# ══════════════════════════════════════════════════════════════════
# Unit tests for build_parser and main()
# ══════════════════════════════════════════════════════════════════

class TestCLIParser:
    def test_build_parser(self):
        parser = build_parser()
        assert parser is not None
        args = parser.parse_args(["info"])
        assert args.command == "info"

    def test_parser_run_watch_flag(self):
        parser = build_parser()
        args = parser.parse_args(["run", "a.xlsx", "b.txt", "--watch", "--dry-run"])
        assert args.command == "run"
        assert args.watch is True

    def test_main_info(self):
        ret = main(["info"])
        assert ret == 0

    def test_main_no_command(self):
        ret = main([])
        assert ret == 0

    def test_main_check_env(self):
        ret = main(["check-env"])
        assert ret in (0, 1)

    def test_main_parse(self, instructions_file):
        ret = main(["parse", str(instructions_file)])
        assert ret == 0

    def test_main_verify(self, cli_workbook):
        ret = main(["verify", str(cli_workbook)])
        assert ret == 0

    def test_main_run_dry_run(self, cli_workbook, instructions_file):
        ret = main(["run", str(cli_workbook), str(instructions_file), "--dry-run"])
        assert ret == 0

    def test_main_run_phase1_dry(self, cli_workbook, instructions_file):
        ret = main(["run", str(cli_workbook), str(instructions_file),
                     "--dry-run", "--phase", "1"])
        assert ret == 0

    def test_main_run_phase2_dry(self, cli_workbook, instructions_file):
        ret = main(["run", str(cli_workbook), str(instructions_file),
                     "--dry-run", "--phase", "2"])
        assert ret == 0

    def test_main_run_actual(self, cli_workbook, instructions_file):
        ret = main(["run", str(cli_workbook), str(instructions_file)])
        # May return 0 or 1 depending on what tasks succeed
        assert ret in (0, 1)

    def test_main_run_with_output(self, cli_workbook, instructions_file, tmp_path):
        out = tmp_path / "result.json"
        ret = main(["run", str(cli_workbook), str(instructions_file),
                     "--dry-run", "--output", str(out)])
        assert ret == 0
        assert out.exists()

    def test_main_parse_with_output(self, instructions_file, tmp_path):
        out = tmp_path / "tasks.json"
        ret = main(["parse", str(instructions_file), "--output", str(out)])
        assert ret == 0
        assert out.exists()

    def test_main_run_missing_workbook(self, instructions_file):
        ret = main(["run", "/nonexistent.xlsx", str(instructions_file)])
        assert ret == 1

    def test_main_run_missing_instructions(self, cli_workbook):
        ret = main(["run", str(cli_workbook), "/nonexistent.txt"])
        assert ret == 1

    def test_main_verify_with_instructions(self, cli_workbook, instructions_file):
        ret = main(["verify", str(cli_workbook), "-i", str(instructions_file)])
        # Verification may pass or fail
        assert ret in (0, 1)

    def test_main_verbose(self, instructions_file):
        ret = main(["-v", "parse", str(instructions_file)])
        assert ret == 0

    def test_main_verify_output(self, cli_workbook, instructions_file, tmp_path):
        out = tmp_path / "verify.json"
        ret = main(["verify", str(cli_workbook), "-i", str(instructions_file),
                     "--output", str(out)])
        assert ret in (0, 1)
        if out.exists():
            data = json.loads(out.read_text())
            assert "overall_passed" in data
