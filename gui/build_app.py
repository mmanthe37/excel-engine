"""
Excel Engine GUI — Build standalone executables.

Usage:
    python gui/build_app.py                # Build for current platform
    python gui/build_app.py --clean        # Clean build artifacts first
    python gui/build_app.py --sign         # Build + code sign (macOS only)
    python gui/build_app.py --sign --notarize  # Build + sign + notarize

Requirements:
    pip install pyinstaller streamlit openpyxl

macOS signing requires:
    - Developer ID Application certificate in Keychain
    - For --notarize: APPLE_ID, APPLE_TEAM_ID, and either
      NOTARY_PASSWORD (app-specific password) or App Store Connect API key

Environment variables for notarization:
    APPLE_ID              Apple ID email
    APPLE_TEAM_ID         Team ID (e.g. 25QSUNYFC9)
    NOTARY_PASSWORD       App-specific password for notarytool
"""
import glob as glob_mod
import os
import subprocess
import sys
import platform
from pathlib import Path

SIGN_IDENTITY = "Developer ID Application: Michael Manthe (25QSUNYFC9)"
TEAM_ID = "25QSUNYFC9"


def _run(cmd, **kwargs):
    """Run a command, printing it first."""
    print(f"  >> {' '.join(str(c) for c in cmd)}")
    return subprocess.run(cmd, **kwargs)


def _codesign_app(app_path: Path, entitlements: Path):
    """Sign a macOS .app bundle inside-out (nested code first, then bundle)."""
    print(f"\nSigning {app_path.name}...")

    # Collect all signable objects inside the bundle
    sign_targets = []
    for pattern in ["**/*.so", "**/*.dylib", "**/*.framework", "**/*.bundle",
                    "**/*.xpc", "**/*.app", "**/Python", "**/python*"]:
        for match in app_path.glob(pattern):
            if match != app_path:
                sign_targets.append(match)
    # Also sign executables in MacOS/
    macos_dir = app_path / "Contents" / "MacOS"
    if macos_dir.exists():
        sign_targets.extend(f for f in macos_dir.iterdir() if f.is_file())
    # Deduplicate and sort deepest-first (inside-out signing)
    sign_targets = sorted(set(sign_targets), key=lambda p: -len(p.parts))

    # Sign nested code WITHOUT entitlements (entitlements go on top-level only)
    for target in sign_targets:
        if not target.exists():
            continue
        result = _run([
            "codesign", "--force", "--sign", SIGN_IDENTITY,
            "--options", "runtime",
            "--timestamp",
            str(target),
        ], capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  Warning: failed to sign {target.name}: {result.stderr.strip()}")

    # Sign the top-level .app bundle WITH entitlements
    result = _run([
        "codesign", "--force", "--sign", SIGN_IDENTITY,
        "--options", "runtime",
        "--timestamp",
        "--entitlements", str(entitlements),
        str(app_path),
    ], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"ERROR signing app bundle: {result.stderr}")
        sys.exit(1)

    # Verify with --deep --strict
    verify = _run(
        ["codesign", "--verify", "--deep", "--strict", str(app_path)],
        capture_output=True, text=True,
    )
    if verify.returncode == 0:
        print(f"Signed and verified: {app_path.name}")
    else:
        print(f"WARNING: Verification issues: {verify.stderr.strip()}")


def _notarize_app(app_path: Path):
    """Submit app for Apple notarization and staple the ticket."""
    apple_id = os.environ.get("APPLE_ID")
    team_id = os.environ.get("APPLE_TEAM_ID", TEAM_ID)
    password = os.environ.get("NOTARY_PASSWORD")

    if not apple_id or not password:
        print("\nSkipping notarization: set APPLE_ID and NOTARY_PASSWORD env vars")
        return False

    # Create a zip for notarization submission
    zip_path = app_path.parent / f"{app_path.stem}-notarize.zip"
    print(f"\nCreating zip for notarization: {zip_path.name}")
    _run([
        "ditto", "-c", "-k", "--keepParent", str(app_path), str(zip_path),
    ], check=True)

    # Submit to Apple
    print("Submitting to Apple notary service (this may take a few minutes)...")
    result = _run([
        "xcrun", "notarytool", "submit", str(zip_path),
        "--apple-id", apple_id,
        "--team-id", team_id,
        "--password", password,
        "--wait",
    ], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(f"Notarization failed: {result.stderr}")
        zip_path.unlink(missing_ok=True)
        return False

    # Staple the notarization ticket to the app
    print("Stapling notarization ticket...")
    staple = _run(
        ["xcrun", "stapler", "staple", str(app_path)],
        capture_output=True, text=True,
    )
    if staple.returncode == 0:
        print("Notarization complete and stapled.")
    else:
        print(f"Staple warning: {staple.stderr.strip()}")

    zip_path.unlink(missing_ok=True)
    return True


def build():
    gui_dir = Path(__file__).parent
    project_dir = gui_dir.parent

    clean = "--clean" in sys.argv
    sign = "--sign" in sys.argv or "--sign-only" in sys.argv
    sign_only = "--sign-only" in sys.argv
    notarize = "--notarize" in sys.argv

    # Platform-specific settings
    system = platform.system()
    if system == "Darwin":
        name = "Excel Engine"
        icon_flag = []
        extra = ["--windowed"]
    elif system == "Windows":
        name = "ExcelEngine"
        icon_flag = []
        extra = ["--windowed"]
    else:
        name = "excel-engine-gui"
        icon_flag = []
        extra = []

    if not sign_only:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--name", name,
            "--onedir",
            "--noconfirm",
            "--collect-all", "streamlit",
            "--collect-all", "altair",
            "--collect-all", "excel_engine",
            "--add-data", f"{gui_dir / 'app.py'}{os.pathsep}.",
            *icon_flag,
            *extra,
            str(gui_dir / "run_app.py"),
        ]

        if clean:
            cmd.insert(3, "--clean")

        print(f"Building for {system}...")
        print(f"Command: {' '.join(cmd)}")
        print()

        subprocess.run(cmd, check=True, cwd=str(project_dir))

        print()
        print("Build complete!")
        print(f"  Output: {project_dir / 'dist' / name}")

    # macOS code signing
    if sign and system == "Darwin":
        app_path = project_dir / "dist" / f"{name}.app"
        if not app_path.exists():
            app_path = project_dir / "dist" / name
        entitlements = gui_dir / "entitlements.plist"
        _codesign_app(app_path, entitlements)

        if notarize:
            _notarize_app(app_path)
    elif sign and system != "Darwin":
        print("\nNote: --sign is only supported on macOS")


if __name__ == "__main__":
    build()
