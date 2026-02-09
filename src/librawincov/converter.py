"""Core conversion functionality for DOCX to PDF."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path


def is_wsl() -> bool:
    """Check if running inside Windows Subsystem for Linux."""
    try:
        with open("/proc/version", "r") as f:
            return "microsoft" in f.read().lower()
    except FileNotFoundError:
        return False


def find_libreoffice() -> tuple[Path, bool]:
    """
    Find LibreOffice installation.

    Returns:
        A tuple of (path, use_wsl) where path is the LibreOffice executable
        and use_wsl indicates whether to use WSL for conversion.

    Raises:
        FileNotFoundError: If LibreOffice cannot be found.
    """
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path:
        candidate = Path(env_path)
        if candidate.exists():
            return candidate, False
        raise FileNotFoundError(f"LIBREOFFICE_PATH not found: {candidate}")

    if platform.system() == "Windows":
        for name in ("soffice.exe", "soffice"):
            found = shutil.which(name)
            if found:
                return Path(found), False

        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for raw in candidates:
            candidate = Path(raw)
            if candidate.exists():
                return candidate, False

        if shutil.which("wsl"):
            return Path("wsl"), True

    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return Path(found), False

    raise FileNotFoundError(
        "LibreOffice not found. Install it, set LIBREOFFICE_PATH, or use WSL."
    )


def _windows_to_wsl_path(win_path: Path) -> str:
    """Convert Windows path to WSL path."""
    path_str = str(win_path.resolve())

    if len(path_str) >= 2 and path_str[1] == ":":
        drive = path_str[0].lower()
        rest = path_str[2:].replace("\\", "/")
        return f"/mnt/{drive}{rest}"

    return path_str.replace("\\", "/")


def _build_convert_command(
    soffice_path: Path, input_path: Path, out_dir: Path, use_wsl: bool
) -> list[str]:
    """Build the LibreOffice conversion command."""
    if use_wsl:
        wsl_input = _windows_to_wsl_path(input_path)
        wsl_out_dir = _windows_to_wsl_path(out_dir)

        return [
            "wsl",
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            wsl_out_dir,
            wsl_input,
        ]
    else:
        return [
            str(soffice_path),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(input_path),
        ]


def convert_docx_to_pdf(input_path: Path, output_path: Path) -> None:
    """
    Convert a DOCX file to PDF using LibreOffice.

    Args:
        input_path: Path to the input DOCX file.
        output_path: Path where the output PDF should be saved.

    Raises:
        RuntimeError: If the conversion fails.
        FileNotFoundError: If LibreOffice cannot be found.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    soffice_path, use_wsl = find_libreoffice()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_out_dir = Path(tmp_dir)
        command = _build_convert_command(soffice_path, input_path, tmp_out_dir, use_wsl)
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode != 0:
            stderr = result.stderr.strip() or "LibreOffice conversion failed."
            raise RuntimeError(stderr)

        tmp_pdf = tmp_out_dir / f"{input_path.stem}.pdf"
        if not tmp_pdf.exists():
            raise RuntimeError("LibreOffice did not produce a PDF output.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(tmp_pdf), str(output_path))
