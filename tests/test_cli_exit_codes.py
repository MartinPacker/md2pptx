"""Subprocess tests for md2pptx startup error exit codes.

Error terminations (missing input file, no parameters, empty input)
must exit nonzero so scripts and CI can detect the failure.
"""

import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "md2pptx"


def run_md2pptx(*args):
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
    )


def test_missing_input_file_exits_nonzero(tmp_path):
    result = run_md2pptx("does-not-exist-140.md", str(tmp_path / "out.pptx"))
    assert result.returncode != 0
    assert "does not exist" in result.stdout
    assert "Traceback" not in result.stdout
    assert "Traceback" not in result.stderr


def test_no_parameters_exits_nonzero():
    result = run_md2pptx()
    assert result.returncode != 0
    assert "No parameters" in result.stdout
    assert "Traceback" not in result.stdout
    assert "Traceback" not in result.stderr


def test_empty_input_file_exits_nonzero(tmp_path):
    empty = tmp_path / "empty.md"
    empty.write_text("", encoding="utf-8")
    result = run_md2pptx(str(empty), str(tmp_path / "out.pptx"))
    assert result.returncode != 0
    assert "Empty input file" in result.stdout
    assert "Traceback" not in result.stdout
    assert "Traceback" not in result.stderr
