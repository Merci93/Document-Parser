"""Tests for the CLI module."""
import subprocess
import sys


def test_cli_help():
    result = subprocess.run(
        [sys.executable, "cli.py", "--help"],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0
    assert "Document Parser CLI" in result.stdout
