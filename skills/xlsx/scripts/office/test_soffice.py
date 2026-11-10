"""Tests for office.soffice binary discovery.

Run from skills/xlsx/scripts (so `office` is importable):
    python -m unittest office.test_soffice -v
"""

import os
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from office import soffice


def _make_fake_soffice(directory):
    exe = Path(directory) / "soffice"
    exe.write_text("#!/bin/sh\necho fake-soffice \"$@\"\n")
    exe.chmod(exe.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    return exe


class FindSofficeTests(unittest.TestCase):
    def test_finds_soffice_on_path(self):
        with tempfile.TemporaryDirectory() as d:
            exe = _make_fake_soffice(d)
            with mock.patch.dict(os.environ, {"PATH": d, "SOFFICE_PATH": ""}):
                self.assertEqual(soffice.find_soffice(), str(exe))

    def test_soffice_path_override_wins_over_path(self):
        with tempfile.TemporaryDirectory() as d:
            exe = _make_fake_soffice(d)
            with mock.patch.dict(os.environ, {"PATH": "/nonexistent", "SOFFICE_PATH": str(exe)}):
                self.assertEqual(soffice.find_soffice(), str(exe))

    def test_finds_macos_app_bundle_when_not_on_path(self):
        with tempfile.TemporaryDirectory() as d:
            bundle = Path(d) / "LibreOffice.app" / "Contents" / "MacOS"
            bundle.mkdir(parents=True)
            exe = _make_fake_soffice(bundle)
            with mock.patch.dict(os.environ, {"PATH": "/nonexistent", "SOFFICE_PATH": ""}):
                with mock.patch.object(soffice, "_KNOWN_SOFFICE_PATHS", [str(exe)]):
                    self.assertEqual(soffice.find_soffice(), str(exe))

    def test_returns_none_when_nothing_found(self):
        with mock.patch.dict(os.environ, {"PATH": "/nonexistent", "SOFFICE_PATH": ""}):
            with mock.patch.object(soffice, "_KNOWN_SOFFICE_PATHS", []):
                self.assertIsNone(soffice.find_soffice())


class RunSofficeTests(unittest.TestCase):
    def test_uses_resolved_binary_not_bare_name(self):
        with tempfile.TemporaryDirectory() as d:
            exe = _make_fake_soffice(d)
            with mock.patch.dict(os.environ, {"PATH": d, "SOFFICE_PATH": ""}):
                with mock.patch("office.soffice.subprocess.run") as run:
                    run.return_value.returncode = 0
                    soffice.run_soffice(["--version"], capture_output=True)
                    argv = run.call_args.args[0]
                    self.assertEqual(argv[0], str(exe))
                    self.assertIn("--version", argv)

    def test_missing_binary_raises_actionable_error(self):
        with mock.patch.dict(os.environ, {"PATH": "/nonexistent", "SOFFICE_PATH": ""}):
            with mock.patch.object(soffice, "_KNOWN_SOFFICE_PATHS", []):
                with self.assertRaises(FileNotFoundError) as ctx:
                    soffice.run_soffice(["--version"], capture_output=True)
                self.assertIn("LibreOffice", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
