"""Tests for office.soffice binary location (issue #2185).

Run from skills/xlsx/scripts:
    python3 -m unittest office.test_soffice -v
"""

import os
import stat
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from office.soffice import find_soffice, run_soffice


class TestFindSoffice(unittest.TestCase):
    def test_uses_path_first(self):
        # This sandbox has soffice on PATH; the resolver must find it.
        found = find_soffice()
        self.assertIsNotNone(found)
        self.assertTrue(found.endswith("soffice"))

    def test_falls_back_to_known_location(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "soffice"
            fake.write_text("#!/bin/sh\nexit 0\n")
            fake.chmod(fake.stat().st_mode | stat.S_IXUSR)
            with (
                mock.patch("shutil.which", return_value=None),
                mock.patch(
                    "office.soffice.KNOWN_SOFFICE_PATHS", (str(fake),)
                ),
            ):
                self.assertEqual(find_soffice(), str(fake))

    def test_returns_none_when_absent(self):
        with (
            mock.patch("shutil.which", return_value=None),
            mock.patch(
                "office.soffice.KNOWN_SOFFICE_PATHS",
                ("/nonexistent/soffice",),
            ),
        ):
            self.assertIsNone(find_soffice())

    def test_run_soffice_raises_helpful_error_when_missing(self):
        with mock.patch("office.soffice.find_soffice", return_value=None):
            with self.assertRaises(FileNotFoundError) as ctx:
                run_soffice(["--version"])
        self.assertIn("LibreOffice", str(ctx.exception))

    def test_run_soffice_uses_located_binary(self):
        with (
            mock.patch(
                "office.soffice.find_soffice", return_value="/fake/soffice"
            ),
            mock.patch(
                "office.soffice.subprocess.run"
            ) as run_mock,
        ):
            run_soffice(["--headless", "--version"])
        argv = run_mock.call_args.args[0]
        self.assertEqual(argv[0], "/fake/soffice")


if __name__ == "__main__":
    unittest.main()
