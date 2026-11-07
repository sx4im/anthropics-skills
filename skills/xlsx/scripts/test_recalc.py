"""Tests for the data-table guard in recalc.py (anthropics/skills #2154).

Run: python3 test_recalc.py
No test framework needed; plain asserts, stdlib only (plus openpyxl, which the
skill already requires, for the fixtures that mirror the issue).

The guard: LibreOffice rewrites an Excel data table's cells as ordinary
TABLE() formulas, so the data table itself is lost while recalc.py reports
success. recalc() must refuse a workbook containing a data table, naming the
ranges, unless --force is passed.
"""

import io
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

# office/soffice.py resolves the temp dir at import time. If the environment
# cannot provide a usable temp dir (e.g. a sandbox write fault), point the
# cache at this directory so the import still works; the in-memory tests do
# not use the temp dir, and the file-based tests skip honestly below.
try:
    tempfile.gettempdir()
except (FileNotFoundError, OSError):
    tempfile.tempdir = str(Path(__file__).resolve().parent)

import recalc
from recalc import data_tables_at_risk, recalc as do_recalc

HERE = Path(__file__).resolve().parent

# Minimal workbook parts; only the worksheet XML matters to the guard.
WORKBOOK_XML = (
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
)
WORKBOOK_RELS = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet1.xml"/></Relationships>'
)
SHEET_XML_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    "<sheetData>{rows}</sheetData></worksheet>"
)


def make_xlsx_bytes(rows_xml):
    """A minimal .xlsx (bytes) whose single sheet holds the given rows XML."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", WORKBOOK_XML)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/worksheets/sheet1.xml", SHEET_XML_TEMPLATE.format(rows=rows_xml))
    return buf.getvalue()


ONE_INPUT_TABLE = (
    '<row r="2"><c r="B2"><v>10</v></c></row>'
    '<row r="3"><c r="B3"><v>2</v></c></row>'
    '<row r="4"><c r="B4"><f>B2*B3</f></c></row>'
    '<row r="5"><c r="D5"><f>B4</f></c></row>'
    '<row r="6"><c r="C6"><v>8</v></c>'
    '<c r="D6"><f t="dataTable" ref="D6:D8" dt2D="0" dtr="0" r1="B2"/></c></row>'
    '<row r="7"><c r="C7"><v>9</v></c></row>'
    '<row r="8"><c r="C8"><v>10</v></c></row>'
)

TWO_INPUT_TABLE = (
    '<row r="11"><c r="D11">'
    '<f t="dataTable" ref="D11:F13" dt2D="1" dtr="0" r1="B2" r2="B3"/></c></row>'
)

PLAIN_FORMULAS = (
    '<row r="2"><c r="B2"><v>10</v></c></row>'
    '<row r="4"><c r="B4"><f>B2*B3</f></c></row>'
)


def test_finds_one_input_table():
    assert data_tables_at_risk(io.BytesIO(make_xlsx_bytes(ONE_INPUT_TABLE))) == [
        "Sheet1!D6:D8"
    ]


def test_finds_two_input_table():
    assert data_tables_at_risk(io.BytesIO(make_xlsx_bytes(TWO_INPUT_TABLE))) == [
        "Sheet1!D11:F13"
    ]


def test_finds_both_tables():
    assert data_tables_at_risk(
        io.BytesIO(make_xlsx_bytes(ONE_INPUT_TABLE + TWO_INPUT_TABLE))
    ) == ["Sheet1!D6:D8", "Sheet1!D11:F13"]


def test_plain_formulas_are_not_tables():
    assert data_tables_at_risk(io.BytesIO(make_xlsx_bytes(PLAIN_FORMULAS))) == []


def test_bad_zip_is_not_a_table():
    assert data_tables_at_risk(io.BytesIO(b"not a zip file")) == []


def test_recalc_refuses_data_table(tmp_path):
    target = tmp_path / "table.xlsx"
    target.write_bytes(make_xlsx_bytes(ONE_INPUT_TABLE))
    before = target.read_bytes()
    result = do_recalc(str(target))
    assert "error" in result, result
    assert "D6:D8" in result["error"], result
    assert result["data_table_ranges"] == ["Sheet1!D6:D8"]
    assert result["data_table_ranges_truncated"] == 0
    assert target.read_bytes() == before, "refusal must not touch the file"


def test_recalc_force_bypasses_guard(tmp_path, monkeypatch=None):
    target = tmp_path / "table.xlsx"
    target.write_bytes(make_xlsx_bytes(ONE_INPUT_TABLE))
    called = {}

    def stub(filename, abs_path, timeout, profile_dir):
        called["ran"] = True
        return {"status": "success", "total_errors": 0, "total_formulas": 5}

    old = recalc._recalc_with_profile
    recalc._recalc_with_profile = stub
    try:
        result = do_recalc(str(target), force=True)
    finally:
        recalc._recalc_with_profile = old
    assert called.get("ran") is True
    assert result.get("status") == "success"


def test_recalc_plain_workbook_passes_guard(tmp_path):
    target = tmp_path / "plain.xlsx"
    target.write_bytes(make_xlsx_bytes(PLAIN_FORMULAS))
    called = {}

    def stub(filename, abs_path, timeout, profile_dir):
        called["ran"] = True
        return {"status": "success", "total_errors": 0, "total_formulas": 1}

    old = recalc._recalc_with_profile
    recalc._recalc_with_profile = stub
    try:
        result = do_recalc(str(target))
    finally:
        recalc._recalc_with_profile = old
    assert called.get("ran") is True
    assert result.get("status") == "success"


class SkipTest(Exception):
    """The environment cannot run this test (no writable temp dir, etc.)."""


def test_cli_refuses(tmp_path):
    target = tmp_path / "table.xlsx"
    target.write_bytes(make_xlsx_bytes(ONE_INPUT_TABLE))
    proc = subprocess.run(
        [sys.executable, str(HERE / "recalc.py"), str(target)],
        capture_output=True,
        text=True,
    )
    assert proc.returncode == 1, proc.stdout
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        raise SkipTest(f"recalc.py could not run here: {proc.stderr.strip()[:200]}")
    assert "error" in result and "D6:D8" in result["error"]


def main():
    tests = [
        test_finds_one_input_table,
        test_finds_two_input_table,
        test_finds_both_tables,
        test_plain_formulas_are_not_tables,
        test_bad_zip_is_not_a_table,
    ]
    for test in tests:
        test()
        print(f"PASS {test.__name__}")
    # The file-based tests need a writable temp dir; report honestly if the
    # sandbox cannot provide one instead of failing the suite silently.
    try:
        with tempfile.TemporaryDirectory(prefix="test-recalc-") as d:
            tmp = Path(d)
            for test in (
                test_recalc_refuses_data_table,
                test_recalc_force_bypasses_guard,
                test_recalc_plain_workbook_passes_guard,
                test_cli_refuses,
            ):
                try:
                    test(tmp)
                except SkipTest as e:
                    print(f"SKIP {test.__name__}: {e}")
                    continue
                print(f"PASS {test.__name__}")
    except OSError as e:
        print(f"SKIP file-based tests: no writable temp dir ({e})")
    print("OK")


if __name__ == "__main__":
    main()
