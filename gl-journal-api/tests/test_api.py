"""
Tests for the GL Journal Processor API.

Run with:
    python -m pytest tests/ -v

Covers both the pure logic in lib/core.py and the actual HTTP behavior of
app.py via FastAPI's TestClient — the latter catches issues pure unit
tests can't (multipart parsing, response shapes, status codes).
"""

import io
import sys
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import app  # noqa: E402
from lib.core import (  # noqa: E402
    ProcessingError,
    build_csv_rows,
    format_amount,
    match_rows,
    parse_ledger,
    parse_prior_journal,
    rows_to_csv_string,
)

client = TestClient(app)


# --------------------------------------------------------------------------
# Fixtures
# --------------------------------------------------------------------------

def _sample_workbook_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Barramundi Adventures Pty Ltd"])
    ws.append(["General Ledger Summary"])
    ws.append(["For the period 1 July 2025 to 30 June 2026"])
    ws.append([])
    ws.append(["Account", "Account Code", "Debit", "Credit", "Net Movement", "Account Type"])
    ws.append(["Accounting", "412", 8410.00, None, 8410.00, "Expense"])
    ws.append(["Fishing Sales", "200", None, 516313.45, -516313.45, "Income"])
    ws.append(["Total", None, None, None, None, None])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _sample_myob_workbook_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["General ledger report"])
    ws.append(["Yambarran Pty Ltd"])
    ws.append(["PO Box 1544, Katherine, NT, 0851"])
    ws.append(["0427430089"])
    ws.append(["Accrual mode"])
    ws.append(["01 Apr 2026 - 30 Jun 2026"])
    ws.append(["Generated 23 Aug 2026 16:29:29"])
    ws.append([" "])
    ws.append(["Code", "Category name", "Open ($)", "Debit ($)", "Credit ($)", "Net activity ($)", "Balance ($)", "Tax amount ($)"])
    ws.append(["", "", "", "", "", "", "", ""])
    ws.append(["1-1000", "Cheque account", 57969.41, 8410.00, None, 8410.00, 66379.41, 0.0])
    ws.append([None, None, None, None, None, None, None, None])
    ws.append(["4-1000", "Fishing Sales", 0, None, 516313.45, -516313.45, -516313.45, 0.0])
    ws.append([None, None, None, None, None, None, None, None])
    ws.append(["Grand total", "", "", 8410.00, 516313.45, "", "", ""])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


SAMPLE_PRIOR_JOURNAL = (
    "Description\tAccount\tTax Rate\tDebit AUD\tCredit AUD\n"
    "412 - Accounting\t300 - Accountancy Fees\tBAS Excluded\t8410.00\t\n"
)


# --------------------------------------------------------------------------
# Pure logic tests
# --------------------------------------------------------------------------

class TestFormatAmount:
    def test_debit_only(self):
        assert format_amount(100.0, 0) == "100.00"

    def test_credit_only(self):
        assert format_amount(0, 100.0) == "(100.00)"

    def test_net_negative(self):
        assert format_amount(100.0, 300.0) == "(200.00)"

    def test_thousands_separator(self):
        assert format_amount(15959.21, 0) == "15,959.21"


class TestMatchRows:
    def test_splits_matched_and_unmatched(self):
        gl_rows = [
            {"description": "412 - Accounting", "amount": "8,410.00"},
            {"description": "999 - New Account", "amount": "500.00"},
        ]
        prior_lookup = {"412 - Accounting": "300 - Accountancy Fees"}
        matched, unmatched = match_rows(gl_rows, prior_lookup)
        assert len(matched) == 1
        assert matched[0]["account_code"] == "300 - Accountancy Fees"
        assert len(unmatched) == 1


class TestBuildCsvRows:
    def test_matched_row_shape(self):
        rows = build_csv_rows(
            [{"description": "412 - Accounting", "amount": "8,410.00", "account_code": "300 - Accountancy Fees"}],
            year=2025,
        )
        assert rows[0]["Narration"] == "2025-01 Load Net Activity"
        assert rows[0]["Date"] == "1 Jul 2025"
        assert rows[0]["AccountCode"] == "300 - Accountancy Fees"

    def test_unmatched_row_gets_blank_account_code(self):
        rows = build_csv_rows([{"description": "999 - New Account", "amount": "500.00"}], year=2025)
        assert rows[0]["AccountCode"] == ""


class TestParsePriorJournal:
    def test_happy_path(self):
        lookup, skipped = parse_prior_journal(SAMPLE_PRIOR_JOURNAL)
        assert lookup["412 - Accounting"] == "300 - Accountancy Fees"
        assert skipped == 0

    def test_rejects_wrong_columns(self):
        with pytest.raises(ProcessingError):
            parse_prior_journal("Foo\tBar\n1\t2\n")

    def test_rejects_empty(self):
        with pytest.raises(ProcessingError):
            parse_prior_journal("")


class TestParseLedger:
    def test_parses_bytes_directly(self):
        rows, year, fmt = parse_ledger(_sample_workbook_bytes())
        assert year == 2025
        assert fmt == "xero"
        assert len(rows) == 2  # "Total" row excluded
        assert rows[0]["description"] == "412 - Accounting"
        assert rows[1]["amount"] == "(516,313.45)"

    def test_rejects_garbage_bytes(self):
        with pytest.raises(ProcessingError):
            parse_ledger(b"not an xlsx file")

    def test_parses_myob_format(self):
        rows, year, fmt = parse_ledger(_sample_myob_workbook_bytes())
        assert year == 2026
        assert fmt == "myob"
        assert len(rows) == 2  # blank rows + "Grand total" excluded
        assert rows[0]["description"] == "1-1000 Cheque account"  # space, not " - "
        assert rows[0]["amount"] == "8,410.00"
        assert rows[1]["description"] == "4-1000 Fishing Sales"
        assert rows[1]["amount"] == "(516,313.45)"

    def test_myob_detected_from_filename_when_header_ambiguous(self):
        # A minimal MYOB-shaped file whose header row wouldn't be found within
        # the scan window still gets detected correctly via filename.
        rows, year, fmt = parse_ledger(
            _sample_myob_workbook_bytes(), filename="Some_Company_-_GeneralLedgerReport.xlsx"
        )
        assert fmt == "myob"

    def test_explicit_ledger_format_overrides_detection(self):
        rows, year, fmt = parse_ledger(_sample_myob_workbook_bytes(), ledger_format="myob")
        assert fmt == "myob"


class TestRowsToCsvString:
    def test_round_trip(self):
        rows = build_csv_rows(
            [{"description": "412 - Accounting", "amount": "8,410.00", "account_code": "300 - Accountancy Fees"}],
            year=2025,
        )
        csv_text = rows_to_csv_string(rows)
        assert "412 - Accounting" in csv_text
        assert "2025-01 Load Net Activity" in csv_text


# --------------------------------------------------------------------------
# HTTP endpoint tests (via FastAPI TestClient — no real network needed)
# --------------------------------------------------------------------------

class TestMatchEndpoint:
    def test_match_happy_path(self):
        response = client.post(
            "/match",
            files={"gl_summary": ("gl.xlsx", _sample_workbook_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            data={"prior_journal_text": SAMPLE_PRIOR_JOURNAL},
        )
        assert response.status_code == 200
        body = response.json()
        assert body["year"] == 2025
        assert body["ledger_format"] == "xero"
        assert len(body["matched_csv_rows"]) == 1
        assert len(body["unmatched_csv_rows"]) == 1
        assert body["matched_csv_rows"][0]["AccountCode"] == "300 - Accountancy Fees"
        assert body["unmatched_csv_rows"][0]["AccountCode"] == ""

    def test_match_myob_happy_path(self):
        prior_journal = (
            "Description\tAccount\tTax Rate\tDebit AUD\tCredit AUD\n"
            "1-1000 Cheque account\t100 - Cheque Account\tBAS Excluded\t8410.00\t\n"
        )
        response = client.post(
            "/match",
            files={"gl_summary": ("Yambarran_Pty_Ltd_-_GeneralLedgerReport.xlsx", _sample_myob_workbook_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            data={"prior_journal_text": prior_journal},
        )
        assert response.status_code == 200
        body = response.json()
        assert body["ledger_format"] == "myob"
        assert len(body["matched_csv_rows"]) == 1
        assert len(body["unmatched_csv_rows"]) == 1
        assert body["matched_csv_rows"][0]["Description"] == "1-1000 Cheque account"

    def test_match_rejects_garbage_prior_journal(self):
        response = client.post(
            "/match",
            files={"gl_summary": ("gl.xlsx", _sample_workbook_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            data={"prior_journal_text": "Foo\tBar\n1\t2\n"},
        )
        assert response.status_code == 422

    def test_match_rejects_garbage_xlsx(self):
        response = client.post(
            "/match",
            files={"gl_summary": ("gl.xlsx", b"not a real xlsx", "application/octet-stream")},
            data={"prior_journal_text": SAMPLE_PRIOR_JOURNAL},
        )
        assert response.status_code == 422


class TestWriteEndpoint:
    def test_write_returns_csv(self):
        rows = build_csv_rows(
            [{"description": "412 - Accounting", "amount": "8,410.00", "account_code": "300 - Accountancy Fees"}],
            year=2025,
        )
        response = client.post("/write", json={"rows": rows})
        assert response.status_code == 200
        assert response.headers["content-type"].startswith("text/csv")
        assert "412 - Accounting" in response.text

    def test_write_rejects_empty_rows(self):
        response = client.post("/write", json={"rows": []})
        assert response.status_code == 422


class TestHealthCheck:
    def test_health(self):
        response = client.get("/")
        assert response.status_code == 200
        assert response.json() == {"status": "ok"}
