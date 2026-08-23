"""
GL Journal Processor — API entrypoint for Vercel Functions.

Two endpoints:
    POST /match  — upload a GL summary (.xlsx) + paste the prior year's
                   journal (text). Returns matched_csv_rows (final, done)
                   and unmatched_csv_rows (AccountCode blank, needs the
                   caller's LLM step + user review before writing).
    POST /write  — given the final combined list of shaped rows (matched
                   + AccountCode-filled unmatched), returns the CSV file
                   content directly in the response body.

Nothing here touches the filesystem. Vercel Functions are stateless: file
uploads are read into memory, and the output CSV is streamed back in the
HTTP response rather than written to disk anywhere.
"""

from __future__ import annotations

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel, Field

from lib.core import (
    ProcessingError,
    build_csv_rows,
    match_rows,
    parse_ledger,
    parse_prior_journal,
    rows_to_csv_string,
)

app = FastAPI(title="GL Journal Processor")


# --------------------------------------------------------------------------
# Response / request models
# --------------------------------------------------------------------------

class MatchResponse(BaseModel):
    year: int
    ledger_format: str
    matched_csv_rows: list[dict[str, str]]
    unmatched_csv_rows: list[dict[str, str]]
    skipped_prior_rows: int


class WriteRequest(BaseModel):
    rows: list[dict[str, str]] = Field(
        ..., description="matched_csv_rows + AccountCode-filled unmatched_csv_rows, concatenated"
    )


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------

@app.post("/match", response_model=MatchResponse)
async def match(
    gl_summary: UploadFile = File(..., description="GL summary .xlsx file (Xero) or GL report .xlsx file (MYOB)"),
    prior_journal_text: str = Form(..., description="Pasted prior year journal, tab-separated"),
    ledger_format: str | None = Form(
        None, description="Optional override: 'xero' or 'myob'. Auto-detected from the file if omitted."
    ),
) -> MatchResponse:
    try:
        if ledger_format is not None and ledger_format not in ("xero", "myob"):
            raise ProcessingError("ledger_format must be 'xero' or 'myob' if provided.")

        gl_bytes = await gl_summary.read()
        gl_rows, year, detected_format = parse_ledger(
            gl_bytes, filename=gl_summary.filename, ledger_format=ledger_format
        )

        if year is None:
            raise ProcessingError(
                "Could not determine the financial year from the GL "
                "summary's period line. Confirm the file includes a "
                "'For the period ...' line."
            )

        prior_lookup, skipped = parse_prior_journal(prior_journal_text)
        matched, unmatched = match_rows(gl_rows, prior_lookup)

        matched_csv_rows = build_csv_rows(matched, year)
        unmatched_csv_rows = build_csv_rows(unmatched, year)

        return MatchResponse(
            year=year,
            ledger_format=detected_format,
            matched_csv_rows=matched_csv_rows,
            unmatched_csv_rows=unmatched_csv_rows,
            skipped_prior_rows=skipped,
        )
    except ProcessingError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@app.post("/write")
async def write(request: WriteRequest) -> PlainTextResponse:
    if not request.rows:
        raise HTTPException(status_code=422, detail="No rows provided to write.")

    csv_content = rows_to_csv_string(request.rows)
    return PlainTextResponse(
        content=csv_content,
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=journal.csv"},
    )


@app.get("/")
def health() -> dict[str, str]:
    return {"status": "ok"}
