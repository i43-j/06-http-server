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
    GET  /swagger2.json — Swagger 2.0 spec (host/basePath/schemes) for
                   consumers like Copilot Studio's "HTTP with Swagger"
                   action, which requires Swagger 2.0 rather than
                   OpenAPI 3.x. Inlined here (not read from disk) since
                   Vercel's Python bundler doesn't reliably ship loose
                   non-.py files alongside the function.

Nothing here touches the filesystem. Vercel Functions are stateless: file
uploads are read into memory, and the output CSV is streamed back in the
HTTP response rather than written to disk anywhere.
"""

from __future__ import annotations

import json

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse, PlainTextResponse
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
# Swagger 2.0 spec (inlined — see module docstring for why)
# --------------------------------------------------------------------------

SWAGGER2_SPEC_JSON = r'''
{
  "swagger": "2.0",
  "info": {
    "title": "GL Journal Processor",
    "description": "Matches a GL summary export (Xero or MYOB) against a pasted prior-year journal, and writes the final combined rows out as a downloadable CSV import file.",
    "version": "1.0.0"
  },
  "host": "06-http-server-delta.vercel.app",
  "basePath": "/",
  "schemes": [
    "https"
  ],
  "consumes": [
    "multipart/form-data",
    "application/json"
  ],
  "produces": [
    "application/json",
    "text/csv"
  ],
  "paths": {
    "/match": {
      "post": {
        "operationId": "matchGlJournal",
        "summary": "Match a GL summary against the prior year's journal",
        "description": "Uploads a GL summary .xlsx file (Xero 'General Ledger Summary' or MYOB 'General Ledger Report') and a pasted, tab-separated prior-year journal. Returns rows that matched by Description (Account filled in) and rows that did not (Account blank, needs review before /write).",
        "consumes": [
          "multipart/form-data"
        ],
        "parameters": [
          {
            "name": "gl_summary",
            "in": "formData",
            "required": true,
            "type": "file",
            "description": "GL summary .xlsx file \u2014 Xero 'General Ledger Summary' export or MYOB 'General Ledger Report' export."
          },
          {
            "name": "prior_journal_text",
            "in": "formData",
            "required": true,
            "type": "string",
            "description": "Pasted prior year journal, tab-separated, with at least 'Description' and 'Account' columns in the header row."
          },
          {
            "name": "ledger_format",
            "in": "formData",
            "required": false,
            "type": "string",
            "enum": [
              "xero",
              "myob"
            ],
            "description": "Optional override for which ledger format to parse the file as. If omitted, auto-detected."
          }
        ],
        "responses": {
          "200": {
            "description": "GL summary parsed and matched successfully.",
            "schema": {
              "$ref": "#/definitions/MatchResponse"
            }
          },
          "422": {
            "description": "Input could not be processed.",
            "schema": {
              "$ref": "#/definitions/HTTPValidationError"
            }
          }
        }
      }
    },
    "/write": {
      "post": {
        "operationId": "writeGlJournalCsv",
        "summary": "Write the final rows out as a CSV file",
        "description": "Takes the final combined list of shaped rows (matched_csv_rows from /match, plus unmatched_csv_rows with Account filled in and reviewed) and returns the CSV file content ready for import.",
        "consumes": [
          "application/json"
        ],
        "parameters": [
          {
            "name": "body",
            "in": "body",
            "required": true,
            "schema": {
              "$ref": "#/definitions/WriteRequest"
            }
          }
        ],
        "responses": {
          "200": {
            "description": "CSV file content.",
            "schema": {
              "type": "string"
            }
          },
          "422": {
            "description": "No rows were provided to write.",
            "schema": {
              "$ref": "#/definitions/HTTPValidationError"
            }
          }
        }
      }
    },
    "/": {
      "get": {
        "operationId": "healthCheck",
        "summary": "Health check",
        "description": "Returns a simple status payload confirming the service is running.",
        "responses": {
          "200": {
            "description": "Service is healthy.",
            "schema": {
              "type": "object",
              "properties": {
                "status": {
                  "type": "string",
                  "example": "ok"
                }
              }
            }
          }
        }
      }
    }
  },
  "definitions": {
    "CsvRow": {
      "type": "object",
      "description": "A single shaped row matching the CSV import template's columns.",
      "properties": {
        "Narration": {
          "type": "string",
          "example": "2026-01 Load Net Activity"
        },
        "Date": {
          "type": "string",
          "example": "30 Jun 2026"
        },
        "Description": {
          "type": "string",
          "example": "412 - Accounting"
        },
        "Account": {
          "type": "string",
          "description": "Account code. Blank for unmatched rows until filled in.",
          "example": "412"
        },
        "TaxRate": {
          "type": "string",
          "example": "BAS Excluded"
        },
        "Amount": {
          "type": "string",
          "description": "Net amount, debit minus credit. Negative wrapped in parentheses, e.g. '(1,234.56)'.",
          "example": "1,234.56"
        },
        "TrackingName1": {
          "type": "string"
        },
        "TrackingOption1": {
          "type": "string"
        },
        "TrackingName2": {
          "type": "string"
        },
        "TrackingOption2": {
          "type": "string"
        }
      }
    },
    "MatchResponse": {
      "type": "object",
      "required": [
        "year",
        "ledger_format",
        "matched_csv_rows",
        "unmatched_csv_rows",
        "skipped_prior_rows"
      ],
      "properties": {
        "year": {
          "type": "integer",
          "example": 2026
        },
        "ledger_format": {
          "type": "string",
          "enum": [
            "xero",
            "myob"
          ]
        },
        "matched_csv_rows": {
          "type": "array",
          "items": {
            "$ref": "#/definitions/CsvRow"
          }
        },
        "unmatched_csv_rows": {
          "type": "array",
          "items": {
            "$ref": "#/definitions/CsvRow"
          }
        },
        "skipped_prior_rows": {
          "type": "integer"
        }
      }
    },
    "WriteRequest": {
      "type": "object",
      "required": [
        "rows"
      ],
      "properties": {
        "rows": {
          "type": "array",
          "description": "matched_csv_rows plus Account-filled unmatched_csv_rows, concatenated.",
          "items": {
            "$ref": "#/definitions/CsvRow"
          }
        }
      }
    },
    "HTTPValidationError": {
      "type": "object",
      "properties": {
        "detail": {
          "type": "string"
        }
      }
    }
  }
}
'''
SWAGGER2_SPEC = json.loads(SWAGGER2_SPEC_JSON)


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


@app.get("/swagger2.json")
def swagger2() -> JSONResponse:
    return JSONResponse(SWAGGER2_SPEC)


@app.get("/")
def health() -> dict[str, str]:
    return {"status": "ok"}
