# GL Journal Processor API

Turns a GL summary export (Xero or MYOB) into a ready-to-import journal CSV,
matching each account against a prior year's journal where possible.

Two calls, in order: **`/match`** (upload + match), then **`/write`**
(turn the final, reviewed rows into a CSV file).

---

## 1. `POST /match`

Upload the GL file and paste the prior year's journal. Get back matched
rows (done) and unmatched rows (need an `AccountCode` filled in before
`/write`).

### Send

`multipart/form-data` with:

| Field                | Type          | Required | Notes                                                                 |
|-----------------------|--------------|----------|------------------------------------------------------------------------|
| `gl_summary`           | file (.xlsx) | yes      | The GL summary (Xero) or GL report (MYOB) export.                     |
| `prior_journal_text`   | text         | yes      | Prior year's journal, tab-separated, pasted as plain text.            |
| `ledger_format`        | text         | no       | `"xero"` or `"myob"`. Only needed if auto-detection is inconclusive.  |

**`prior_journal_text` format** — tab-separated, must include `Description`
and `Account` columns (other columns like `Tax Rate` are fine, just ignored):

```
Description	Account	Tax Rate	Debit AUD	Credit AUD
412 - Accounting	300 - Accountancy Fees	BAS Excluded	8410.00	
```

### Example (curl)

```bash
curl -X POST https://your-deployment.vercel.app/match \
  -F "gl_summary=@Yambarran_Pty_Ltd_-_GeneralLedgerReport.xlsx" \
  -F "prior_journal_text@prior_journal.txt"
```

(`prior_journal.txt` is just the tab-separated text above, saved to a file
and passed with `-F "prior_journal_text@prior_journal.txt"` — or pass it
inline with `-F "prior_journal_text=Description\tAccount\n..."`.)

### Get back

```json
{
  "year": 2026,
  "ledger_format": "myob",
  "matched_csv_rows": [
    {
      "Narration": "2026-01 Load Net Activity",
      "Date": "1 Jul 2026",
      "Description": "1-1000 Cheque account",
      "AccountCode": "100 - Cheque Account",
      "TaxRate": "BAS Excluded",
      "Amount": "14,184.80",
      "TrackingName1": "", "TrackingOption1": "",
      "TrackingName2": "", "TrackingOption2": ""
    }
  ],
  "unmatched_csv_rows": [
    {
      "Narration": "2026-01 Load Net Activity",
      "Date": "1 Jul 2026",
      "Description": "4-1000 Fishing Sales",
      "AccountCode": "",
      "TaxRate": "BAS Excluded",
      "Amount": "(516,313.45)",
      "TrackingName1": "", "TrackingOption1": "",
      "TrackingName2": "", "TrackingOption2": ""
    }
  ],
  "skipped_prior_rows": 0
}
```

- **`matched_csv_rows`** — `Description` matched exactly against the prior
  journal; `AccountCode` is already filled in. Ready to write as-is.
- **`unmatched_csv_rows`** — no match found; `AccountCode` is `""`. Fill
  these in (manually, or with your own judgment/LLM step) before writing.
- **`year`** — the financial year start year, read off the period line in
  the GL file. Used to build `Narration`/`Date` on every row.
- **`ledger_format`** — which format was detected (`"xero"` or `"myob"`),
  so you can confirm it picked the right one.

### Errors (HTTP 422)

Returned as `{"detail": "..."}` for anything expected-but-invalid, e.g.:
- The `.xlsx` won't open, or isn't a recognizable GL export.
- `prior_journal_text` is empty or missing the `Description`/`Account`
  columns.
- The financial year couldn't be determined from the file.
- Format couldn't be auto-detected and no `ledger_format` was passed.

---

## 2. `POST /write`

Turn a final, complete list of rows into a downloadable CSV.

### Send

`application/json`:

```json
{
  "rows": [ /* matched_csv_rows + AccountCode-filled unmatched_csv_rows, concatenated */ ]
}
```

Each row must have all ten fields (`Narration`, `Date`, `Description`,
`AccountCode`, `TaxRate`, `Amount`, `TrackingName1`, `TrackingOption1`,
`TrackingName2`, `TrackingOption2`) — exactly the shape `/match` already
returns each row in, so normally you just concatenate the two arrays from
`/match` (after filling in `AccountCode` on the unmatched ones) and send
that straight through.

### Example (curl)

```bash
curl -X POST https://your-deployment.vercel.app/write \
  -H "Content-Type: application/json" \
  -d '{"rows": [ ... ]}' \
  -o journal.csv
```

### Get back

The raw CSV file (`Content-Type: text/csv`, with a
`Content-Disposition: attachment; filename=journal.csv` header) — save the
response body directly to a `.csv` file.

### Errors (HTTP 422)

`{"detail": "No rows provided to write."}` if `rows` is empty.

---

## Typical flow

1. `POST /match` with the GL file + pasted prior journal.
2. Review `unmatched_csv_rows` and fill in `AccountCode` for each (however
   you decide the right account — manually or otherwise).
3. Concatenate `matched_csv_rows` + your now-filled-in `unmatched_csv_rows`.
4. `POST /write` with that combined list → save the returned CSV.

## Health check

`GET /` → `{"status": "ok"}`
