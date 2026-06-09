# iClassReport Processor

Turns two iClassPro CSV exports into the canonical **iClassReport** ST Ledgerwood
requires, plus a draft 2090 clearing journal entry and a bank cross-check.
Same hosting pattern as `ninja-park-tool` (GitHub → Streamlit → embed on site).

## Inputs (run both for the same period, 1st → last of the month)

| Upload | iClassPro report | Drives |
|---|---|---|
| FIN-4 | **Program Deposit Split Report** (Payments Received basis) | Open Gym / Birthdays / Retail gross + tax, refunds |
| FIN-24 | **Gateway Transactions Report** | Total Gross / Fees / Net Deposited |

Upload order doesn't matter — the app detects which file is which.

## What it outputs

- **iClassReport** in the exact ST Ledgerwood layout (on-screen + CSV download + Google Sheet).
- **Bank cross-check** — enter the month's iClassPro card payouts that hit 2090; the app shows the
  variance (a small one is expected end-of-month timing).
- **Draft 2090 clearing JE** — revenue split across 4200 / 4000 / 4100 / 2230 / 6040.

## Report logic (locked, validated against April 2026 to the penny)

- **Each taxable program gets its own row** (Open Gym and Birthdays first, then the rest alphabetically).
- Taxable vs non-taxable is detected from the report (tax collected > 0), so new programs classify themselves
  and show up automatically as a new row.
- **Refunds** are read from the `(Refund)` charge-category lines and split taxable / non-taxable.
- **Card Processing Fees** (gross / fees / net deposited) come from the Gateway report, on a report-as-run basis —
  whatever the period shows is reported, so nothing is left behind. Small month-end in-transit differences ride
  into the next month's report.
- **EFA / ClassWallet** is intentionally excluded — it's booked directly to 4100 from the bank, not through 2090.

The draft JE still groups for posting: Open Gym + Birthdays → 4200, other taxable → 4000, non-taxable → 4100.

## Deploy

1. Push these files to a GitHub repo (`app.py`, `requirements.txt`).
2. Create a Streamlit app pointed at `app.py`.
3. (Optional, for Google Sheets output) In Streamlit **Settings → Secrets**, add your Google service-account
   JSON under `gcp_service_account`, and share the target sheet (named `iClassReport_Output`) with the
   service-account email. Without this secret the app still works — it just shows the table and CSV download.

## Notes

- The draft JE's revenue allocations are solid; the received-vs-deposited **cutoff** at month-end is a
  recognition-timing call for ST Ledgerwood. Confirm her preference once, then it's standard each month.
- Version 1.0
