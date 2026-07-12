"""
iClassReport Processor  -  Flip Side Lowell, LLC
================================================
Upload the payout-basis FIN-4 export and get the canonical iClassReport
ST Ledgerwood requires, a draft 2090 clearing JE, and exact validation
against the merchant statement.

Input:  FIN-4 Program Deposit Split Report, exported with the saved
        PAYOUT DATE preset for the calendar month being closed
        (CSV or XLSX). Payout basis natively matches what the bank
        received, so no date math and no Gateway (FIN-24) report needed.

Validation: enter the four numbers exactly as they appear in the merchant
statement's Summary box (Total Sales, Total Refunds, Total Processing
Fees, Net Amount Settled). Everything must tie to the penny.

Validated against June 2026 (ties to the merchant statement, bank, and
Xero to the penny).  Version 1.4
"""
import csv
import io
import re
import datetime

import pandas as pd
import streamlit as st

# --- CONFIGURATION ---------------------------------------------------------
VERSION = "1.5"

# Open Gym and Birthdays lead the report; every other taxable program follows
# (alphabetical), one row each. Taxable vs non-taxable is detected from the
# report (tax collected > 0), so new programs classify themselves.
OPEN_GYM_NAMES = {"open gym"}
BIRTHDAY_NAMES = {"birthdays", "birthday"}

# FIN-4 tender columns that settle through the iClassPro gateway (these are
# the dollars that arrive as "iClassPro Inc. PAYOUT" in the bank / 2090).
GATEWAY_TENDERS = ["Credit Card", "Credit Card - Swipe/insert/tap",
                   "Credit Card Present", "eCheck"]
CASH_TENDERS = ["Cash", "Check"]
EXTERNAL_TENDERS = ["External Credit Card", "Nacha"]

st.set_page_config(page_title=f"iClassReport Processor {VERSION}", page_icon="N", layout="wide")


# =========================================================================== #
#  CORE LOGIC  (pure functions)
# =========================================================================== #
def parse_money(s):
    """'$1,234.56' / '$(98.55)' / '--' / '' -> float."""
    s = str(s).replace("\xa0", " ").replace("$", "").replace(",", "").strip()
    if s in ("", "--"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def _read_rows(raw_text):
    raw_text = raw_text.lstrip("\ufeff")
    return list(csv.reader(io.StringIO(raw_text)))


def load_upload(uploaded_file):
    """Streamlit UploadedFile (.csv or .xlsx) -> list-of-rows of strings."""
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith((".xlsx", ".xls")):
        try:
            from openpyxl import load_workbook
        except ImportError:
            st.error("Excel uploads need `openpyxl` - add a line saying `openpyxl` "
                     "to requirements.txt in GitHub, or export as CSV instead.")
            return None
        wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=True)
        rows = []
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            sheet_rows = [["" if c is None else str(c) for c in r]
                          for r in ws.iter_rows(values_only=True)]
            if any(any(c.strip() for c in r) for r in sheet_rows):
                rows = sheet_rows
                break
        if not rows:
            st.error(f"'{uploaded_file.name}' contains no data. If this was a "
                     "payout-filtered Gateway export, that's expected - the gateway "
                     "report can't filter by payout date (Batch Date is unpopulated). "
                     "Use statement fees instead.")
            return None
        return rows
    return _read_rows(uploaded_file.read().decode("utf-8-sig", errors="ignore"))


def parse_statement_pdf(uploaded_file):
    """iClassPro merchant statement PDF -> dict with the Summary box numbers,
    or None if it can't be read."""
    try:
        from pypdf import PdfReader
    except ImportError:
        st.error("Statement upload needs `pypdf` - add a line saying `pypdf==5.4.0` to "
                 "requirements.txt in GitHub, or enter the numbers manually below.")
        return None
    try:
        uploaded_file.seek(0)
        text = "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(uploaded_file.read())).pages)
    except Exception:
        return None

    def grab(pat):
        m = re.search(pat, text)
        return float(m.group(1).replace(",", "")) if m else None

    out = {
        "total_sales": grab(r"Total Sales\s*\$([\d,]+\.\d{2})"),
        "refunds": grab(r"Total Refunds\s*\(\$([\d,]+\.\d{2})\)"),
        "fees": grab(r"Total Processing Fees\s*\(\$([\d,]+\.\d{2})\)"),
        "net": grab(r"Net Amount Settled\s*\$([\d,]+\.\d{2})"),
    }
    if out["refunds"] is None and re.search(r"Total Refunds\s*\$0\.00", text):
        out["refunds"] = 0.0
    m = re.search(r"Statement Period:\s*(\d{2}/\d{2}/\d{4})\s*to\s*(\d{2}/\d{2}/\d{4})", text)
    out["period"] = f"{m.group(1)} - {m.group(2)}" if m else ""
    if out["total_sales"] is None or out["fees"] is None or out["net"] is None:
        return None
    if out["refunds"] is None:
        out["refunds"] = 0.0
    return out


def detect_report_type(rows):
    """Return 'gateway', 'program_split', or None by inspecting the first rows."""
    flat = " ".join(" ".join(str(c) for c in r) for r in rows[:8]).lower()
    if "final settlement amount" in flat or "transaction id" in flat:
        return "gateway"
    if "charge category" in flat or "payments received" in flat:
        return "program_split"
    return None


def parse_program_split(rows):
    """FIN-4 -> (dict program -> {total, tax, gross, refunds[], gw, cash, external},
                 unapplied_gw_net, period_str)

    gw / cash / external are the Program Total row's tender-column dollars:
      gw       = tenders that settle through the iClassPro gateway
      cash     = Cash + Check (deposited manually, hits the bank next month)
      external = External Credit Card + Nacha (settled outside the gateway)
    unapplied_gw_net = net gateway-tender dollars on the Unapplied rows.
    """
    period = ""
    for row in rows[:6]:
        for cell in row:
            m = re.search(r"\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4}", str(cell))
            if m:
                period = m.group(0)
                break
        if period:
            break

    hidx = None
    for i, r in enumerate(rows[:10]):
        low = [str(c).strip().lower() for c in r]
        if "program" in low and "charge category" in low:
            hidx = i
            break
    if hidx is None:
        hidx = 1
    header = [c.strip() for c in rows[hidx]]

    def col(name, default=None):
        for i, c in enumerate(header):
            if c.strip().lower() == name.lower():
                return i
        return default

    i_prog = col("Program", 1)
    i_cat = col("Charge Category", 2)
    i_total = col("Total", 3)
    i_tax = col("Taxes", 4)
    i_gw = [col(n) for n in GATEWAY_TENDERS if col(n) is not None]
    i_cash = [col(n) for n in CASH_TENDERS if col(n) is not None]
    i_ext = [col(n) for n in EXTERNAL_TENDERS if col(n) is not None]

    def tender_sum(row, idxs):
        return sum(parse_money(row[i]) for i in idxs if len(row) > i)

    progs = {}
    unapplied_gw_net = 0.0
    unapplied_gw_refunds = 0.0
    cur = None
    for r in rows[hidx + 1:]:
        if len(r) <= i_cat:
            continue
        name = r[i_prog].strip()
        cat = r[i_cat].strip()
        if name.lower().startswith("total payments received"):
            break
        if name.lower().startswith("unapplied"):
            amt_gw = tender_sum(r, i_gw)
            unapplied_gw_net += amt_gw
            if "refund" in name.lower() and "deleted" not in name.lower():
                unapplied_gw_refunds += amt_gw
            continue
        if cat == "Program Total:":
            cur = name
            progs[cur] = {
                "total": parse_money(r[i_total]),
                "tax": parse_money(r[i_tax]),
                "gw": tender_sum(r, i_gw),
                "cash": tender_sum(r, i_cash),
                "external": tender_sum(r, i_ext),
                "gross": 0.0,
                "refunds": [],
            }
        elif cur and name == "" and cat:
            amt = parse_money(r[i_total])
            if "refund" in cat.lower():
                progs[cur]["refunds"].append(amt)
            else:
                progs[cur]["gross"] += amt
    return progs, round(unapplied_gw_net, 2), round(unapplied_gw_refunds, 2), period


def build_iclassreport(progs, unapplied_gw_net, gateway, use_tax=0.0):
    taxable = [p for p, v in progs.items() if abs(v["tax"]) > 0.001]
    open_gym = [p for p in progs if p.strip().lower() in OPEN_GYM_NAMES]
    birthdays = [p for p in progs if p.strip().lower() in BIRTHDAY_NAMES]

    ref_taxable = sum(-x for p in taxable for x in progs[p]["refunds"])
    ref_nontaxable = sum(-x for p in progs if p not in taxable for x in progs[p]["refunds"])

    # one row per taxable program: Open Gym, Birthdays, then the rest alphabetically
    ordered = list(open_gym) + list(birthdays) + sorted(
        p for p in taxable if p not in open_gym and p not in birthdays)
    program_rows = [(p, round(progs[p]["gross"], 2), round(progs[p]["tax"], 2)) for p in ordered]
    total_tax = round(sum(progs[p]["tax"] for p in taxable), 2)

    # ---- Draft 2090 clearing JE - GATEWAY-SETTLED TENDERS ONLY -------------
    # Each program's tax is allocated to the gateway share pro-rata (most
    # programs are 100% card, so this only matters for mixed cash/card ones).
    retail_progs = [p for p in taxable if p not in open_gym and p not in birthdays]

    def gw_tax(p):
        v = progs[p]
        return v["tax"] * (v["gw"] / v["total"]) if abs(v["total"]) > 0.001 else 0.0

    venue = round(sum(progs[p]["gw"] - gw_tax(p) for p in open_gym + birthdays), 2)
    merch = round(sum(progs[p]["gw"] - gw_tax(p) for p in retail_progs), 2)
    member = round(sum(progs[p]["gw"] for p in progs
                       if p not in taxable) + unapplied_gw_net, 2)
    je_tax = round(sum(gw_tax(p) for p in progs), 2)

    je_credits = round(venue + merch + member + je_tax, 2)
    je_debits = round(gateway["net"] + gateway["fees"], 2)
    # absorb sub-dollar pro-rata rounding into the largest revenue line
    rounding = round(je_debits - je_credits, 2)
    if abs(rounding) <= 1.00 and abs(rounding) > 0.001:
        member = round(member + rounding, 2)
        je_credits = round(venue + merch + member + je_tax, 2)

    cash_collected = round(sum(v["cash"] for v in progs.values()), 2)
    external_total = round(sum(v["external"] for v in progs.values()), 2)

    notes = []
    extra_taxable = [p for p in taxable if p not in open_gym and p not in birthdays]
    if extra_taxable:
        notes.append("Taxable programs beyond Open Gym / Birthdays: " + ", ".join(sorted(extra_taxable)))
    if gateway.get("unsettled_count"):
        notes.append(f"{gateway['unsettled_count']} gateway transaction(s) NOT settled "
                     f"(${gateway['unsettled_amt']:,.2f}) - processed but not completed.")
    if abs(unapplied_gw_net) > 0.001:
        notes.append(f"Unapplied payments/refunds via gateway tenders net to "
                     f"${unapplied_gw_net:,.2f} (included in the Member Fees JE line).")
    if abs(cash_collected) > 0.001:
        notes.append(f"Cash/Check collected this window: ${cash_collected:,.2f}. NOT in this JE - "
                     "deposit it, code the deposit to 2090, and allocate it next month.")
    if abs(external_total) > 0.001:
        notes.append(f"External Credit Card / Nacha payments: ${external_total:,.2f}. NOT in this JE - "
                     "these settle outside the iClassPro gateway (e.g. ClassWallet) and are coded "
                     "straight from the bank feed.")

    return {
        "program_rows": program_rows,
        "ref_taxable": round(ref_taxable, 2),
        "ref_nontaxable": round(ref_nontaxable, 2),
        "fees_section": {
            "Total Collected (Gross)": gateway["gross"],
            "Total Fees": gateway["fees"],
            "Total Net Deposited": gateway["net"],
        },
        "use_tax": round(use_tax, 2),
        "total_tax": total_tax,
        "je_debits": [
            ("2090 Revenue Clearing (iClassPro)", round(gateway["net"], 2)),
            ("6040 Credit Card Merchant Fees", round(gateway["fees"], 2)),
        ],
        "je_credits": [
            ("4200 Sales:Venue (Open Gym + Birthdays)", venue),
            ("4000 Sales:Merchandise (other taxable)", merch),
            ("4100 Sales:Member Fees (non-taxable + unapplied)", member),
            ("2230 Sales Tax Payable", je_tax),
        ],
        "je_balance_ok": abs(je_debits - je_credits) < 0.005,
        "je_debit_total": je_debits,
        "je_credit_total": je_credits,
        "cash_collected": cash_collected,
        "external_total": external_total,
        "notes": notes,
        "taxable_programs": sorted(taxable),
        "nontaxable_programs": sorted(p for p in progs if p not in taxable),
    }


def report_to_grid(report, period_label, stmt=None):
    """Flatten the report into the iClassReport layout (one row per taxable program)."""
    grid = [["iClassReport", period_label, ""],
            ["Program", "Total Collected (Gross)", "Tax Collected (Portion)"]]
    for name, gross, tax in report["program_rows"]:
        grid.append([name, f"{gross:,.2f}", f"{tax:,.2f}"])
    grid.append(["", "", ""])
    grid.append(["Refunds", "", ""])
    grid.append(["Taxable", f"{report['ref_taxable']:,.2f}", ""])
    grid.append(["Non-taxable", f"{report['ref_nontaxable']:,.2f}", ""])
    grid.append(["", "", ""])
    grid.append(["Card Processing Fees", "", ""])
    for k, v in report["fees_section"].items():
        grid.append([k, f"{v:,.2f}", ""])
    grid.append(["", "", ""])
    grid.append(["Use Tax (Gross)", f"{report['use_tax']:,.2f}", ""])
    if stmt:
        grid.append(["", "", ""])
        grid.append(["Merchant Statement Summary", stmt.get("period", ""), ""])
        grid.append(["Total Sales", f"{stmt['total_sales']:,.2f}", ""])
        grid.append(["Total Refunds", f"({stmt['refunds']:,.2f})", ""])
        grid.append(["Total Processing Fees", f"({stmt['fees']:,.2f})", ""])
        grid.append(["Net Amount Settled", f"{stmt['net']:,.2f}", ""])
    return grid


def build_xlsx_bytes(report, period_label, stmt=None):
    """Build the iClassReport as a single-sheet Excel workbook with numeric
    cells and light formatting. Returns bytes."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    bold = Font(bold=True)
    money = "#,##0.00"

    wb = Workbook()

    # ---- Sheet 1: iClassReport ----
    ws = wb.active
    ws.title = "iClassReport"

    def row(vals, bold_cells=(), money_cells=()):
        ws.append(vals)
        r = ws.max_row
        for c in bold_cells:
            ws.cell(row=r, column=c).font = bold
        for c in money_cells:
            ws.cell(row=r, column=c).number_format = money

    row(["iClassReport", period_label, ""], bold_cells=(1,))
    row(["Program", "Total Collected (Gross)", "Tax Collected (Portion)"], bold_cells=(1, 2, 3))
    for name, gross, tax in report["program_rows"]:
        row([name, gross, tax], money_cells=(2, 3))
    row(["", "", ""])
    row(["Refunds", "", ""], bold_cells=(1,))
    row(["Taxable", report["ref_taxable"], ""], money_cells=(2,))
    row(["Non-taxable", report["ref_nontaxable"], ""], money_cells=(2,))
    row(["", "", ""])
    row(["Card Processing Fees", "", ""], bold_cells=(1,))
    for k, v in report["fees_section"].items():
        row([k, v, ""], money_cells=(2,))
    row(["", "", ""])
    row(["Use Tax (Gross)", report["use_tax"], ""], money_cells=(2,))
    if stmt:
        paren = "#,##0.00;(#,##0.00)"
        row(["", "", ""])
        row(["Merchant Statement Summary", stmt.get("period", ""), ""], bold_cells=(1,))
        row(["Total Sales", stmt["total_sales"], ""], money_cells=(2,))
        row(["Total Refunds", -stmt["refunds"], ""])
        ws.cell(row=ws.max_row, column=2).number_format = paren
        row(["Total Processing Fees", -stmt["fees"], ""])
        ws.cell(row=ws.max_row, column=2).number_format = paren
        row(["Net Amount Settled", stmt["net"], ""], money_cells=(2,))

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =========================================================================== #
#  UI
# =========================================================================== #
st.title("iClassReport Processor")
st.caption(f"Flip Side Lowell - v{VERSION} - FIN-4 Program Deposit Split + FIN-24 Gateway Transactions to iClassReport")

# ---- STEP 1 ----------------------------------------------------------------
st.subheader("Step 1 - Export the report")
st.info("In iClassPro, run the **FIN-4 Program Deposit Split** using your saved **payout-date "
        "preset** for the calendar month you're closing, and export it (CSV or Excel). "
        "That's the only file needed - fees come from the merchant statement below.")

# ---- STEP 2 ----------------------------------------------------------------
st.subheader("Step 2 - Upload it")
file_a = st.file_uploader("FIN-4 Program Deposit Split (CSV or XLSX)", type=["csv", "xlsx"], key="a")

# ---- STEP 3 ----------------------------------------------------------------
st.subheader("Step 3 - Merchant statement")
stmt_pdf = st.file_uploader("Upload the iClassPro monthly merchant statement (PDF) - the Summary "
                            "box is read automatically", type=["pdf"], key="stmt")
stmt = None
if stmt_pdf is not None:
    stmt = parse_statement_pdf(stmt_pdf)
    if stmt:
        st.success(f"Statement read ({stmt['period'] or 'period not found'}):  "
                   f"Total Sales ${stmt['total_sales']:,.2f}  -  Refunds ${stmt['refunds']:,.2f}  -  "
                   f"Fees ${stmt['fees']:,.2f}  -  Net Settled ${stmt['net']:,.2f}")
    else:
        st.error("Couldn't read the Summary box from that PDF - enter the numbers manually below.")

with st.expander("Or enter the Summary box numbers manually",
                 expanded=(stmt_pdf is not None and stmt is None)):
    st.caption("Only used when no statement PDF is uploaded (or it can't be read). Enter refunds "
               "and fees as positive numbers - the parentheses on the statement just mean "
               "they're subtractions.")
    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1:
        man_total_sales = st.number_input("Total Sales $", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    with sc2:
        man_refunds = st.number_input("Total Refunds $", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    with sc3:
        man_fees = st.number_input("Total Processing Fees $", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    with sc4:
        man_net = st.number_input("Net Amount Settled $", min_value=0.0, value=0.0, step=0.01, format="%.2f")

if stmt is None and (man_total_sales > 0 or man_refunds > 0 or man_fees > 0 or man_net > 0):
    stmt = {"total_sales": man_total_sales, "refunds": man_refunds,
            "fees": man_fees, "net": man_net, "period": ""}

stmt_total_sales = stmt["total_sales"] if stmt else 0.0
stmt_refunds = stmt["refunds"] if stmt else 0.0
stmt_fees = stmt["fees"] if stmt else 0.0
stmt_net = stmt["net"] if stmt else 0.0

use_tax = st.number_input("Use Tax (Gross) $ (optional)", min_value=0.0, value=0.0, step=0.01, format="%.2f")

if file_a is not None:
    rows = load_upload(file_a)
    if rows is None:
        st.stop()
    kind = detect_report_type(rows)
    if kind == "gateway":
        st.error("That's a Gateway Transactions (FIN-24) export - it's no longer used. "
                 "Upload the FIN-4 Program Deposit Split (payout-date preset) instead.")
    elif kind != "program_split":
        st.error("Could not identify this file as a FIN-4 Program Deposit Split export.")
    elif stmt_fees <= 0:
        st.warning("Enter the statement's **Total Processing Fees** in Step 3 - the fees section "
                   "and the 6040 JE line can't be built without it.")
    else:
        progs, unapplied_gw_net, unapplied_gw_refunds, period = parse_program_split(rows)
        fin4_total = round(sum(v["total"] for v in progs.values()) + unapplied_gw_net, 2)
        gateway = {
            "gross": fin4_total,
            "fees": round(stmt_fees, 2),
            "net": round(fin4_total - stmt_fees, 2),
            "count": 0, "settled_net": round(fin4_total - stmt_fees, 2),
            "unsettled_amt": 0.0, "unsettled_count": 0,
            "period": period, "date_min": None, "date_max": None,
        }
        report = build_iclassreport(progs, unapplied_gw_net, gateway, use_tax=use_tax)
        period_label = period or datetime.date.today().strftime("%m/%d/%Y")
        report_refunds = round(report["ref_taxable"] + report["ref_nontaxable"]
                               - unapplied_gw_refunds, 2)

        st.success(f"Processed period: {period_label} (payout basis)  -  {len(progs)} programs")

        if stmt and stmt.get("period") and period and stmt["period"][:10][:2] != period[:2]:
            st.warning(f"The statement covers {stmt['period']} but the FIN-4 export covers "
                       f"{period} - make sure both are for the same month.")

        # ---- sanity: payout-basis exports have no cash ----
        if abs(report["cash_collected"]) > 0.005:
            st.warning(f"This export shows ${report['cash_collected']:,.2f} of Cash/Check - a "
                       "payout-basis export should show $0.00 cash. It was probably run on "
                       "payments-received basis. Check the preset before sending anything out.")

        # ---- iClassReport table ----
        st.subheader("iClassReport")
        grid = report_to_grid(report, period_label, stmt=stmt)
        df_display = pd.DataFrame(grid[2:], columns=grid[1])
        st.dataframe(df_display, width="stretch", hide_index=True)

        # ---- statement validation ----
        st.subheader("Merchant statement validation")
        if stmt_total_sales > 0 or stmt_refunds > 0 or stmt_net > 0:
            checks = [
                ("Total Sales", round(fin4_total + report_refunds, 2), stmt_total_sales),
                ("Total Refunds", report_refunds, stmt_refunds),
                ("Net Amount Settled", gateway["net"], stmt_net),
            ]
            all_ok = True
            for label, got, want in checks:
                if want <= 0:
                    continue
                var = round(got - want, 2)
                if abs(var) < 0.005:
                    st.success(f"{label}: report ${got:,.2f} = statement ${want:,.2f}  (exact match)")
                else:
                    all_ok = False
                    st.error(f"{label}: report ${got:,.2f} vs statement ${want:,.2f}  ->  variance "
                             f"${var:,.2f}. This should be $0.00. Check that the export used the "
                             "payout-date preset for the right month, and that the Step 3 numbers "
                             "were copied from the statement's Summary box.")
            if all_ok:
                st.success("All statement totals match exactly. Report is good to send and the JE "
                           "is good to post.")
            st.caption("Note: the statement's separate 'Bank Transfer Fees' line (e.g. $0.54) is an "
                       "account fee, not a processing fee - it is excluded from this check and does "
                       "not reduce the payouts.")
        else:
            st.info("Enter the statement's Summary box numbers in Step 3 to validate. These should "
                    "match to the penny.")

        # ---- draft JE ----
        st.subheader("Draft 2090 clearing JE")
        st.caption("Post in Xero dated the last day of the month.")
        je_rows = [(k, f"{v:,.2f}", "") for k, v in report["je_debits"]] + \
                  [(k, "", f"{v:,.2f}") for k, v in report["je_credits"]] + \
                  [("TOTALS", f"{report['je_debit_total']:,.2f}", f"{report['je_credit_total']:,.2f}")]
        st.table(pd.DataFrame(je_rows, columns=["Account", "Debit", "Credit"]))
        if report["je_balance_ok"]:
            st.success("JE balances: debits = credits.")
        else:
            st.error(f"JE DOES NOT BALANCE (DR ${report['je_debit_total']:,.2f} vs "
                     f"CR ${report['je_credit_total']:,.2f}). Do not post - check the export file "
                     "and the Step 3 fee entry.")

        # ---- notes ----
        if report["notes"]:
            st.subheader("Notes / exceptions")
            for note in report["notes"]:
                st.write("- " + note)
        st.write("- Payout basis: cash sales never appear on this report (cash has no payout). "
                 "Capture the cash split when you make the monthly cash deposit.")

        # ---- export: copy/paste + download ----
        st.subheader("Copy & paste")
        st.caption("Hover the box and click the copy icon (top-right), then paste into Google Sheets "
                   "or Excel - it splits into columns automatically.")
        tsv = "\n".join("\t".join(str(c) for c in row) for row in grid)
        st.code(tsv, language=None)

        st.caption("Draft 2090 clearing JE:")
        je_tsv = "\n".join([f"{k}\t{v:,.2f}\t" for k, v in report["je_debits"]] +
                           [f"{k}\t\t{v:,.2f}" for k, v in report["je_credits"]])
        st.code(je_tsv, language=None)

        try:
            xlsx_bytes = build_xlsx_bytes(report, period_label, stmt=stmt)
            st.download_button(
                "Download iClassReport (Excel)", xlsx_bytes,
                file_name=f"iClassReport_{period_label.replace('/', '-')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except ImportError:
            st.error("Excel export needs `openpyxl` - add a line saying `openpyxl` to "
                     "requirements.txt in GitHub. Falling back to CSV:")
            csv_buf = io.StringIO()
            csv.writer(csv_buf).writerows(grid)
            st.download_button("Download iClassReport (CSV)", csv_buf.getvalue(),
                               file_name=f"iClassReport_{period_label.replace('/', '-')}.csv",
                               mime="text/csv")
else:
    st.info("Upload the payout-basis FIN-4 export to begin.")
