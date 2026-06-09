"""
iClassReport Processor  -  Flip Side Lowell, LLC
================================================
Upload two iClassPro exports and get the canonical iClassReport ST Ledgerwood
requires, a draft 2090 clearing JE, and a bank cross-check.

Inputs (CSV exports, Payments-Received / month-to-date basis):
  A) FIN-4  Program Deposit Split Report
  B) FIN-24 Gateway Transactions Report

Output is a one-click copy block (pastes straight into Sheets/Excel) plus a CSV
download - no secrets or accounts required, so anyone on the team can run it.

Logic validated against April 2026 (matches the accepted report to the penny)
and May 2026.  Version 1.1
"""
import csv
import io
import re
import datetime

import pandas as pd
import streamlit as st

# --- CONFIGURATION ---------------------------------------------------------
VERSION = "1.1"

# Open Gym and Birthdays lead the report; every other taxable program follows
# (alphabetical), one row each. Taxable vs non-taxable is detected from the
# report (tax collected > 0), so new programs classify themselves.
OPEN_GYM_NAMES = {"open gym"}
BIRTHDAY_NAMES = {"birthdays", "birthday"}

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
    raw_text = raw_text.lstrip("﻿")
    return list(csv.reader(io.StringIO(raw_text)))


def detect_report_type(raw_text):
    """Return 'gateway', 'program_split', or None by inspecting the first rows."""
    rows = _read_rows(raw_text)
    flat = " ".join(" ".join(r) for r in rows[:3]).lower()
    if "final settlement amount" in flat or "transaction id" in flat:
        return "gateway"
    if "charge category" in flat or "payments received" in flat:
        return "program_split"
    return None


def parse_program_split(raw_text):
    """FIN-4 -> (dict program -> {total, tax, gross, refunds[]}, period_str)."""
    rows = _read_rows(raw_text)
    period = ""
    if rows and rows[0]:
        for cell in rows[0]:
            m = re.search(r"\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4}", cell)
            if m:
                period = m.group(0)
                break

    hidx = None
    for i, r in enumerate(rows[:5]):
        low = [c.strip().lower() for c in r]
        if "program" in low and "charge category" in low:
            hidx = i
            break
    if hidx is None:
        hidx = 1
    header = [c.strip() for c in rows[hidx]]

    def col(name, default):
        for i, c in enumerate(header):
            if c.strip().lower() == name.lower():
                return i
        return default

    i_prog = col("Program", 1)
    i_cat = col("Charge Category", 2)
    i_total = col("Total", 3)
    i_tax = col("Taxes", 4)

    progs = {}
    cur = None
    for r in rows[hidx + 1:]:
        if len(r) <= i_cat:
            continue
        name = r[i_prog].strip()
        cat = r[i_cat].strip()
        if name.lower().startswith("total payments received"):
            break
        if cat == "Program Total:":
            cur = name
            progs[cur] = {
                "total": parse_money(r[i_total]),
                "tax": parse_money(r[i_tax]),
                "gross": 0.0,
                "refunds": [],
            }
        elif cur and name == "" and cat:
            amt = parse_money(r[i_total])
            if "refund" in cat.lower():
                progs[cur]["refunds"].append(amt)
            else:
                progs[cur]["gross"] += amt
    return progs, period


def parse_gateway(raw_text):
    """FIN-24 -> dict with gross/fees/net + settled/unsettled splits + period."""
    rows = _read_rows(raw_text)
    hidx = None
    for i, r in enumerate(rows[:5]):
        low = [c.strip().lower() for c in r]
        if "amount" in low and "final settlement amount" in low:
            hidx = i
            break
    if hidx is None:
        hidx = 0
    header = [c.strip() for c in rows[hidx]]

    def col(name):
        for i, c in enumerate(header):
            if c.strip().lower() == name.lower():
                return i
        return None

    i_date = col("Date")
    i_amt = col("Amount")
    i_fee = col("Fees")
    i_net = col("Final Settlement Amount")
    i_status = col("Status")

    gross = fees = net = settled_net = unsettled_amt = 0.0
    n = unsettled_n = 0
    dates = []
    for r in rows[hidx + 1:]:
        if i_amt is None or len(r) <= i_amt:
            continue
        if i_date is not None and (len(r) <= i_date or "/" not in str(r[i_date])):
            continue  # skip totals row
        a = parse_money(r[i_amt])
        f = parse_money(r[i_fee]) if i_fee is not None else 0.0
        s = parse_money(r[i_net]) if i_net is not None else (a - f)
        status = (r[i_status].strip().lower() if i_status is not None and len(r) > i_status else "")
        gross += a
        fees += f
        net += s
        n += 1
        if i_date is not None:
            dates.append(str(r[i_date]))
        if status == "settled":
            settled_net += s
        else:
            unsettled_amt += a
            unsettled_n += 1
    period = ""
    if dates:
        try:
            ds = [datetime.datetime.strptime(d, "%m/%d/%Y") for d in dates]
            period = f"{min(ds):%m/%d/%Y} - {max(ds):%m/%d/%Y}"
        except Exception:
            period = ""
    return {
        "gross": round(gross, 2),
        "fees": round(fees, 2),
        "net": round(net, 2),
        "count": n,
        "settled_net": round(settled_net, 2),
        "unsettled_amt": round(unsettled_amt, 2),
        "unsettled_count": unsettled_n,
        "period": period,
    }


def build_iclassreport(progs, gateway, use_tax=0.0):
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

    # Draft 2090 clearing JE (net of refunds, ex-tax)
    retail_progs = [p for p in taxable if p not in open_gym and p not in birthdays]
    venue = round(sum(progs[p]["total"] - progs[p]["tax"] for p in open_gym + birthdays), 2)
    merch = round(sum(progs[p]["total"] - progs[p]["tax"] for p in retail_progs), 2)
    member = round(sum(progs[p]["total"] for p in progs if p not in taxable), 2)

    notes = []
    extra_taxable = [p for p in taxable if p not in open_gym and p not in birthdays]
    if extra_taxable:
        notes.append("Taxable programs beyond Open Gym / Birthdays: " + ", ".join(sorted(extra_taxable)))
    if gateway.get("unsettled_count"):
        notes.append(f"{gateway['unsettled_count']} gateway transaction(s) NOT settled "
                     f"(${gateway['unsettled_amt']:,.2f}) - processed but not completed.")

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
        "je": {
            "4200 Sales:Venue (Open Gym + Birthdays)": venue,
            "4000 Sales:Merchandise (other taxable)": merch,
            "4100 Sales:Member Fees (non-taxable)": member,
            "2230 Sales Tax Payable": total_tax,
            "6040 Credit Card Merchant Fees (DR)": gateway["fees"],
        },
        "notes": notes,
        "taxable_programs": sorted(taxable),
        "nontaxable_programs": sorted(p for p in progs if p not in taxable),
    }


def report_to_grid(report, period_label):
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
    return grid


# =========================================================================== #
#  UI
# =========================================================================== #
st.title("iClassReport Processor")
st.caption(f"Flip Side Lowell - v{VERSION} - FIN-4 Program Deposit Split + FIN-24 Gateway Transactions to iClassReport")

c1, c2 = st.columns(2)
with c1:
    file_a = st.file_uploader("Upload report A (CSV)", type=["csv"], key="a")
with c2:
    file_b = st.file_uploader("Upload report B (CSV)", type=["csv"], key="b")

with st.expander("Optional inputs"):
    bank_deposits = st.number_input(
        "iClassPro card payouts that hit the bank / 2090 this month ($) - for the cross-check",
        min_value=0.0, value=0.0, step=0.01, format="%.2f")
    use_tax = st.number_input("Use Tax (Gross) $", min_value=0.0, value=0.0, step=0.01, format="%.2f")

if file_a and file_b:
    texts = {}
    for f in (file_a, file_b):
        f.seek(0)
        texts[f.name] = f.read().decode("utf-8-sig", errors="ignore")

    program_text = gateway_text = None
    for name, txt in texts.items():
        kind = detect_report_type(txt)
        if kind == "program_split":
            program_text = txt
        elif kind == "gateway":
            gateway_text = txt

    if not program_text or not gateway_text:
        st.error("Could not identify both files. Need one Program Deposit Split (FIN-4) "
                 "and one Gateway Transactions (FIN-24) export.")
    else:
        progs, period = parse_program_split(program_text)
        gateway = parse_gateway(gateway_text)
        report = build_iclassreport(progs, gateway, use_tax=use_tax)
        period_label = period or gateway["period"] or datetime.date.today().strftime("%m/%d/%Y")

        st.success(f"Processed period: {period_label}  -  {gateway['count']} gateway transactions  -  "
                   f"{len(progs)} programs")

        # ---- iClassReport table ----
        st.subheader("iClassReport")
        grid = report_to_grid(report, period_label)
        df_display = pd.DataFrame(grid[2:], columns=grid[1])
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        # ---- cross-check ----
        st.subheader("Bank cross-check")
        net = report["fees_section"]["Total Net Deposited"]
        if bank_deposits > 0:
            var = round(net - bank_deposits, 2)
            msg = (f"Net Deposited **${net:,.2f}** vs. bank/2090 card payouts **${bank_deposits:,.2f}** "
                   f"-> variance **${var:,.2f}**")
            (st.success if abs(var) < 50 else st.warning)(
                msg + "  \nA small variance here is expected end-of-month timing "
                "(late-month transactions that settle next month).")
        else:
            st.info(f"Net Deposited (per gateway): **${net:,.2f}**. Enter your bank/2090 card-payout "
                    "total in *Optional inputs* to see the timing variance.")

        # ---- draft JE ----
        st.subheader("Draft 2090 clearing JE")
        st.caption("Report-as-run basis. Revenue allocations post to the accounts below; "
                   "any small month-end in-transit clears next month.")
        je_df = pd.DataFrame([(k, f"{v:,.2f}") for k, v in report["je"].items()],
                             columns=["Account", "Amount"])
        st.table(je_df)

        # ---- notes ----
        if report["notes"]:
            st.subheader("Notes / exceptions")
            for note in report["notes"]:
                st.write("- " + note)

        # ---- export: copy/paste + download ----
        st.subheader("Copy & paste")
        st.caption("Hover the box and click the copy icon (top-right), then paste into Google Sheets "
                   "or Excel - it splits into columns automatically.")
        tsv = "\n".join("\t".join(str(c) for c in row) for row in grid)
        st.code(tsv, language=None)

        st.caption("Draft 2090 clearing JE:")
        je_tsv = "\n".join(f"{k}\t{v:,.2f}" for k, v in report["je"].items())
        st.code(je_tsv, language=None)

        csv_buf = io.StringIO()
        csv.writer(csv_buf).writerows(grid)
        st.download_button("Download iClassReport (CSV)", csv_buf.getvalue(),
                           file_name=f"iClassReport_{period_label.replace('/', '-')}.csv",
                           mime="text/csv")
else:
    st.info("Upload both CSV exports to begin. Order doesn't matter - the app detects which is which.")
