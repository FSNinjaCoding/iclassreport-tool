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
    raw_text = raw_text.lstrip("\ufeff")
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
    birthdays = [p for p in progs if
