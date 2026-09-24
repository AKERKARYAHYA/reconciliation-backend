# -*- coding: utf-8 -*-
"""
BMCI Bank Reconciliation — Streamlit application
================================================
Two-phase ERP ⇄ RLV (bank statement) reconciliation with strict chronological
and weekday constraints.

Rules
-----
* Process ONLY 'Credit' transactions (debit rows are ignored entirely).
* Chronological constraint : ERP date <= RLV date.
* Strict date window (ERP → RLV):
    - ERP on Saturday  → match ONLY with the following Monday.
    - any other day    → match with the SAME day or DAY + 1.
* A row matched in one phase is excluded from every later phase (no double count).

Phase 1 — Direct reference search
    ERP `reference` found literally inside RLV `description`,
    identical Credit amount, within the date window. (1 ERP ⇄ 1 RLV)

Phase 2 — Aggregation on "VERSEMENT ESP"
    RLV restricted to descriptions containing "VERSEMENT ESP".
    ERP grouped by reference (a match may involve 1 or 2 references on the same date).
    A combination of 2 to 8 RLV rows (same date) whose Credit sum equals the
    ERP Credit sum, within the date window.

Phase 3 — Same-reference ERP aggregation → single RLV row
    Among the rows still unmatched, an ERP reference carrying several Credit lines
    on the same date is summed and matched to ONE RLV row whose Credit equals that
    total (any description), within the same date window.

Run
---
    pip install streamlit pandas openpyxl
    streamlit run bmci_reconciliation_app.py
"""

import io
import re
import unicodedata
import itertools
import datetime
from collections import defaultdict

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ============================================================================
#  Helpers
# ============================================================================
def deaccent(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


def find_col(columns, *keywords):
    """First column whose (accent-insensitive) header contains any keyword."""
    for col in columns:
        h = deaccent(col)
        if any(kw in h for kw in keywords):
            return col
    return None


def to_num(series: pd.Series) -> pd.Series:
    """Parse amounts, tolerating spaces and comma decimals (e.g. '1 234,56')."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    cleaned = (series.astype(str)
               .str.replace(" ", "", regex=False)
               .str.replace(" ", "", regex=False)
               .str.replace(",", ".", regex=False)
               .str.replace(r"[^0-9.\-]", "", regex=True))
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def parse_date_val(v):
    """Robust per-value date parser: real dates, dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd."""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(v).normalize()
    s = str(v).strip()
    if not s:
        return pd.NaT
    parts = re.split(r"[-/]", s.split()[0])
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        a, b, c = (int(p) for p in parts)
        try:
            return pd.Timestamp(year=a, month=b, day=c) if a > 31 else pd.Timestamp(year=c, month=b, day=a)
        except ValueError:
            return pd.NaT
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def read_upload(upload):
    """Read an uploaded Excel/CSV file into a DataFrame (first sheet for Excel)."""
    name = upload.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        sheets = pd.read_excel(upload, sheet_name=None)
        # prefer a sheet named like erp/rlv, else the first
        for s in sheets:
            if deaccent(s) in ("erp", "rlv"):
                return sheets[s]
        return list(sheets.values())[0]
    return pd.read_csv(upload, sep=None, engine="python")


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize a raw ERP/RLV frame to: date, credit, reference, desc, cents.
    Keeps Credit transactions only."""
    cols = list(df.columns)
    date_c = find_col(cols, "date")
    cred_c = find_col(cols, "cred")
    ref_c = find_col(cols, "ref")

    out = pd.DataFrame()
    out["date"] = df[date_c].map(parse_date_val) if date_c else pd.NaT
    out["credit"] = to_num(df[cred_c]) if cred_c else 0.0
    out["reference"] = (df[ref_c].astype(str).str.strip().replace({"nan": "", "None": ""})
                        if ref_c else "")
    # Searchable description = every text column except the date column.
    text_cols = [c for c in cols if c != date_c
                 and not pd.api.types.is_numeric_dtype(df[c])
                 and not pd.api.types.is_datetime64_any_dtype(df[c])]
    out["desc"] = df[text_cols].astype(str).agg(" | ".join, axis=1) if text_cols else ""
    out["desc_lower"] = out["desc"].map(deaccent)
    out["cents"] = (out["credit"] * 100).round().astype("int64")
    out = out[out["credit"] > 0].reset_index(drop=True)
    return out


def date_ok(ed, rd) -> bool:
    """ERP→RLV strict date window (already enforces ERP date <= RLV date)."""
    if pd.isna(ed) or pd.isna(rd):
        return False
    if ed.weekday() == 5:                          # Saturday → following Monday only
        return rd == ed + pd.Timedelta(days=2)
    return rd == ed or rd == ed + pd.Timedelta(days=1)   # same day or +1


def find_combo(items, target, kmin=2, kmax=8, guard=2_000_000):
    """Subset of 2..8 rows whose cents sum == target. Returns indices or None."""
    cand = sorted([(i, c) for i, c in items if 0 < c <= target], key=lambda x: -x[1])
    n = len(cand)
    if n < kmin:
        return None
    cents = [c for _, c in cand]
    count = 0
    for k in range(kmin, min(kmax, n) + 1):
        if sum(cents[:k]) < target:        # k largest can't reach target → try larger k
            continue
        if sum(cents[n - k:]) > target:    # k smallest already exceed → stop
            break
        for combo in itertools.combinations(range(n), k):
            count += 1
            if count > guard:              # safety valve against combinatorial blow-up
                return None
            s = 0
            for j in combo:
                s += cents[j]
                if s > target:
                    break
            if s == target:
                return [cand[j][0] for j in combo]
    return None


# ============================================================================
#  Reconciliation engine
# ============================================================================
def reconcile(erp: pd.DataFrame, rlv: pd.DataFrame):
    erp_matched, rlv_matched, matches = set(), set(), []
    mid = 0

    # ---------- Phase 1 : direct reference search ----------
    rlv_by_amt = defaultdict(list)
    for ri in rlv.index:
        rlv_by_amt[rlv.at[ri, "cents"]].append(ri)

    for ei in erp.index:
        ref = erp.at[ei, "reference"]
        ed = erp.at[ei, "date"]
        if not ref or pd.isna(ed):
            continue
        ref_l = deaccent(ref)
        cands = [ri for ri in rlv_by_amt.get(erp.at[ei, "cents"], [])
                 if ri not in rlv_matched
                 and date_ok(ed, rlv.at[ri, "date"])
                 and ref_l in rlv.at[ri, "desc_lower"]]
        if cands:
            ri = min(cands, key=lambda x: (rlv.at[x, "date"], x))
            erp_matched.add(ei)
            rlv_matched.add(ri)
            mid += 1
            matches.append({"mid": mid, "type": "Phase 1", "erp": [ei], "rlv": [ri]})

    # ---------- Phase 2 : aggregation on "VERSEMENT ESP" ----------
    groups = defaultdict(lambda: {"cents": 0, "idx": []})       # (date, reference) → sum
    for i in erp.index:
        if i in erp_matched or pd.isna(erp.at[i, "date"]):
            continue
        key = (erp.at[i, "date"], erp.at[i, "reference"])
        groups[key]["cents"] += erp.at[i, "cents"]
        groups[key]["idx"].append(i)

    esp = [ri for ri in rlv.index
           if ri not in rlv_matched and "versement esp" in rlv.at[ri, "desc_lower"]
           and not pd.isna(rlv.at[ri, "date"])]
    rlv_by_date = defaultdict(list)
    for ri in esp:
        rlv_by_date[rlv.at[ri, "date"]].append(ri)

    def allowed_dates(ed):
        return [ed + pd.Timedelta(days=2)] if ed.weekday() == 5 else [ed, ed + pd.Timedelta(days=1)]

    def try_match(target, ed):
        for rd in allowed_dates(ed):
            rows = [ri for ri in rlv_by_date.get(rd, []) if ri not in rlv_matched]
            if len(rows) >= 2:
                combo = find_combo([(ri, rlv.at[ri, "cents"]) for ri in rows], target)
                if combo:
                    return combo
        return None

    # singles (1 reference)
    for key in sorted(groups, key=lambda k: (k[0], str(k[1]))):
        date, _ = key
        g = groups[key]
        if any(i in erp_matched for i in g["idx"]):
            continue
        combo = try_match(g["cents"], date)
        if combo:
            for i in g["idx"]:
                erp_matched.add(i)
            for ri in combo:
                rlv_matched.add(ri)
            mid += 1
            matches.append({"mid": mid, "type": "Phase 2", "erp": list(g["idx"]), "rlv": combo})

    # pairs (2 references sharing the same date)
    by_date = defaultdict(list)
    for key in groups:
        by_date[key[0]].append(key)
    for date, keys in sorted(by_date.items()):
        avail = [k for k in keys if not any(i in erp_matched for i in groups[k]["idx"])]
        for k1, k2 in itertools.combinations(sorted(avail, key=lambda k: str(k[1])), 2):
            idxs = groups[k1]["idx"] + groups[k2]["idx"]
            if any(i in erp_matched for i in idxs):
                continue
            combo = try_match(groups[k1]["cents"] + groups[k2]["cents"], date)
            if combo:
                for i in idxs:
                    erp_matched.add(i)
                for ri in combo:
                    rlv_matched.add(ri)
                mid += 1
                matches.append({"mid": mid, "type": "Phase 2", "erp": list(idxs), "rlv": combo})

    # ---------- Phase 3 : same-reference ERP aggregation → single RLV row ----------
    # A reference carrying several Credit lines on one date: sum them and look for
    # ONE remaining RLV row whose Credit equals that total (any description),
    # still within the strict date window.
    groups3 = defaultdict(lambda: {"cents": 0, "idx": []})
    for i in erp.index:
        if i in erp_matched or pd.isna(erp.at[i, "date"]) or not erp.at[i, "reference"]:
            continue
        key = (erp.at[i, "date"], erp.at[i, "reference"])
        groups3[key]["cents"] += erp.at[i, "cents"]
        groups3[key]["idx"].append(i)

    for key in sorted(groups3, key=lambda k: (k[0], str(k[1]))):
        date, _ = key
        g = groups3[key]
        if len(g["idx"]) < 2 or any(i in erp_matched for i in g["idx"]):
            continue
        cands = [ri for ri in rlv_by_amt.get(g["cents"], [])
                 if ri not in rlv_matched and date_ok(date, rlv.at[ri, "date"])]
        if cands:
            ri = min(cands, key=lambda x: (rlv.at[x, "date"], x))
            for i in g["idx"]:
                erp_matched.add(i)
            rlv_matched.add(ri)
            mid += 1
            matches.append({"mid": mid, "type": "Phase 3", "erp": list(g["idx"]), "rlv": [ri]})

    return matches, erp_matched, rlv_matched


# ============================================================================
#  Output workbook
# ============================================================================
def build_outputs(erp, rlv, matches, erp_matched, rlv_matched):
    """Return (excel_bytes, matched_df, unmatched_df)."""
    mrows = []
    for m in matches:
        e_idx, r_idx = m["erp"], m["rlv"]
        e_tot = round(sum(erp.at[i, "credit"] for i in e_idx), 2)
        r_tot = round(sum(rlv.at[i, "credit"] for i in r_idx), 2)
        for k in range(max(len(e_idx), len(r_idx))):
            row = {"Match ID": m["mid"], "Match Type": m["type"]}
            if k < len(e_idx):
                i = e_idx[k]
                row.update({"ERP Date": erp.at[i, "date"], "ERP Référence": erp.at[i, "reference"],
                            "ERP Crédit": erp.at[i, "credit"]})
            else:
                row.update({"ERP Date": None, "ERP Référence": "", "ERP Crédit": None})
            if k < len(r_idx):
                j = r_idx[k]
                row.update({"RLV Date": rlv.at[j, "date"], "RLV Description": rlv.at[j, "desc"],
                            "RLV Crédit": rlv.at[j, "credit"]})
            else:
                row.update({"RLV Date": None, "RLV Description": "", "RLV Crédit": None})
            row["ERP Total"] = e_tot if k == 0 else None
            row["RLV Total"] = r_tot if k == 0 else None
            mrows.append(row)
    matched_df = pd.DataFrame(mrows, columns=["Match ID", "Match Type", "ERP Date", "ERP Référence",
                                              "ERP Crédit", "RLV Date", "RLV Description", "RLV Crédit",
                                              "ERP Total", "RLV Total"])

    urows = [{"Source": "ERP", "Date": erp.at[i, "date"], "Référence": erp.at[i, "reference"],
              "Description": erp.at[i, "desc"], "Crédit": erp.at[i, "credit"]}
             for i in erp.index if i not in erp_matched]
    urows += [{"Source": "RLV", "Date": rlv.at[j, "date"], "Référence": rlv.at[j, "reference"],
               "Description": rlv.at[j, "desc"], "Crédit": rlv.at[j, "credit"]}
              for j in rlv.index if j not in rlv_matched]
    unmatched_df = pd.DataFrame(urows, columns=["Source", "Date", "Référence", "Description", "Crédit"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="dd/mm/yyyy") as xl:
        matched_df.to_excel(xl, sheet_name="Matched_Results", index=False)
        unmatched_df.to_excel(xl, sheet_name="Unmatched_Data", index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    navy, white = "1F3864", "FFFFFF"
    for ws in wb.worksheets:
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(1, c)
            cell.font = Font(name="Arial", size=10, bold=True, color=white)
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[cell.column_letter].width = 22
        for col in ws.iter_cols(min_row=2):
            for cell in col:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
        ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), matched_df, unmatched_df


# ============================================================================
#  Streamlit UI
# ============================================================================
st.set_page_config(page_title="Réconciliation BMCI", page_icon="🏦", layout="wide")
st.title("🏦 Réconciliation BMCI — ERP ⇄ Relevé bancaire")
st.caption("Rapprochement en deux phases (référence directe + agrégation « VERSEMENT ESP ») "
           "avec contrainte chronologique et règle Samedi → Lundi. Crédits uniquement.")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1️⃣ Fichier ERP")
    erp_up = st.file_uploader("ERP (Excel ou CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="erp")
with c2:
    st.subheader("2️⃣ Fichier RLV (banque)")
    rlv_up = st.file_uploader("RLV (Excel ou CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="rlv")

if erp_up and rlv_up:
    try:
        with st.spinner("Rapprochement en cours…"):
            erp = prepare(read_upload(erp_up))
            rlv = prepare(read_upload(rlv_up))
            matches, em, rm = reconcile(erp, rlv)
            xlsx_bytes, matched_df, unmatched_df = build_outputs(erp, rlv, matches, em, rm)
    except Exception as e:
        st.error(f"Impossible de traiter les fichiers : {e}")
        st.stop()

    p1 = sum(1 for m in matches if m["type"] == "Phase 1")
    p2 = sum(1 for m in matches if m["type"] == "Phase 2")
    p3 = sum(1 for m in matches if m["type"] == "Phase 3")
    esp_n = int(rlv["desc_lower"].str.contains("versement esp").sum())

    k = st.columns(6)
    k[0].metric("Crédits ERP", len(erp))
    k[1].metric("Crédits RLV", len(rlv))
    k[2].metric("Matchs Phase 1", p1)
    k[3].metric("Matchs Phase 2", p2)
    k[4].metric("Matchs Phase 3", p3)
    k[5].metric("Non rapprochés", f"{len(erp) - len(em)} / {len(rlv) - len(rm)}",
                help="ERP non rapprochés / RLV non rapprochés")

    st.caption(f"Lignes RLV contenant « VERSEMENT ESP » : {esp_n}")

    st.subheader("✅ Matched_Results")
    st.dataframe(matched_df, use_container_width=True, hide_index=True)
    st.subheader("⚠️ Unmatched_Data")
    st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Télécharger le rapprochement (Excel — 2 feuilles)",
        data=xlsx_bytes,
        file_name="BMCI_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("⬆️ Téléversez les fichiers ERP et RLV pour lancer le rapprochement.")
