import os
import re
import sys
import math
from datetime import datetime, timedelta
import pandas as pd
import xlrd  # ==1.2.0 for formatting_info=True
from xlutils.copy import copy as xl_copy
from xlwt import XFStyle, Font, Alignment, Borders, Pattern
from tkinter import Tk, filedialog

# ----------------------------
# Configuration
# ----------------------------
INVENTORY_DIR = os.path.expanduser("~/Desktop/Inventory")

QVC_FILE   = "QVC Inventory.xls"
JCP_FILE   = "JCP Inventory.xls"
MACYS_FILE = "Macy's Inventory.xls"

QVC_SHEET   = "Sheet1"
JCP_SHEET   = "Sheet1"
MACYS_SHEET = "Sheet 1"   # Macy's fixed name

DATA_START_ROW = 6  # first data row (1-based)

# Column indices (0-based) for retailer files
COL_B  = 1   # Vendor SKU
COL_D  = 3   # Quantity
COL_Y  = 24  # Macy's UPC
COL_AA = 26  # Second quantity column (QVC/Macy's)

# Master sheet config
MASTER_SHEET = "Export"
MASTER_COL_A_SKU     = "Entry Name"
MASTER_COL_B_QTY     = "Quantity"
MASTER_COL_C_UPC     = "UPC"
MASTER_COL_D_JCP_SKU = "JCP SKU NAME"

# Merchant Reference Data (MRD)
MRD_BASENAME_NO_EXT = "Merchant Reference Data"  # tries .xlsx then .xls
MRD_COL_VENDOR_SKU = "Vendor SKU"
MRD_COL_MERCHANT   = "Merchant"     # "Macy's" or "QVC Drop Ship"
MRD_COL_UPC        = "UPC"

# Filenames
HISTORY_PATH = os.path.join(INVENTORY_DIR, "Inventory History.xlsx")  # persistent
# Report name is generated daily as "Inventory Report MM.DD.YY.xlsx"

# ----------------------------
# Normalization & matching
# ----------------------------
def normalize_basic(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isfinite(v) and v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

def normalize_casefold(s): return normalize_basic(s).casefold()
def normalize_nospace(s): return normalize_basic(s).replace(" ", "")
def normalize_casefold_nospace(s): return normalize_casefold(s).replace(" ", "")

def normalize_upc(v):
    s = normalize_basic(v)
    s = s.replace("E+", "e+")
    if "e+" in s.lower():
        try:
            f = float(s)
            s = str(int(f)) if f.is_integer() else f"{f:.0f}"
        except Exception:
            pass
    return re.sub(r"\D+", "", s)

def quantity_int(x):
    if pd.isna(x): return 0
    try: return int(math.floor(float(x)))
    except Exception:
        digits = re.sub(r"[^0-9\-]+", "", str(x))
        return int(digits) if digits else 0

def build_map(keys, vals, func):
    d = {}
    for k, v in zip(keys, vals):
        key = func(k)
        if key:
            d[key] = quantity_int(v)
    return d

def make_lookup_dicts(df_master):
    df = df_master.copy()
    df[MASTER_COL_B_QTY] = df[MASTER_COL_B_QTY].apply(quantity_int)

    sku = df[MASTER_COL_A_SKU].astype(str).fillna("")
    qty = df[MASTER_COL_B_QTY]
    upc = df[MASTER_COL_C_UPC].astype(str).fillna("")
    jcp = df[MASTER_COL_D_JCP_SKU].astype(str).fillna("")

    return {
        "sku": [build_map(sku, qty, normalize_basic),
                build_map(sku, qty, normalize_casefold),
                build_map(sku, qty, normalize_nospace),
                build_map(sku, qty, normalize_casefold_nospace)],
        "jcp": [build_map(jcp, qty, normalize_basic),
                build_map(jcp, qty, normalize_casefold),
                build_map(jcp, qty, normalize_nospace),
                build_map(jcp, qty, normalize_casefold_nospace)],
        "upc": [build_map(upc, qty, normalize_upc)]
    }

# --- matchers that also tell us where the match came from ---
def match_master_sku(sku, master_maps):
    raw = normalize_basic(sku)
    if not raw: return (False, 0, None)
    for keyfunc, mp in zip(
        [normalize_basic, normalize_casefold, normalize_nospace, normalize_casefold_nospace],
        master_maps["sku"]
    ):
        k = keyfunc(raw)
        if k in mp:
            return True, mp[k], "Master SKU"
    return (False, 0, None)

def match_master_upc(upc, master_maps):
    k = normalize_upc(upc)
    if not k: return (False, 0, None)
    for d in master_maps["upc"]:
        if k in d:
            return True, d[k], "Master UPC"
    return (False, 0, None)

def match_master_jcp_sku(sku, master_maps):
    raw = normalize_basic(sku)
    if not raw: return (False, 0, None)
    for keyfunc, mp in zip(
        [normalize_basic, normalize_casefold, normalize_nospace, normalize_casefold_nospace],
        master_maps["jcp"]
    ):
        k = keyfunc(raw)
        if k in mp:
            return True, mp[k], "Master JCP SKU"
    return (False, 0, None)

# ----------------------------
# XLRD utilities: last data row & open helpers
# ----------------------------
def _cell_nonempty(sheet, r, c):
    try:
        t = sheet.cell_type(r, c)
        v = sheet.cell_value(r, c)
    except Exception:
        return False
    if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return False
    if t == xlrd.XL_CELL_TEXT:
        return str(v).strip() != ""
    return True

def last_data_row(sheet, key_cols, start_row_1based):
    start0 = start_row_1based - 1
    n = sheet.nrows
    for r in range(n - 1, start0 - 1, -1):
        for c in key_cols:
            if _cell_nonempty(sheet, r, c):
                return r
    return start0 - 1

def open_xls(path: str, sheet_name: str):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File not found: {path}")
    rb = xlrd.open_workbook(path, formatting_info=True)
    try:
        idx = rb.sheet_names().index(sheet_name)
    except ValueError:
        raise ValueError(f"Sheet '{sheet_name}' not found in '{path}'")
    r_sheet = rb.sheet_by_index(idx)
    wb_copy = xl_copy(rb)
    w_sheet = wb_copy.get_sheet(idx)
    return rb, wb_copy, r_sheet, w_sheet

# ----------------------------
# Style cloning (read from original, apply to writes)
# ----------------------------
def clone_cell_style_from_xlrd(rb, r_sheet, r, c):
    try:
        xf_idx = r_sheet.cell_xf_index(r, c)
        xf = rb.xf_list[xf_idx]
    except Exception:
        return XFStyle()

    style = XFStyle()
    # Number format
    try:
        fmt = rb.format_map.get(xf.format_key)
        if fmt and getattr(fmt, "format_str", None):
            style.num_format_str = fmt.format_str
    except Exception:
        pass
    # Font
    try:
        f_rd = rb.font_list[xf.font_index]
        f = Font()
        f.name = f_rd.name; f.bold = f_rd.bold; f.italic = f_rd.italic
        f.height = f_rd.height; f.underline = f_rd.underline_type
        f.colour_index = f_rd.colour_index
        style.font = f
    except Exception:
        pass
    # Alignment
    try:
        a_rd = xf.alignment
        a = Alignment()
        a.horz = a_rd.hor_align; a.vert = a_rd.vert_align
        a.wrap = a_rd.text_wrapped; a.rota = a_rd.rotation
        a.indent = a_rd.indent_level
        style.alignment = a
    except Exception:
        pass
    # Borders
    try:
        b_rd = xf.border
        b = Borders()
        b.left=b_rd.left_line_style; b.right=b_rd.right_line_style
        b.top=b_rd.top_line_style; b.bottom=b_rd.bottom_line_style
        b.left_colour=b_rd.left_colour_index; b.right_colour=b_rd.right_colour_index
        b.top_colour=b_rd.top_colour_index; b.bottom_colour=b_rd.bottom_colour_index
        style.borders = b
    except Exception:
        pass
    # Fill / Pattern
    try:
        bg = xf.background
        p = Pattern()
        p.pattern = bg.fill_pattern
        p.pattern_fore_colour = bg.pattern_colour_index
        p.pattern_back_colour = bg.background_colour_index
        style.pattern = p
    except Exception:
        pass
    return style

def sample_column_style(rb, r_sheet, col_index, start_row_1based, last_row_0based):
    start0 = start_row_1based - 1
    for r in range(start0, last_row_0based + 1):
        try:
            return clone_cell_style_from_xlrd(rb, r_sheet, r, col_index)
        except Exception:
            continue
    return XFStyle()

def write_with_style(ws_writable, row_idx, col_idx, value, style):
    ws_writable.write(row_idx, col_idx, int(value), style)

# ----------------------------
# MRD (Merchant Reference Data)
# ----------------------------
MRD_COLS_REQUIRED = [MRD_COL_VENDOR_SKU, MRD_COL_MERCHANT, MRD_COL_UPC]

def _ensure_mrd_columns(df):
    cols = list(df.columns)
    if len(cols) >= 3:
        lc = [str(c).strip().lower() for c in cols]
        def match(target):
            for i, name in enumerate(lc):
                if target in name: return cols[i]
            return None
        cand_vendor   = match("vendor") or match("sku") or cols[0]
        cand_merchant = match("merchant") or cols[1]
        cand_upc      = match("upc") or cols[2]
        df = df.rename(columns={
            cand_vendor: MRD_COL_VENDOR_SKU,
            cand_merchant: MRD_COL_MERCHANT,
            cand_upc: MRD_COL_UPC
        })
    missing = [c for c in MRD_COLS_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Merchant Reference Data missing required columns: {missing}")
    return df[MRD_COLS_REQUIRED]

def load_mrd_dataframe():
    xlsx = os.path.join(INVENTORY_DIR, MRD_BASENAME_NO_EXT + ".xlsx")
    xls  = os.path.join(INVENTORY_DIR, MRD_BASENAME_NO_EXT + ".xls")
    path = xlsx if os.path.isfile(xlsx) else (xls if os.path.isfile(xls) else None)
    if not path:
        return None
    df = pd.read_excel(path, sheet_name=0, dtype={MRD_COL_UPC: str})
    df = _ensure_mrd_columns(df)
    for c in MRD_COLS_REQUIRED:
        df[c] = df[c].astype(str).fillna("").map(normalize_basic)
    return df

def filter_mrd_for_merchant(df_mrd, merchant_keyword):
    if df_mrd is None: return None
    mk = merchant_keyword.casefold()
    mask = df_mrd[MRD_COL_MERCHANT].str.casefold().str.contains(mk, na=False)
    out = df_mrd.loc[mask].copy()
    return out if not out.empty else None

def build_mrd_lookup(df_mrd):
    if df_mrd is None: return None
    lut = {}
    for _, row in df_mrd.iterrows():
        vsku = normalize_basic(row[MRD_COL_VENDOR_SKU])
        item = {"vendor_sku": vsku,
                "merchant": normalize_basic(row[MRD_COL_MERCHANT]),
                "upc": normalize_basic(row[MRD_COL_UPC])}
        for k in {vsku, normalize_casefold(vsku), normalize_nospace(vsku), normalize_casefold_nospace(vsku)}:
            if not k: continue
            lut.setdefault(k, []).append(item)
    return lut

# Try MRD SKU first, then MRD UPC
def mrd_try_master_match(mrd_rows_for_sku, master_maps):
    for row in mrd_rows_for_sku or []:
        ok, q, src = match_master_sku(row["vendor_sku"], master_maps)
        if ok: return True, q, "MRD SKU → " + src
    for row in mrd_rows_for_sku or []:
        ok, q, src = match_master_upc(row["upc"], master_maps)
        if ok: return True, q, "MRD UPC → " + src
    return False, 0, None

# ----------------------------
# Reporting accumulators
# ----------------------------
class Report:
    def __init__(self):
        self.changes = []   # list of dicts
        self.unmatched = [] # list of (id, retailer)
        self.by_src = {"Master SKU":0, "Master UPC":0, "Master JCP SKU":0,
                       "MRD SKU → Master SKU":0, "MRD UPC → Master UPC":0}
        self.updated_per_retailer = {"QVC":0, "JCP":0, "Macy's":0}
        self.unmatched_per_retailer = {"QVC":0, "JCP":0, "Macy's":0}
        self.deduction_applied = {"QVC":0, "JCP":0}

    def log_change(self, retailer, rownum_1b, sku, upc, source, old_qty, new_qty, deducted=False):
        delta = int(new_qty) - int(old_qty)
        self.changes.append({
            "Retailer": retailer,
            "Row": rownum_1b,
            "SKU": sku,
            "UPC": upc,
            "Match Source": source,
            "Old Qty": int(old_qty),
            "New Qty": int(new_qty),
            "Delta": int(delta),
            "Deduction Applied": bool(deducted),
        })
        self.updated_per_retailer[retailer] += 1
        if deducted and retailer in self.deduction_applied:
            self.deduction_applied[retailer] += 1
        if source in self.by_src:
            self.by_src[source] += 1

    def log_unmatched(self, ident, retailer):
        self.unmatched.append((ident, retailer))
        self.unmatched_per_retailer[retailer] += 1

    def totals(self):
        total_updates = sum(self.updated_per_retailer.values())
        total_unmatched = sum(self.unmatched_per_retailer.values())
        net = sum(int(x["Delta"]) for x in self.changes)
        gross = sum(abs(int(x["Delta"])) for x in self.changes)
        total_old = sum(int(x["Old Qty"]) for x in self.changes)
        total_new = sum(int(x["New Qty"]) for x in self.changes)
        pos = sum(1 for x in self.changes if x["Delta"] > 0)
        neg = sum(1 for x in self.changes if x["Delta"] < 0)
        zero = sum(1 for x in self.changes if x["Delta"] == 0)
        avg = (net / total_updates) if total_updates else 0
        return {
            "total_updates": total_updates,
            "total_unmatched": total_unmatched,
            "net": net,
            "gross": gross,
            "total_old": total_old,
            "total_new": total_new,
            "pos": pos,
            "neg": neg,
            "zero": zero,
            "avg": avg
        }

# ----------------------------
# Per-retailer processors (no Excel app; style clone) + reporting
# ----------------------------
def process_qvc(path, master_maps, mrd_lut_qvc, start_row_1b, report: Report):
    rb, wb, r, w = open_xls(path, QVC_SHEET)
    last = last_data_row(r, [COL_B], start_row_1b)
    if last < start_row_1b - 1:
        wb.save(path); return

    style_D  = sample_column_style(rb, r, COL_D,  start_row_1b, last)
    style_AA = sample_column_style(rb, r, COL_AA, start_row_1b, last)

    for rix in range(start_row_1b - 1, last + 1):
        sku = r.cell_value(rix, COL_B)
        sku_norm = normalize_basic(sku)
        if not sku_norm:
            continue

        old_qty = quantity_int(r.cell_value(rix, COL_D))

        ok, q, src = match_master_sku(sku, master_maps)
        if not ok and mrd_lut_qvc is not None:
            rows = []
            for k in [sku_norm, normalize_casefold(sku_norm), normalize_nospace(sku_norm), normalize_casefold_nospace(sku_norm)]:
                rows.extend(mrd_lut_qvc.get(k, []))
            if rows:
                ok, q, src = mrd_try_master_match(rows, master_maps)

        if ok:
            new_q = max(q - 5, 0)  # deduction
            write_with_style(w, rix, COL_D,  new_q, style_D)
            write_with_style(w, rix, COL_AA, new_q, style_AA)
            report.log_change("QVC", rix+1, sku_norm, "", src, old_qty, new_q, deducted=True)
        else:
            report.log_unmatched(sku_norm, "QVC")

    wb.save(path)

def process_jcp(path, master_maps, start_row_1b, report: Report):
    rb, wb, r, w = open_xls(path, JCP_SHEET)
    last = last_data_row(r, [COL_B], start_row_1b)
    if last < start_row_1b - 1:
        wb.save(path); return

    style_D = sample_column_style(rb, r, COL_D, start_row_1b, last)

    for rix in range(start_row_1b - 1, last + 1):
        sku = r.cell_value(rix, COL_B)
        sku_norm = normalize_basic(sku)
        if not sku_norm:
            continue

        old_qty = quantity_int(r.cell_value(rix, COL_D))

        ok, q, src = match_master_sku(sku, master_maps)
        if not ok:
            ok, q, src = match_master_jcp_sku(sku, master_maps)

        if ok:
            new_q = max(q - 5, 0)  # deduction
            write_with_style(w, rix, COL_D, new_q, style_D)
            report.log_change("JCP", rix+1, sku_norm, "", src, old_qty, new_q, deducted=True)
        else:
            report.log_unmatched(sku_norm, "JCP")

    wb.save(path)

def process_macys(path, master_maps, mrd_lut_macys, start_row_1b, report: Report):
    rb, wb, r, w = open_xls(path, MACYS_SHEET)
    last = last_data_row(r, [COL_B, COL_Y], start_row_1b)
    if last < start_row_1b - 1:
        wb.save(path); return

    style_D  = sample_column_style(rb, r, COL_D,  start_row_1b, last)
    style_AA = sample_column_style(rb, r, COL_AA, start_row_1b, last)

    for rix in range(start_row_1b - 1, last + 1):
        sku = r.cell_value(rix, COL_B)
        upc = r.cell_value(rix, COL_Y)
        sku_norm = normalize_basic(sku)
        upc_norm = normalize_upc(upc)

        old_qty = quantity_int(r.cell_value(rix, COL_D))

        ok, q, src = (False, 0, None)
        if sku_norm:
            ok, q, src = match_master_sku(sku, master_maps)
        if not ok and upc_norm:
            ok, q, src = match_master_upc(upc, master_maps)
        if not ok and mrd_lut_macys is not None and sku_norm:
            rows = []
            for k in [sku_norm, normalize_casefold(sku_norm), normalize_nospace(sku_norm), normalize_casefold_nospace(sku_norm)]:
                rows.extend(mrd_lut_macys.get(k, []))
            if rows:
                ok, q, src = mrd_try_master_match(rows, master_maps)

        if ok:
            new_q = q  # no deduction
            write_with_style(w, rix, COL_D,  new_q, style_D)
            write_with_style(w, rix, COL_AA, new_q, style_AA)
            report.log_change("Macy's", rix+1, sku_norm, upc_norm, src, old_qty, new_q, deducted=False)
        else:
            ident = sku_norm if sku_norm else upc_norm
            if ident:
                report.log_unmatched(ident, "Macy's")

    wb.save(path)

# ----------------------------
# History (append) + Trends
# ----------------------------
def append_to_history(history_path, changes_df, run_id, run_ts):
    """Append current run's Changes to the persistent history workbook."""
    cols = ["Run ID","Run Timestamp","Retailer","Row","SKU","UPC","Match Source",
            "Old Qty","New Qty","Delta","Deduction Applied"]
    changes_df = changes_df.copy()
    changes_df.insert(0, "Run ID", run_id)
    changes_df.insert(1, "Run Timestamp", run_ts)

    if os.path.isfile(history_path):
        try:
            hist = pd.read_excel(history_path, sheet_name="History")
            hist = pd.concat([hist, changes_df[cols]], ignore_index=True)
        except Exception:
            # If the existing file is malformed, start fresh
            hist = changes_df[cols]
    else:
        hist = changes_df[cols]

    with pd.ExcelWriter(history_path, engine="xlsxwriter") as writer:
        hist.to_excel(writer, index=False, sheet_name="History")

    return hist

def compute_trends(history_df):
    """Return several trend DataFrames based on last 30 days + last 10 runs."""
    if history_df is None or history_df.empty:
        empty = pd.DataFrame()
        return {"Last 10 Runs (Net Δ)": empty, "30d Net Δ by Retailer": empty,
                "30d Top Movers": empty, "30d Repeated Changes by SKU": empty}

    # Ensure proper types
    df = history_df.copy()
    if not pd.api.types.is_datetime64_any_dtype(df["Run Timestamp"]):
        df["Run Timestamp"] = pd.to_datetime(df["Run Timestamp"])
    df["Delta"] = pd.to_numeric(df["Delta"], errors="coerce").fillna(0).astype(int)

    # Last 10 runs net change (overall + by retailer)
    last10_ids = (df.drop_duplicates(["Run ID"])
                    .sort_values("Run Timestamp", ascending=False)["Run ID"].head(10))
    last10 = df[df["Run ID"].isin(last10_ids)]
    last10_overall = (last10.groupby("Run ID")["Delta"].sum()
                      .reset_index().rename(columns={"Delta":"Net Δ (All)"}))
    last10_by_ret = (last10.pivot_table(index="Run ID", columns="Retailer", values="Delta", aggfunc="sum")
                     .reset_index())
    last10_out = pd.merge(last10_overall, last10_by_ret, on="Run ID", how="left").sort_values("Run ID", ascending=False)

    # 30-day window
    cutoff = df["Run Timestamp"].max() - timedelta(days=30)
    last30 = df[df["Run Timestamp"] >= cutoff]

    # 30d Net Δ by Retailer
    d30_by_ret = (last30.groupby("Retailer")["Delta"].sum()
                  .reset_index().rename(columns={"Delta":"30d Net Δ"}).sort_values("30d Net Δ", ascending=False))

    # 30d Top Movers (by absolute cumulative delta) – SKU level + retailer
    movers = (last30.groupby(["Retailer","SKU"])["Delta"].sum()
              .reset_index())
    movers["Abs Δ"] = movers["Delta"].abs()
    top_movers = movers.sort_values("Abs Δ", ascending=False).head(20)

    # 30d Repeated Changes – count of times a SKU changed (frequency)
    freq = (last30.groupby(["Retailer","SKU"])["Delta"]
                    .apply(lambda s: (s != 0).sum()).reset_index(name="Change Count"))
    freq = freq.sort_values("Change Count", ascending=False).head(20)

    return {
        "Last 10 Runs (Net Δ)": last10_out,
        "30d Net Δ by Retailer": d30_by_ret,
        "30d Top Movers": top_movers,
        "30d Repeated Changes by SKU": freq
    }

# ----------------------------
# Report writer (streamlined KPIs + Analysis + Trends)
# ----------------------------
def write_inventory_report(report, trends, date_str, analysis_lines):
    report_path = os.path.join(INVENTORY_DIR, f"Inventory Report {date_str}.xlsx")

    totals = report.totals()
    tot_updates   = totals["total_updates"]
    tot_unmatched = totals["total_unmatched"]
    unmatched_rate = (tot_unmatched / (tot_updates + tot_unmatched)) if (tot_updates + tot_unmatched) else 0.0

    # Minimal but powerful KPIs
    summary_rows = [
        ["Total SKUs Updated", tot_updates],
        ["Total Unmatched", tot_unmatched],
        ["Unmatched Rate", round(unmatched_rate, 4)],
        ["Net Quantity Change", totals["net"]],
        ["Gross Quantity Change", totals["gross"]],
        ["Old Qty (updated rows)", totals["total_old"]],
        ["New Qty (updated rows)", totals["total_new"]],
        ["Positive Δ", totals["pos"]],
        ["Negative Δ", totals["neg"]],
        ["Zero Δ", totals["zero"]],
        ["Average Δ / updated SKU", round(totals["avg"], 4)],
        ["QVC: SKUs Updated", report.updated_per_retailer["QVC"]],
        ["JCP: SKUs Updated", report.updated_per_retailer["JCP"]],
        ["Macy's: SKUs Updated", report.updated_per_retailer["Macy's"]],
        ["QVC: -5 Deduction Applied", report.deduction_applied["QVC"]],
        ["JCP: -5 Deduction Applied", report.deduction_applied["JCP"]],
    ]
    df_summary = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    # Changes (full log) – useful for drilldowns
    df_changes = pd.DataFrame(report.changes,
                              columns=["Retailer","Row","SKU","UPC","Match Source","Old Qty","New Qty","Delta","Deduction Applied"])

    # Top increases/decreases (today’s run)
    if len(df_changes):
        df_top_inc = df_changes.sort_values("Delta", ascending=False).head(10)
        df_top_dec = df_changes.sort_values("Delta", ascending=True).head(10)
    else:
        df_top_inc = pd.DataFrame(columns=df_changes.columns)
        df_top_dec = pd.DataFrame(columns=df_changes.columns)

    # Unmatched (today’s run)
    df_unmatched = pd.DataFrame(report.unmatched, columns=["Unmatched SKU/UPC","Retailer"])

    # Write (overwrite per day)
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        # Core tabs
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        df_top_inc.to_excel(writer, index=False, sheet_name="Top Increases")
        df_top_dec.to_excel(writer, index=False, sheet_name="Top Decreases")
        df_unmatched.to_excel(writer, index=False, sheet_name="Unmatched")
        df_changes.to_excel(writer, index=False, sheet_name="Changes")

        # Trends tabs
        for name, df in trends.items():
            (df if df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name=name[:31])

        # Analysis (plain-English)
        ws = writer.book.add_worksheet("Analysis")
        wrap = writer.book.add_format({"text_wrap": True, "valign": "top"})
        bold = writer.book.add_format({"bold": True})
        ws.write(0, 0, f"Inventory Analysis — {date_str}", bold)
        row = 2
        for para in analysis_lines:
            ws.write(row, 0, para, wrap)
            row += 2  # blank line between paragraphs
        ws.set_column(0, 0, 110)

        # Autosize helper
        def autosize(ws_name, df, extras=None):
            if df is None or df.empty: return
            ws = writer.sheets[ws_name]
            for i, col in enumerate(df.columns):
                width = max(len(str(col)), *(len(str(x)) for x in df[col].astype(str).head(1000)))
                ws.set_column(i, i, min(max(width + 2, 8), 60))
            ws.set_row(0, None, writer.book.add_format({"bold": True}))
            if extras:
                for idx, w in extras.items():
                    ws.set_column(idx, idx, w)

        autosize("Summary", df_summary, {0:40, 1:18})
        autosize("Top Increases", df_top_inc)
        autosize("Top Decreases", df_top_dec)
        autosize("Unmatched", df_unmatched, {0:30,1:12})
        autosize("Changes", df_changes)

        for name in trends.keys():
            autosize(name[:31], trends[name])

    return report_path

# ----------------------------
# Natural-language analysis
# ----------------------------
def build_analysis(report, trends, run_date_str):
    t = report.totals()
    updated = t["total_updates"]; unmatched = t["total_unmatched"]
    net = t["net"]; gross = t["gross"]
    unmatched_rate = (unmatched / (updated + unmatched)) if (updated + unmatched) else 0.0

    # Retailer highlights
    byret = report.updated_per_retailer
    unmatched_byret = report.unmatched_per_retailer
    ded_qvc = report.deduction_applied["QVC"]
    ded_jcp = report.deduction_applied["JCP"]

    # Pull Macy's values into variables to avoid quoting inside f-strings
    macys_updated = byret.get("Macy's", 0)
    macys_unmatched = unmatched_byret.get("Macy's", 0)
    qvc_updated = byret.get("QVC", 0)
    jcp_updated = byret.get("JCP", 0)
    qvc_unmatched = unmatched_byret.get("QVC", 0)
    jcp_unmatched = unmatched_byret.get("JCP", 0)

    lines = []
    lines.append(
        f"Run date {run_date_str}: {updated} SKUs were updated and {unmatched} were unmatched "
        f"({unmatched_rate:.1%} unmatched rate). Net inventory change across updated SKUs was {net:+,} "
        f"({gross:,} gross movement)."
    )
    lines.append(
        f"By retailer — QVC updated {qvc_updated} (deductions applied on {ded_qvc}), "
        f"JCP updated {jcp_updated} (deductions applied on {ded_jcp}), "
        f"Macy's updated {macys_updated}. "
        f"Unmatched rows: QVC {qvc_unmatched}, JCP {jcp_unmatched}, Macy's {macys_unmatched}."
    )

    # Trend callouts
    last10 = trends.get("Last 10 Runs (Net Δ)")
    if last10 is not None and not last10.empty:
        most_recent = last10.iloc[0]
        lines.append(
            "Recent momentum (last 10 runs): "
            f"Most recent net change was {int(most_recent.get('Net Δ (All)', 0)):+,}."
        )
    d30 = trends.get("30d Net Δ by Retailer")
    if d30 is not None and not d30.empty:
        top = d30.iloc[0]
        lines.append(
            f"Past 30 days: highest net movement by retailer was {top['Retailer']} "
            f"at {int(top['30d Net Δ']):+,}."
        )

    movers = trends.get("30d Top Movers")
    if movers is not None and not movers.empty:
        m = movers.head(3)
        bullets = "; ".join([f"{r['Retailer']} / {r['SKU']} ({int(r['Delta']):+d})" for _, r in m.iterrows()])
        lines.append(f"Top movers in the last 30 days: {bullets}.")

    freq = trends.get("30d Repeated Changes by SKU")
    if freq is not None and not freq.empty:
        f = freq.head(3)
        bullets = "; ".join([f"{r['Retailer']} / {r['SKU']} ({int(r['Change Count'])} changes)" for _, r in f.iterrows()])
        lines.append(f"Items with frequent updates (may need attention): {bullets}.")

    lines.append(
        "Recommendations: review repeatedly-updated SKUs for supply or listing issues; "
        "resolve frequently-unmatched SKUs by enriching Master and Merchant Reference Data; "
        "spot-check largest positive/negative deltas against recent purchase orders and returns."
    )
    return lines

# ----------------------------
# Main
# ----------------------------
def main():
    if not os.path.isdir(INVENTORY_DIR):
        raise FileNotFoundError(f"Inventory folder missing: {INVENTORY_DIR}")

    # Select Master file
    root = Tk(); root.withdraw()
    master_path = filedialog.askopenfilename(
        title="Select Master Excel file",
        filetypes=[("Excel", "*.xls *.xlsx *.xlsm *.xlsb")]
    )
    root.destroy()
    if not master_path:
        print("No Master selected."); return

    # Load Master (preserve UPC as string)
    df_master = pd.read_excel(master_path, sheet_name=MASTER_SHEET, dtype={MASTER_COL_C_UPC: str})
    master_maps = make_lookup_dicts(df_master)

    # Load MRD (optional)
    def load_mrd_dataframe_local():
        xlsx = os.path.join(INVENTORY_DIR, MRD_BASENAME_NO_EXT + ".xlsx")
        xls  = os.path.join(INVENTORY_DIR, MRD_BASENAME_NO_EXT + ".xls")
        path = xlsx if os.path.isfile(xlsx) else (xls if os.path.isfile(xls) else None)
        if not path: return None
        df = pd.read_excel(path, sheet_name=0, dtype={MRD_COL_UPC: str})
        df = _ensure_mrd_columns(df)
        for c in (MRD_COL_VENDOR_SKU, MRD_COL_MERCHANT, MRD_COL_UPC):
            df[c] = df[c].astype(str).fillna("").map(normalize_basic)
        return df

    df_mrd = load_mrd_dataframe_local()
    mrd_lut_qvc = mrd_lut_macys = None
    if df_mrd is not None:
        df_qvc   = filter_mrd_for_merchant(df_mrd, "QVC Drop Ship")
        df_macys = filter_mrd_for_merchant(df_mrd, "Macy")
        mrd_lut_qvc   = build_mrd_lookup(df_qvc)   if df_qvc   is not None else None
        mrd_lut_macys = build_mrd_lookup(df_macys) if df_macys is not None else None

    report = Report()

    # Process all three without launching Excel
    process_qvc(os.path.join(INVENTORY_DIR, QVC_FILE),     master_maps, mrd_lut_qvc,     DATA_START_ROW, report)
    process_jcp(os.path.join(INVENTORY_DIR, JCP_FILE),     master_maps,                  DATA_START_ROW, report)
    process_macys(os.path.join(INVENTORY_DIR, MACYS_FILE), master_maps, mrd_lut_macys,   DATA_START_ROW, report)

    # Build current run DataFrame (Changes)
    df_changes = pd.DataFrame(report.changes,
                              columns=["Retailer","Row","SKU","UPC","Match Source","Old Qty","New Qty","Delta","Deduction Applied"])

    # Persist to history
    run_ts = datetime.now()
    run_id = run_ts.strftime("%Y%m%d-%H%M%S")
    hist_df = append_to_history(HISTORY_PATH, df_changes, run_id, run_ts)

    # Compute trends from history
    trends = compute_trends(hist_df)

    # Build plain-English analysis
    date_str = run_ts.strftime("%m.%d.%y")
    analysis_lines = build_analysis(report, trends, date_str)

    # Write daily report
    report_path = write_inventory_report(report, trends, date_str, analysis_lines)
    print(f"Inventory report written to: {report_path}")
    print(f"History updated: {HISTORY_PATH}")
    print("Inventory update completed.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
