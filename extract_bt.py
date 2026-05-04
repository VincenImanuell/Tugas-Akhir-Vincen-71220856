"""
Ekstrak semua hasil backtesting dari folder BT_RSI/
Output: summary_bt.xlsx dengan sheet per rumusan masalah
"""

import os, glob, re
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── MAPPING: filename stem → (sektor, [metode]) ──────────────────────────────
# metode: "Pearson", "Spearman", "Kointegrasi"
# Nama file sesuai urutan Y-X hasil ADF / kointegrasi

PAIR_META = {
    # Pearson & Spearman (overlap)
    "IMPC-HEXA":  ("Industri",                   ["Pearson", "Spearman"]),
    "ADRO-ITMG":  ("Energi",                      ["Pearson", "Spearman"]),
    "SPMA-UNIC":  ("Barang Baku",                 ["Pearson", "Spearman"]),
    "UNIC-SPMA":  ("Barang Baku",                 ["Pearson", "Spearman"]),
    "ADHI-PTPP":  ("Infrastruktur",               ["Pearson", "Spearman"]),
    "SMDR-TMAS":  ("Transportasi & Logistik",     ["Pearson", "Spearman"]),
    "PNBN-PNLF":  ("Keuangan",                    ["Pearson"]),
    "POLL-URBN":  ("Properti",                    ["Pearson", "Spearman"]),
    "DMMX-MLPT":  ("Teknologi",                   ["Pearson", "Spearman"]),
    "AMRT-ADES":  ("Barang Konsumsi Primer",      ["Pearson", "Spearman"]),
    "MAPI-PANR":  ("Barang Konsumsi Non-Primer",  ["Pearson", "Spearman"]),
    "KAEF-INAF":  ("Kesehatan",                   ["Pearson", "Spearman"]),
    # Spearman only
    "BBNI-BMRI":  ("Keuangan",                    ["Spearman"]),
    # Kointegrasi
    "SOSS-SKRN":  ("Industri",                    ["Kointegrasi"]),
    "KOPI-DWGL":  ("Energi",                      ["Kointegrasi"]),
    "PICO-ALMI":  ("Barang Baku",                 ["Kointegrasi"]),
    "JAST-OASA":  ("Infrastruktur",               ["Kointegrasi"]),
    "AKSI-HELI":  ("Transportasi & Logistik",     ["Kointegrasi"]),
    "JMAS-BBRI":  ("Keuangan",                    ["Kointegrasi"]),
    "JMAS-BBNI":  ("Keuangan",                    ["Kointegrasi"]),
    "NZIA-PAMG":  ("Properti",                    ["Kointegrasi"]),
    "ATIC-DIVA":  ("Teknologi",                   ["Kointegrasi"]),
    "SKBM-MGRO":  ("Barang Konsumsi Primer",      ["Kointegrasi"]),
    "GDYR-GLOB":  ("Barang Konsumsi Non-Primer",  ["Kointegrasi"]),
    "MIKA-PRIM":  ("Kesehatan",                   ["Kointegrasi"]),
}

FOLDER = "BT_RSI"


def pct(s):
    """Parse '12.34%' or numeric → float percentage value."""
    if pd.isna(s):
        return np.nan
    s = str(s).strip().replace("%", "")
    try:
        return float(s)
    except:
        return np.nan


def read_summary(path):
    """Baca sheet Comparison Summary dan kembalikan dict metrics."""
    try:
        df = pd.read_excel(path, sheet_name="Comparison Summary", header=None)
    except Exception as e:
        print(f"  ERROR membaca {path}: {e}")
        return None

    data = {}
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        val_b = row.iloc[1] if len(row) > 1 else np.nan
        val_r = row.iloc[2] if len(row) > 2 else np.nan
        if label:
            data[label] = (val_b, val_r)

    def g(key, idx=0):
        v = data.get(key, (np.nan, np.nan))
        return v[idx] if isinstance(v, tuple) else v

    stem = os.path.splitext(os.path.basename(path))[0]
    meta = PAIR_META.get(stem, ("Unknown", ["Unknown"]))

    result = {
        "Pair":           stem,
        "Sektor":         meta[0],
        "Metode":         ", ".join(meta[1]),
        "Pair Type":      str(g("Pair Type")),
        "Y":              str(g("Stock Y (Dependen)")),
        "X":              str(g("Stock X (Independen)")),
        # Base
        "Base_Return":    pct(g("Total Return (%)", 0)),
        "Base_Sharpe":    pct(g("Sharpe Ratio", 0)),
        "Base_MDD":       pct(g("Max Drawdown (%)", 0)),
        "Base_WinRate":   pct(g("Win Rate (%)", 0)),
        "Base_Volatility":pct(g("Annual Volatility (%)", 0)),
        "Base_Trades":    g("Total Trades", 0),
        "Base_TP":        g("TP Hits", 0),
        "Base_SL":        g("SL Hits", 0),
        "Base_ZExit":     g("Z-Exits", 0),
        # RSI
        "RSI_Return":     pct(g("Total Return (%)", 1)),
        "RSI_Sharpe":     pct(g("Sharpe Ratio", 1)),
        "RSI_MDD":        pct(g("Max Drawdown (%)", 1)),
        "RSI_WinRate":    pct(g("Win Rate (%)", 1)),
        "RSI_Volatility": pct(g("Annual Volatility (%)", 1)),
        "RSI_Trades":     g("Total Trades", 1),
        "RSI_TP":         g("TP Hits", 1),
        "RSI_SL":         g("SL Hits", 1),
        "RSI_ZExit":      g("Z-Exits", 1),
        # Benchmark
        "BH_Y":           pct(g(f"Buy & Hold Y ({str(g('Stock Y (Dependen)'))})")),
        "BH_X":           pct(g(f"Buy & Hold X ({str(g('Stock X (Independen)'))})")),
        # vs BH
        "Base_vs_BH_Y":   pct(g("vs B&H Y (Base)")),
        "RSI_vs_BH_Y":    pct(g("vs B&H Y (RSI)")),
        "Base_vs_BH_X":   pct(g("vs B&H X (Base)")),
        "RSI_vs_BH_X":    pct(g("vs B&H X (RSI)")),
        # Parameters
        "z_entry":        g("z_entry (±)"),
        "z_exit":         g("z_exit"),
        "Z_Window":       g("Z-Score Window"),
        "SL":             str(g("Stop Loss (best)")),
        "TP":             str(g("Take Profit (best)")),
        "RSI_Period":     g("RSI Period (best)"),
        "RSI_Threshold":  g("RSI Threshold (best)"),
    }
    return result


def read_grid_search(path):
    return {}


# ── MAIN ─────────────────────────────────────────────────────────────────────

files = sorted(glob.glob(os.path.join(FOLDER, "*.xlsx")))
print(f"Ditemukan {len(files)} file di {FOLDER}/\n")

rows = []

for f in files:
    stem = os.path.splitext(os.path.basename(f))[0]
    print(f"  Membaca: {stem}")
    r = read_summary(f)
    if r:
        rows.append(r)

df_all = pd.DataFrame(rows)

print(f"\nTotal pair berhasil dibaca: {len(df_all)}")

# ── HELPER STYLES ─────────────────────────────────────────────────────────────

def make_wb_styles():
    return {
        "hdr":   Font(name="Arial", bold=True, color="FFFFFF", size=11),
        "hdr_f": PatternFill("solid", start_color="1F4E79"),
        "sec_f": PatternFill("solid", start_color="2E75B6"),
        "lbl_f": PatternFill("solid", start_color="D6E4F0"),
        "grn_f": PatternFill("solid", start_color="E2EFDA"),
        "red_f": PatternFill("solid", start_color="FFDAD9"),
        "gld_f": PatternFill("solid", start_color="FFF2CC"),
        "ctr":   Alignment(horizontal="center"),
        "lft":   Alignment(horizontal="left"),
        "bdr":   Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin")),
        "fnt":   Font(name="Arial", size=11),
    }

S = make_wb_styles()

def write_df_to_sheet(ws, df, title=None, color_col=None, higher_is_better=True):
    """Tulis DataFrame ke sheet dengan header styling."""
    start_row = 1
    if title:
        ws.append([title])
        ws[ws.max_row][0].font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
        ws[ws.max_row][0].fill = S["hdr_f"]
        ws.merge_cells(f"A1:{get_column_letter(len(df.columns))}1")
        ws.append([])
        start_row = 3

    ws.append(list(df.columns))
    for cell in ws[ws.max_row]:
        cell.font = S["hdr"]
        cell.fill = S["hdr_f"]
        cell.alignment = S["ctr"]
        cell.border = S["bdr"]

    for _, row_data in df.iterrows():
        ws.append(list(row_data))
        row_idx = ws.max_row
        for cell in ws[row_idx]:
            cell.font = S["fnt"]
            cell.alignment = S["ctr"]
            cell.border = S["bdr"]
        # Color coding pada kolom tertentu
        if color_col and color_col in df.columns:
            col_idx = list(df.columns).index(color_col) + 1
            val = row_data[color_col]
            try:
                val = float(str(val).replace("%",""))
                fill = S["grn_f"] if (val > 0) == higher_is_better else S["red_f"]
                ws.cell(row_idx, col_idx).fill = fill
            except:
                pass

    for col in ws.columns:
        w = max(len(str(c.value or "")) for c in col) + 4
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w, 30)
    ws.freeze_panes = f"A{start_row+1}"


def agg_metrics(df, prefix):
    """Agregasi metrik base atau rsi."""
    return {
        "Return (%)":      df[f"{prefix}_Return"].mean(),
        "Sharpe Ratio":    df[f"{prefix}_Sharpe"].mean(),
        "MDD (%)":         df[f"{prefix}_MDD"].mean(),
        "Win Rate (%)":    df[f"{prefix}_WinRate"].mean(),
        "Volatility (%)":  df[f"{prefix}_Volatility"].mean(),
        "Avg Trades":      df[f"{prefix}_Trades"].mean(),
        "Positive Pairs":  (df[f"{prefix}_Return"] > 0).sum(),
        "Total Pairs":     len(df),
    }


# ── BUILD OUTPUT EXCEL ────────────────────────────────────────────────────────

wb = Workbook()

# ─── Sheet 0: Raw Data Lengkap ───────────────────────────────────────────────
ws0 = wb.active
ws0.title = "Raw Data"
write_df_to_sheet(ws0, df_all, title="Semua Pair — Data Lengkap")

# ─── Sheet RQ1a: Per Metode (Base) ────────────────────────────────────────────
ws1 = wb.create_sheet("RQ1 - Per Metode (Base)")
METODE_LIST = ["Pearson", "Spearman", "Kointegrasi"]

rows_rq1 = []
for m in METODE_LIST:
    subset = df_all[df_all["Metode"].str.contains(m, na=False)]
    for _, r in subset.iterrows():
        rows_rq1.append({
            "Metode": m, "Pair": r["Pair"], "Sektor": r["Sektor"],
            "Return (%)": r["Base_Return"], "Sharpe": r["Base_Sharpe"],
            "MDD (%)": r["Base_MDD"], "Win Rate (%)": r["Base_WinRate"],
            "Volatility (%)": r["Base_Volatility"], "Trades": r["Base_Trades"],
            "vs B&H Y (%)": r["Base_vs_BH_Y"],
        })

df_rq1 = pd.DataFrame(rows_rq1)

# Ringkasan per metode
rows_rq1_sum = []
for m in METODE_LIST:
    subset = df_all[df_all["Metode"].str.contains(m, na=False)]
    if len(subset) == 0: continue
    ag = agg_metrics(subset, "Base")
    ag["Metode"] = m
    ag["Jumlah Pair"] = len(subset)
    rows_rq1_sum.append(ag)

df_rq1_sum = pd.DataFrame(rows_rq1_sum)[["Metode","Jumlah Pair","Return (%)","Sharpe Ratio",
                                           "MDD (%)","Win Rate (%)","Volatility (%)","Avg Trades","Positive Pairs"]]

write_df_to_sheet(ws1, df_rq1_sum, title="RQ1 — Perbandingan Antar Metode (Strategi Basis)", color_col="Return (%)")
ws1.append([])
ws1.append(["Detail per Pair"])
ws1[ws1.max_row][0].font = Font(name="Arial", bold=True, color="FFFFFF")
ws1[ws1.max_row][0].fill = S["sec_f"]
ws1.append([])

for row_data in df_rq1.values.tolist():
    ws1.append(row_data)

# ─── Sheet RQ1b: Per Metode (RSI) ─────────────────────────────────────────────
ws1r = wb.create_sheet("RQ1 - Per Metode (RSI)")
rows_rq1r = []
for m in METODE_LIST:
    subset = df_all[df_all["Metode"].str.contains(m, na=False)]
    for _, r in subset.iterrows():
        rows_rq1r.append({
            "Metode": m, "Pair": r["Pair"], "Sektor": r["Sektor"],
            "Return (%)": r["RSI_Return"], "Sharpe": r["RSI_Sharpe"],
            "MDD (%)": r["RSI_MDD"], "Win Rate (%)": r["RSI_WinRate"],
            "Volatility (%)": r["RSI_Volatility"], "Trades": r["RSI_Trades"],
            "vs B&H Y (%)": r["RSI_vs_BH_Y"],
        })

df_rq1r = pd.DataFrame(rows_rq1r)
rows_rq1r_sum = []
for m in METODE_LIST:
    subset = df_all[df_all["Metode"].str.contains(m, na=False)]
    if len(subset) == 0: continue
    ag = agg_metrics(subset, "RSI")
    ag["Metode"] = m
    ag["Jumlah Pair"] = len(subset)
    rows_rq1r_sum.append(ag)

df_rq1r_sum = pd.DataFrame(rows_rq1r_sum)[["Metode","Jumlah Pair","Return (%)","Sharpe Ratio",
                                              "MDD (%)","Win Rate (%)","Volatility (%)","Avg Trades","Positive Pairs"]]
write_df_to_sheet(ws1r, df_rq1r_sum, title="RQ1 — Perbandingan Antar Metode (RSI-Filtered)", color_col="Return (%)")

# ─── Sheet RQ2: Per Sektor ────────────────────────────────────────────────────
ws2 = wb.create_sheet("RQ2 - Per Sektor")
SEKTOR_LIST = sorted(df_all["Sektor"].unique())

rows_rq2 = []
for s in SEKTOR_LIST:
    subset = df_all[df_all["Sektor"] == s]
    for prefix, label in [("Base", "Basis"), ("RSI", "RSI")]:
        ag = agg_metrics(subset, prefix)
        rows_rq2.append({
            "Sektor": s, "Strategi": label,
            "Jumlah Pair": ag["Total Pairs"],
            "Return (%)": round(ag["Return (%)"], 2),
            "Sharpe Ratio": round(ag["Sharpe Ratio"], 4),
            "MDD (%)": round(ag["MDD (%)"], 2),
            "Win Rate (%)": round(ag["Win Rate (%)"], 2),
            "Volatility (%)": round(ag["Volatility (%)"], 2),
            "Avg Trades": round(ag["Avg Trades"], 1),
            "Positive Pairs": int(ag["Positive Pairs"]),
        })

df_rq2 = pd.DataFrame(rows_rq2)
write_df_to_sheet(ws2, df_rq2, title="RQ2 — Kinerja Per Sektor", color_col="Return (%)")

# Detail pair per sektor
ws2.append([])
ws2.append(["Detail Pair per Sektor"])
ws2[ws2.max_row][0].font = Font(name="Arial", bold=True, color="FFFFFF")
ws2[ws2.max_row][0].fill = S["sec_f"]

for s in SEKTOR_LIST:
    subset = df_all[df_all["Sektor"] == s]
    ws2.append([])
    ws2.append([s])
    ws2[ws2.max_row][0].font = Font(name="Arial", bold=True, size=12)
    ws2[ws2.max_row][0].fill = S["lbl_f"]

    header = ["Pair","Metode","Return_Base(%)","Sharpe_Base","MDD_Base(%)","WR_Base(%)","Trades_Base",
              "Return_RSI(%)","Sharpe_RSI","MDD_RSI(%)","WR_RSI(%)","Trades_RSI"]
    ws2.append(header)
    for cell in ws2[ws2.max_row]:
        cell.font = S["hdr"]
        cell.fill = S["hdr_f"]
        cell.alignment = S["ctr"]
        cell.border = S["bdr"]

    for _, r in subset.iterrows():
        ws2.append([
            r["Pair"], r["Metode"],
            r["Base_Return"], r["Base_Sharpe"], r["Base_MDD"], r["Base_WinRate"], r["Base_Trades"],
            r["RSI_Return"],  r["RSI_Sharpe"],  r["RSI_MDD"],  r["RSI_WinRate"],  r["RSI_Trades"],
        ])

# ─── Sheet RQ3: Basis vs RSI ──────────────────────────────────────────────────
ws3 = wb.create_sheet("RQ3 - Basis vs RSI")

df_rq3 = df_all[["Pair","Sektor","Metode",
                   "Base_Return","RSI_Return",
                   "Base_Sharpe","RSI_Sharpe",
                   "Base_MDD","RSI_MDD",
                   "Base_WinRate","RSI_WinRate",
                   "Base_Trades","RSI_Trades"]].copy()

df_rq3["Δ_Return"]   = df_rq3["RSI_Return"]   - df_rq3["Base_Return"]
df_rq3["Δ_Sharpe"]   = df_rq3["RSI_Sharpe"]   - df_rq3["Base_Sharpe"]
df_rq3["Δ_MDD"]      = df_rq3["RSI_MDD"]      - df_rq3["Base_MDD"]
df_rq3["Δ_WinRate"]  = df_rq3["RSI_WinRate"]  - df_rq3["Base_WinRate"]
df_rq3["Δ_Trades"]   = df_rq3["RSI_Trades"]   - df_rq3["Base_Trades"]

df_rq3 = df_rq3[["Pair","Sektor","Metode",
                   "Base_Return","RSI_Return","Δ_Return",
                   "Base_Sharpe","RSI_Sharpe","Δ_Sharpe",
                   "Base_MDD","RSI_MDD","Δ_MDD",
                   "Base_WinRate","RSI_WinRate","Δ_WinRate",
                   "Base_Trades","RSI_Trades","Δ_Trades"]]

write_df_to_sheet(ws3, df_rq3, title="RQ3 — Perbandingan Basis vs RSI-Filtered", color_col="Δ_Return")

# Ringkasan agregat RQ3
ws3.append([])
ws3.append(["Ringkasan Agregat"])
ws3[ws3.max_row][0].font = Font(name="Arial", bold=True, color="FFFFFF")
ws3[ws3.max_row][0].fill = S["sec_f"]

metrics_3 = {
    "Avg Return Basis (%)":      df_all["Base_Return"].mean(),
    "Avg Return RSI (%)":        df_all["RSI_Return"].mean(),
    "Avg Sharpe Basis":          df_all["Base_Sharpe"].mean(),
    "Avg Sharpe RSI":            df_all["RSI_Sharpe"].mean(),
    "Avg MDD Basis (%)":         df_all["Base_MDD"].mean(),
    "Avg MDD RSI (%)":           df_all["RSI_MDD"].mean(),
    "Avg Win Rate Basis (%)":    df_all["Base_WinRate"].mean(),
    "Avg Win Rate RSI (%)":      df_all["RSI_WinRate"].mean(),
    "RSI lebih baik (Return)":   (df_all["RSI_Return"] > df_all["Base_Return"]).sum(),
    "RSI lebih baik (MDD)":      (df_all["RSI_MDD"] < df_all["Base_MDD"]).sum(),
    "RSI lebih baik (Sharpe)":   (df_all["RSI_Sharpe"] > df_all["Base_Sharpe"]).sum(),
}
for k, v in metrics_3.items():
    ws3.append([k, round(v, 4) if isinstance(v, float) else v])

# ─── Sheet RQ4: Keseluruhan ────────────────────────────────────────────────────
ws4 = wb.create_sheet("RQ4 - Kinerja Keseluruhan")

# Tabel A: Semua pair, basis + RSI side by side
cols_rq4 = ["Pair","Sektor","Metode","Y","X",
            "Base_Return(%)","Base_Sharpe","Base_MDD(%)","Base_WinRate(%)","Base_Trades",
            "RSI_Return(%)","RSI_Sharpe","RSI_MDD(%)","RSI_WinRate(%)","RSI_Trades",
            "BH_Y(%)","BH_X(%)",
            "Base_vs_BH_Y(%)","RSI_vs_BH_Y(%)",
            "Base_vs_BH_X(%)","RSI_vs_BH_X(%)"] # <--- Tambahkan dua string ini di akhir

df_rq4 = df_all[["Pair","Sektor","Metode","Y","X",
                   "Base_Return","Base_Sharpe","Base_MDD","Base_WinRate","Base_Trades",
                   "RSI_Return","RSI_Sharpe","RSI_MDD","RSI_WinRate","RSI_Trades",
                   "BH_Y","BH_X",
                   "Base_vs_BH_Y","RSI_vs_BH_Y",
                   "Base_vs_BH_X","RSI_vs_BH_X"]].copy() # <--- Tambahkan dua string ini di akhir
df_rq4.columns = cols_rq4
write_df_to_sheet(ws4, df_rq4, title="RQ4 — Kinerja Keseluruhan Out-of-Sample", color_col="RSI_Return(%)")

# Ringkasan overall
ws4.append([])
ws4.append(["Statistik Keseluruhan"])
ws4[ws4.max_row][0].font = Font(name="Arial", bold=True, color="FFFFFF")
ws4[ws4.max_row][0].fill = S["sec_f"]

n = len(df_all)
summary_rows = [
    ["Total Pair",                  n],
    ["Pair Basis Positif (Return)", (df_all["Base_Return"] > 0).sum()],
    ["Pair RSI Positif (Return)",   (df_all["RSI_Return"] > 0).sum()],
    ["Avg Return Basis (%)",        round(df_all["Base_Return"].mean(), 2)],
    ["Avg Return RSI (%)",          round(df_all["RSI_Return"].mean(), 2)],
    ["Median Return Basis (%)",     round(df_all["Base_Return"].median(), 2)],
    ["Median Return RSI (%)",       round(df_all["RSI_Return"].median(), 2)],
    ["Avg Sharpe Basis",            round(df_all["Base_Sharpe"].mean(), 4)],
    ["Avg Sharpe RSI",              round(df_all["RSI_Sharpe"].mean(), 4)],
    ["Avg MDD Basis (%)",           round(df_all["Base_MDD"].mean(), 2)],
    ["Avg MDD RSI (%)",             round(df_all["RSI_MDD"].mean(), 2)],
    ["Avg Win Rate Basis (%)",      round(df_all["Base_WinRate"].mean(), 2)],
    ["Avg Win Rate RSI (%)",        round(df_all["RSI_WinRate"].mean(), 2)],
    ["Avg Trades Basis",            round(df_all["Base_Trades"].mean(), 1)],
    ["Avg Trades RSI",              round(df_all["RSI_Trades"].mean(), 1)],
    ["Avg B&H Y (%)",               round(df_all["BH_Y"].mean(), 2)],
    ["Avg B&H X (%)",               round(df_all["BH_X"].mean(), 2)],
    ["Best Pair (RSI Return)",      df_all.loc[df_all["RSI_Return"].idxmax(), "Pair"]],
    ["Worst Pair (RSI Return)",     df_all.loc[df_all["RSI_Return"].idxmin(), "Pair"]],
]
for r in summary_rows:
    ws4.append(r)
    for cell in ws4[ws4.max_row]:
        cell.font = S["fnt"]
        cell.border = S["bdr"]

# ─── Sheet Parameter Terpilih (dipakai di out-of-sample) ─────────────────────
ws5 = wb.create_sheet("Parameter Terpilih (OOS)")
df_params = df_all[["Pair","Sektor","Metode","Y","X",
                    "Z_Window","z_entry","z_exit","SL","TP","RSI_Period","RSI_Threshold"]].copy()
df_params.columns = ["Pair","Sektor","Metode","Y (Dependen)","X (Independen)",
                     "Z Window","±Z Entry","Z Exit","TL","TP","RSI Period","RSI Threshold"]
write_df_to_sheet(ws5, df_params, title="Parameter Terpilih dari Grid Search — Dipakai di Out-of-Sample")

# ─── Simpan ───────────────────────────────────────────────────────────────────
out_path = "summary_bt.xlsx"
wb.save(out_path)
print(f"\nOutput tersimpan: {out_path}")
print(f"Sheet: Raw Data | RQ1 (Basis+RSI) | RQ2 Per Sektor | RQ3 Basis vs RSI | RQ4 Keseluruhan | Parameter Terpilih (OOS)")