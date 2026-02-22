# Dashboard.py
# Dashboard UTAUT (SmartPLS + SPSS) - FINAL
# Run: python -m streamlit run Dashboard.py

import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Dashboard UTAUT (SmartPLS + SPSS)", layout="wide")

st.markdown(
    """
<style>
.block-container{padding-top:1.1rem;}
.kpi{
  padding:14px 16px;border-radius:14px;
  background:rgba(99,102,241,.06);
  border:1px solid rgba(99,102,241,.18);
}
.muted{color:rgba(0,0,0,.55);font-size:.92rem;}
div[data-testid="stDataFrame"] {border-radius: 14px; overflow: hidden;}
</style>
""",
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_MAIN = BASE_DIR / "Validitas & Realibilitas.xlsx"
DEFAULT_BOOT = BASE_DIR / "Hipotesis.xlsx"
DEFAULT_IMG = BASE_DIR / "Model Smart Pls 250.png"
DEFAULT_OUT = BASE_DIR / "OUTPUT.xls"
DEFAULT_PROFIL = BASE_DIR / "Profil Responden.xlsx"
DEFAULT_Q2 = BASE_DIR / "Q2.xlsx"

CONSTRUCTS = ["BI", "EE", "FC", "PE", "SI"]

# =========================
# HELPERS
# =========================
def fmt(x, nd=3):
    try:
        if x is None:
            return "-"
        x = float(x)
        if np.isnan(x):
            return "-"
        return f"{x:.{nd}f}"
    except Exception:
        return "-"

def safe_df(df: pd.DataFrame) -> pd.DataFrame:
    """Aman untuk st.dataframe / st.table (hindari pyarrow Expected bytes got float)."""
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    out = out.replace({pd.NA: "", np.nan: ""})
    for c in out.columns:
        if out[c].dtype == "object" or str(out[c].dtype).startswith("string"):
            out[c] = out[c].astype("string").fillna("").replace({"nan": "", "None": ""})
    return out

def round_numeric_df(df: pd.DataFrame, nd: int = 3) -> pd.DataFrame:
    """Bulatkan semua kolom numerik jadi nd desimal (untuk tampilan)."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for c in out.columns:
        s = pd.to_numeric(out[c], errors="coerce")
        if s.notna().sum() > 0:
            out[c] = s.round(nd)
    return out

def download_button_df(df: pd.DataFrame, filename: str, label: str, key: str):
    if df is None or df.empty:
        return
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(label, csv, file_name=filename, mime="text/csv", key=key)

@st.cache_data
def read_excel_all_sheets(file_or_path) -> dict:
    """Baca semua sheet dengan header=None (anti geser SmartPLS)."""
    xl = pd.ExcelFile(file_or_path)
    out = {}
    for sh in xl.sheet_names:
        try:
            out[sh] = pd.read_excel(file_or_path, sheet_name=sh, header=None, dtype=object)
        except Exception:
            pass
    return out

def _clean_series_text(s: pd.Series) -> pd.Series:
    s = s.astype(str)
    s = s.mask(s.str.lower().isin(["nan", "none"]), "")
    return s

def guess_report_col(df: pd.DataFrame) -> int:
    """Pilih kolom label (yang paling banyak teks / '->')."""
    best, best_score = None, -1
    for c in df.columns:
        s = _clean_series_text(df[c])
        score = 0
        score += s.str.contains(r"\-\>", regex=True, na=False).sum() * 6
        score += s.str.contains(r"[A-Za-z]", regex=True, na=False).sum() * 2
        score += (s.str.len() > 2).sum()
        if score > best_score:
            best_score = score
            best = c
    return best if best is not None else df.columns[0]

def find_row_contains(df: pd.DataFrame, keyword: str, start=0, look=6000):
    key = str(keyword).lower()
    end = min(len(df), start + look)
    for i in range(start, end):
        row = df.iloc[i].astype(str).fillna("").str.lower().tolist()
        if any(key in cell for cell in row):
            return i
    return None

def _to_num(x):
    return pd.to_numeric(x, errors="coerce")

def format_numeric_cols(df: pd.DataFrame, nd: int = 3) -> pd.DataFrame:
    """Format kolom angka (hipotesis)."""
    if df is None or df.empty:
        return df
    out = df.copy()
    num_cols = [
        "Original sample (O)",
        "Sample mean (M)",
        "Standard deviation (STDEV)",
        "T statistics (|O/STDEV|)",
        "P values",
    ]
    for c in num_cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").map(
                lambda x: "-" if pd.isna(x) else f"{x:.{nd}f}"
            )
    return out

def clean_regression_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Keep hanya baris variabel regresi yang valid: (Constant), PE, EE, SI, FC."""
    if df is None or df.empty or "Variabel" not in df.columns:
        return df
    out = df.copy()
    out["Variabel"] = out["Variabel"].astype(str).str.strip()
    allowed = {"(Constant)", "Constant", "PE", "EE", "SI", "FC"}
    out = out[out["Variabel"].isin(allowed)].copy()
    out["Variabel"] = out["Variabel"].replace({"Constant": "(Constant)"})
    order = {"(Constant)": 0, "PE": 1, "EE": 2, "SI": 3, "FC": 4}
    out["__ord"] = out["Variabel"].map(lambda v: order.get(v, 99))
    out = out.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)
    return out

# ====== PROFIL RESPONDEN HELPERS (profil: header normal, bukan header=None) ======
def _norm_col(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def load_profil_responden(file_prof) -> pd.DataFrame | None:
    if file_prof is None:
        return None
    if isinstance(file_prof, str):
        if str(file_prof).strip() == "":
            return None
        p = Path(file_prof)
        if not p.exists():
            return None
        df = pd.read_excel(p)  # header=0
    else:
        df = pd.read_excel(file_prof)  # header=0

    df = df.dropna(axis=1, how="all")

    rename_map = {}
    for c in df.columns:
        nc = _norm_col(c)
        if nc in ["gender", "jenis kelamin", "jk", "kelamin"]:
            rename_map[c] = "Jenis Kelamin"
        elif nc in ["usia", "umur", "age"]:
            rename_map[c] = "Usia"
        elif nc in ["pendidikan", "pendidikan terakhir", "education"]:
            rename_map[c] = "Pendidikan"
        elif nc in ["pekerjaan", "job", "occupation"]:
            rename_map[c] = "Pekerjaan"
        elif nc in ["domisili", "kota", "asal", "alamat", "residence"]:
            rename_map[c] = "Domisili"
        elif nc in ["penghasilan", "income", "gaji"]:
            rename_map[c] = "Penghasilan"
        elif nc in ["pengeluaran", "spending"]:
            rename_map[c] = "Pengeluaran"
        elif nc in ["lama penggunaan", "durasi penggunaan", "how long", "lama pakai"]:
            rename_map[c] = "Lama Penggunaan"

    if rename_map:
        df = df.rename(columns=rename_map)

    return safe_df(df)

# =========================
# SMARTPLS - MATRIX PARSERS
# =========================
def _find_header_row_for_matrix(df: pd.DataFrame, start: int, labels: list, max_look=250):
    labels_u = [x.upper() for x in labels]
    for i in range(start, min(start + max_look, len(df))):
        row = df.iloc[i].astype(str).fillna("").str.upper().str.strip()
        hits = sum([(row == lab).any() for lab in labels_u])
        if hits >= 2:
            return i
    return None

def _map_cols_by_header(df: pd.DataFrame, header_row: int, report_col: int, labels: list):
    want = set([x.upper() for x in labels])
    m = {}
    for c in df.columns:
        if c == report_col:
            continue
        h = str(df.at[header_row, c]).strip().upper()
        if h in want:
            m[h] = c
    return m

def extract_outer_loadings_matrix(df: pd.DataFrame, report_col: int) -> pd.DataFrame:
    start = find_row_contains(df, "outer loadings")
    if start is None:
        return pd.DataFrame()

    header = _find_header_row_for_matrix(df, start, CONSTRUCTS)
    if header is None:
        return pd.DataFrame()

    col_map = _map_cols_by_header(df, header, report_col, CONSTRUCTS)
    if len(col_map) < 2:
        return pd.DataFrame()

    rows = []
    for i in range(header + 1, min(header + 1600, len(df))):
        ind = str(df.at[i, report_col]).strip()
        if ind == "" or ind.lower() in ["nan", "none"]:
            continue

        low = ind.lower()
        if any(k in low for k in [
            "construct reliability",
            "discriminant validity",
            "fornell",
            "r-square",
            "path coefficients",
            "f-square",
            "cross-validated",
        ]):
            break

        vals = []
        num_count = 0
        for k in CONSTRUCTS:
            c = col_map.get(k)
            v = _to_num(df.at[i, c]) if c is not None else np.nan
            if pd.notna(v):
                num_count += 1
            vals.append(v)

        if num_count >= 1:
            rows.append([ind] + vals)

    out = pd.DataFrame(rows, columns=["Indikator"] + CONSTRUCTS)
    out = out[~out["Indikator"].str.contains("outer loadings", case=False, na=False)]
    return out.reset_index(drop=True)

def extract_reliability_validity_overview(df: pd.DataFrame, report_col: int) -> pd.DataFrame:
    start = find_row_contains(df, "construct reliability and validity")
    if start is None:
        return pd.DataFrame()

    header = None
    for i in range(start, min(start + 500, len(df))):
        row = df.iloc[i].astype(str).fillna("").str.lower()
        if row.str.contains("cronbach", na=False).any() and (
            row.str.contains("average variance extracted", na=False).any()
            or row.str.contains("ave", na=False).any()
        ):
            header = i
            break
    if header is None:
        return pd.DataFrame()

    col_ca = col_rhoa = col_rhoc = col_ave = None
    for c in df.columns:
        if c == report_col:
            continue
        h = str(df.at[header, c]).strip().lower()
        if "cronbach" in h:
            col_ca = c
        elif "rho_a" in h or "rho a" in h:
            col_rhoa = c
        elif "rho_c" in h or "rho c" in h:
            col_rhoc = c
        elif "average variance extracted" in h or h == "ave":
            col_ave = c

    if col_ca is None and col_rhoa is None and col_rhoc is None and col_ave is None:
        return pd.DataFrame()

    rows = []
    seen = set()

    for r in range(header + 1, min(header + 400, len(df))):
        k = str(df.at[r, report_col]).strip().upper()
        if k == "" or k.lower() in ["nan", "none"]:
            continue

        low = k.lower()
        if any(x in low for x in [
            "outer loadings",
            "discriminant",
            "fornell",
            "r-square",
            "path",
            "f-square",
            "cross-validated",
        ]):
            break

        if k in CONSTRUCTS and k not in seen:
            rows.append([
                k,
                _to_num(df.at[r, col_ca]) if col_ca is not None else np.nan,
                _to_num(df.at[r, col_rhoa]) if col_rhoa is not None else np.nan,
                _to_num(df.at[r, col_rhoc]) if col_rhoc is not None else np.nan,
                _to_num(df.at[r, col_ave]) if col_ave is not None else np.nan,
            ])
            seen.add(k)

        if len(seen) == len(CONSTRUCTS):
            break

    out = pd.DataFrame(rows, columns=["Konstruk", "Cronbach's alpha", "rho_a", "rho_c", "AVE"])
    if out.empty:
        return out
    out["Konstruk"] = pd.Categorical(out["Konstruk"], categories=CONSTRUCTS, ordered=True)
    out = out.sort_values("Konstruk").reset_index(drop=True)
    return out

def extract_fornell_larcker_matrix(df: pd.DataFrame, report_col: int) -> pd.DataFrame:
    start = find_row_contains(df, "fornell-larcker")
    if start is None:
        start = find_row_contains(df, "fornell")
    if start is None:
        return pd.DataFrame()

    header = _find_header_row_for_matrix(df, start, CONSTRUCTS)
    if header is None:
        return pd.DataFrame()

    col_map = _map_cols_by_header(df, header, report_col, CONSTRUCTS)
    if len(col_map) < 2:
        return pd.DataFrame()

    rows = []
    seen = set()
    for i in range(header + 1, min(header + 200, len(df))):
        rname = str(df.at[i, report_col]).strip().upper()
        if rname == "" or rname.lower() in ["nan", "none"]:
            continue

        low = rname.lower()
        if any(x in low for x in ["r-square", "path", "f-square", "cross-validated", "construct reliability"]):
            break

        if rname in CONSTRUCTS and rname not in seen:
            vals = []
            for k in CONSTRUCTS:
                c = col_map.get(k)
                vals.append(_to_num(df.at[i, c]) if c is not None else np.nan)
            rows.append([rname] + vals)
            seen.add(rname)

        if len(seen) == len(CONSTRUCTS):
            break

    out = pd.DataFrame(rows, columns=["Konstruk"] + CONSTRUCTS)
    if out.empty:
        return out
    out["Konstruk"] = pd.Categorical(out["Konstruk"], categories=CONSTRUCTS, ordered=True)
    out = out.sort_values("Konstruk").reset_index(drop=True)
    return out

def extract_rsquare_overview(df: pd.DataFrame, report_col: int) -> dict:
    start = find_row_contains(df, "r-square")
    if start is None:
        start = find_row_contains(df, "r square")
    if start is None:
        return {"R2": np.nan, "AdjR2": np.nan}

    header = None
    for i in range(start, min(start + 400, len(df))):
        row = df.iloc[i].astype(str).fillna("").str.lower()
        if row.str.contains("r-square adjusted", na=False).any() or row.str.contains("adjusted", na=False).any():
            header = i
            break
    if header is None:
        return {"R2": np.nan, "AdjR2": np.nan}

    col_r2 = col_adj = None
    for c in df.columns:
        if c == report_col:
            continue
        h = str(df.at[header, c]).strip().lower()
        if ("r-square" in h or "r square" in h) and "adjust" not in h:
            col_r2 = c
        if "adjust" in h and ("r" in h):
            col_adj = c

    bi_row = None
    for i in range(header + 1, min(header + 150, len(df))):
        if str(df.at[i, report_col]).strip().upper() == "BI":
            bi_row = i
            break
    if bi_row is None:
        return {"R2": np.nan, "AdjR2": np.nan}

    r2 = _to_num(df.at[bi_row, col_r2]) if col_r2 is not None else np.nan
    adj = _to_num(df.at[bi_row, col_adj]) if col_adj is not None else np.nan
    return {"R2": float(r2) if pd.notna(r2) else np.nan, "AdjR2": float(adj) if pd.notna(adj) else np.nan}

def extract_fsquare_matrix(df: pd.DataFrame, report_col: int) -> pd.DataFrame:
    start = find_row_contains(df, "f-square")
    if start is None:
        start = find_row_contains(df, "f square")
    if start is None:
        return pd.DataFrame()

    header = _find_header_row_for_matrix(df, start, CONSTRUCTS)
    if header is None:
        return pd.DataFrame()

    col_map = _map_cols_by_header(df, header, report_col, CONSTRUCTS)
    col_bi = col_map.get("BI")
    if col_bi is None:
        return pd.DataFrame()

    rows = []
    seen = set()
    for i in range(header + 1, min(header + 250, len(df))):
        rname = str(df.at[i, report_col]).strip().upper()
        if rname == "" or rname.lower() in ["nan", "none"]:
            continue

        low = rname.lower()
        if any(x in low for x in ["cross-validated", "q2", "construct reliability", "discriminant", "fornell", "path coefficients"]):
            break

        if rname in CONSTRUCTS and rname != "BI" and rname not in seen:
            val = _to_num(df.at[i, col_bi])
            if pd.notna(val):
                rows.append([f"{rname} → BI", float(val)])
                seen.add(rname)

        if len(seen) == (len(CONSTRUCTS) - 1):
            break

    return pd.DataFrame(rows, columns=["Path", "f²"]).reset_index(drop=True)

def extract_q2_redundancy(df: pd.DataFrame, report_col: int) -> pd.DataFrame:
    start = find_row_contains(df, "cross-validated redundancy")
    if start is None:
        start = find_row_contains(df, "q² (=1-sse/sso)")
    if start is None:
        start = find_row_contains(df, "q2 (=1-sse/sso)")
    if start is None:
        return pd.DataFrame()

    header = None
    for i in range(start, min(start + 450, len(df))):
        row = df.iloc[i].astype(str).fillna("").str.lower()
        if row.str.contains("sso", na=False).any() and row.str.contains("sse", na=False).any():
            header = i
            break
    if header is None:
        return pd.DataFrame()

    col_sso = col_sse = col_q2 = None
    for c in df.columns:
        if c == report_col:
            continue
        h = str(df.at[header, c]).strip().lower()
        if h == "sso":
            col_sso = c
        elif h == "sse":
            col_sse = c
        elif "q" in h:
            col_q2 = c

    if col_q2 is None:
        return pd.DataFrame()

    rows = []
    seen = set()
    for i in range(header + 1, min(header + 200, len(df))):
        k = str(df.at[i, report_col]).strip().upper()
        if k == "" or k.lower() in ["nan", "none"]:
            continue

        if k in CONSTRUCTS and k not in seen:
            rows.append([
                k,
                _to_num(df.at[i, col_sso]) if col_sso is not None else np.nan,
                _to_num(df.at[i, col_sse]) if col_sse is not None else np.nan,
                _to_num(df.at[i, col_q2]),
            ])
            seen.add(k)

        if len(seen) == len(CONSTRUCTS):
            break

    out = pd.DataFrame(rows, columns=["Konstruk", "SSO", "SSE", "Q²"]).reset_index(drop=True)
    if out.empty:
        return out
    out["Konstruk"] = pd.Categorical(out["Konstruk"], categories=CONSTRUCTS, ordered=True)
    out = out.sort_values("Konstruk").reset_index(drop=True)
    return out

# =========================
# HIPOTESIS (Bootstrap) - DIKUNCI SESUAI FILE KAMU
# =========================
def extract_hypothesis_table_bootstrap_complete(df_boot: pd.DataFrame) -> pd.DataFrame:
    if df_boot is None or df_boot.empty:
        return pd.DataFrame()

    report_col = guess_report_col(df_boot)
    start = find_row_contains(df_boot, "Mean, STDEV, T values, p values")
    if start is None:
        start = find_row_contains(df_boot, "mean, stdev, t values, p values")
    if start is None:
        start = find_row_contains(df_boot, "Path coefficients")
    if start is None:
        return pd.DataFrame()

    header = None
    for i in range(start, min(start + 40, len(df_boot))):
        row = df_boot.iloc[i].astype(str)
        row = row.mask(row.str.lower().isin(["nan", "none"]), "")
        rowl = row.str.lower()
        ok = (
            rowl.str.contains("original sample", na=False).any()
            and rowl.str.contains("sample mean", na=False).any()
            and rowl.str.contains("standard deviation", na=False).any()
            and rowl.str.contains("t statistics", na=False).any()
            and rowl.str.contains("p values", na=False).any()
        )
        if ok:
            header = i
            break
    if header is None:
        return pd.DataFrame()

    def find_col_contains(text):
        t = text.lower()
        for c in df_boot.columns:
            h = str(df_boot.at[header, c]).lower()
            if t in h:
                return c
        return None

    col_O = find_col_contains("original sample")
    col_M = find_col_contains("sample mean")
    col_SD = find_col_contains("standard deviation")
    col_T = find_col_contains("t statistics")
    col_P = find_col_contains("p values")
    if col_O is None or col_P is None:
        return pd.DataFrame()

    targets = {"EE -> BI", "FC -> BI", "PE -> BI", "SI -> BI"}
    rows = []
    for i in range(header + 1, min(header + 120, len(df_boot))):
        lab = str(df_boot.at[i, report_col]).strip()
        if lab.lower() in ["", "nan", "none"]:
            continue
        lab_norm = re.sub(r"\s+", " ", lab.replace("→", "->")).strip()

        if lab_norm in targets:
            O = pd.to_numeric(df_boot.at[i, col_O], errors="coerce")
            M = pd.to_numeric(df_boot.at[i, col_M], errors="coerce") if col_M is not None else np.nan
            SD = pd.to_numeric(df_boot.at[i, col_SD], errors="coerce") if col_SD is not None else np.nan
            T = pd.to_numeric(df_boot.at[i, col_T], errors="coerce") if col_T is not None else np.nan
            P = pd.to_numeric(df_boot.at[i, col_P], errors="coerce") if col_P is not None else np.nan
            rows.append([lab_norm.replace("->", "→"), O, M, SD, T, P])

        if len(rows) == 4:
            break

    out = pd.DataFrame(
        rows,
        columns=[
            "Path",
            "Original sample (O)",
            "Sample mean (M)",
            "Standard deviation (STDEV)",
            "T statistics (|O/STDEV|)",
            "P values",
        ],
    )
    if out.empty:
        return out

    order = ["EE → BI", "FC → BI", "PE → BI", "SI → BI"]
    out["__ord"] = out["Path"].apply(lambda x: order.index(x) if x in order else 99)
    out = out.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)
    out["Keputusan (α=0.05)"] = np.where(out["P values"] < 0.05, "Diterima", "Ditolak")
    return out

# =========================
# SPSS OUTPUT.xls PARSER
# =========================
def load_spss_output_xls(path: str) -> dict:
    xls = pd.ExcelFile(path, engine="xlrd")  # butuh xlrd==2.0.1
    frames = {}
    for sh in xls.sheet_names:
        frames[sh] = pd.read_excel(xls, sheet_name=sh, header=None, dtype=object)
    return frames

def _sheet_find(df: pd.DataFrame, keyword: str):
    key = keyword.lower()
    for i in range(len(df)):
        row = df.iloc[i].astype(str).fillna("").str.lower()
        if row.str.contains(re.escape(key), na=False).any():
            return i
    return None

def _col_by_any_header(df, header_rows, want_keywords):
    ncol = df.shape[1]
    merged = []
    for c in range(ncol):
        parts = []
        for r in header_rows:
            if 0 <= r < len(df):
                parts.append(str(df.iat[r, c]).strip().lower())
        merged.append(" ".join([p for p in parts if p and p != "nan"]))

    for c, text in enumerate(merged):
        ok = True
        for k in want_keywords:
            if k not in text:
                ok = False
                break
        if ok:
            return c
    return None

def extract_spss_model_summary(frames: dict) -> dict:
    for sh, df in frames.items():
        pos = _sheet_find(df, "Model Summary")
        if pos is None:
            continue
        for h in range(pos, min(pos + 40, len(df))):
            r2_col = _col_by_any_header(df, [h, h + 1, h + 2], ["r square"])
            adj_col = _col_by_any_header(df, [h, h + 1, h + 2], ["adjusted", "r square"])
            if r2_col is None:
                continue
            for r in range(h + 1, min(h + 15, len(df))):
                r2 = pd.to_numeric(df.iat[r, r2_col], errors="coerce")
                adj = pd.to_numeric(df.iat[r, adj_col], errors="coerce") if adj_col is not None else np.nan
                if pd.notna(r2):
                    return {"R2": float(r2), "AdjR2": float(adj) if pd.notna(adj) else np.nan}
    return {"R2": np.nan, "AdjR2": np.nan}

def extract_spss_anova(frames: dict) -> dict:
    for sh, df in frames.items():
        pos = _sheet_find(df, "ANOVA")
        if pos is None:
            continue
        for h in range(pos, min(pos + 50, len(df))):
            f_col = _col_by_any_header(df, [h, h + 1, h + 2], ["f"])
            sig_col = _col_by_any_header(df, [h, h + 1, h + 2], ["sig"])
            if f_col is None:
                continue
            for r in range(h + 1, min(h + 30, len(df))):
                label = str(df.iat[r, 0]).strip().lower()
                if "regression" in label:
                    F = pd.to_numeric(df.iat[r, f_col], errors="coerce")
                    SigF = pd.to_numeric(df.iat[r, sig_col], errors="coerce") if sig_col is not None else np.nan
                    if pd.notna(F):
                        return {"F": float(F), "SigF": float(SigF) if pd.notna(SigF) else np.nan}
    return {"F": np.nan, "SigF": np.nan}

def extract_spss_coefficients_and_vif(frames: dict):
    for sh, df in frames.items():
        pos = _sheet_find(df, "Coefficients")
        if pos is None:
            continue

        for h in range(pos, min(pos + 60, len(df))):
            b_col = _col_by_any_header(df, [h, h + 1, h + 2], ["b"])
            beta_col = _col_by_any_header(df, [h, h + 1, h + 2], ["beta"])
            t_col = _col_by_any_header(df, [h, h + 1, h + 2], ["t"])
            sig_col = _col_by_any_header(df, [h, h + 1, h + 2], ["sig"])
            vif_col = _col_by_any_header(df, [h, h + 1, h + 2], ["vif"])
            tol_col = _col_by_any_header(df, [h, h + 1, h + 2], ["tolerance"])

            if t_col is None or sig_col is None:
                continue

            var_col = 1 if df.shape[1] > 1 else 0
            rows = []
            vif_rows = []
            for r in range(h + 1, min(h + 160, len(df))):
                varname = str(df.iat[r, var_col]).strip()
                if varname == "" or varname.lower() == "nan":
                    continue
                if "model summary" in varname.lower() or "anova" in varname.lower():
                    break

                B = pd.to_numeric(df.iat[r, b_col], errors="coerce") if b_col is not None else np.nan
                Beta = pd.to_numeric(df.iat[r, beta_col], errors="coerce") if beta_col is not None else np.nan
                tval = pd.to_numeric(df.iat[r, t_col], errors="coerce") if t_col is not None else np.nan
                Sig = pd.to_numeric(df.iat[r, sig_col], errors="coerce") if sig_col is not None else np.nan
                VIF = pd.to_numeric(df.iat[r, vif_col], errors="coerce") if vif_col is not None else np.nan
                Tol = pd.to_numeric(df.iat[r, tol_col], errors="coerce") if tol_col is not None else np.nan

                if pd.isna(tval) and pd.isna(Sig) and pd.isna(B) and pd.isna(Beta):
                    continue

                rows.append([varname, B, Beta, tval, Sig, Tol, VIF])
                if pd.notna(VIF) and "constant" not in varname.lower():
                    vif_rows.append([varname, VIF])

            coef_df = pd.DataFrame(rows, columns=["Variabel", "B", "Beta", "t", "Sig", "Tolerance", "VIF"])
            vif_df = pd.DataFrame(vif_rows, columns=["Variabel", "VIF"])
            if not coef_df.empty:
                return coef_df, vif_df

    return pd.DataFrame(), pd.DataFrame()

# =========================
# SIDEBAR
# =========================
st.sidebar.title("📌 Input Dashboard")
n_resp = st.sidebar.number_input("Jumlah responden", min_value=1, max_value=2000, value=250, step=1)

mode = st.sidebar.radio("Sumber file", ["Pakai file di folder ini (default)", "Upload file"], key="mode_file")

if mode == "Pakai file di folder ini (default)":
    file_main = st.sidebar.text_input("File SmartPLS (Validitas & Realibilitas)", value=str(DEFAULT_MAIN))
    file_boot = st.sidebar.text_input("File Hipotesis (Bootstrap)", value=str(DEFAULT_BOOT))
    file_img = st.sidebar.text_input("Gambar model SmartPLS (opsional)", value=str(DEFAULT_IMG))
    file_out = st.sidebar.text_input("File OUTPUT (SPSS)", value=str(DEFAULT_OUT))
    file_prof = st.sidebar.text_input("File Profil Responden (opsional)", value=str(DEFAULT_PROFIL))
    file_q2 = st.sidebar.text_input("File Q² (opsional)", value=str(DEFAULT_Q2))
else:
    file_main = st.sidebar.file_uploader("Upload SmartPLS (XLSX/XLS/CSV)", type=["xlsx", "xls", "csv"])
    file_boot = st.sidebar.file_uploader("Upload Hipotesis (XLSX/XLS/CSV)", type=["xlsx", "xls", "csv"])
    file_img = st.sidebar.file_uploader("Upload gambar model (opsional)", type=["png", "jpg", "jpeg"])
    file_out = st.sidebar.file_uploader("Upload OUTPUT SPSS (.xls)", type=["xls"])
    file_prof = st.sidebar.file_uploader("Upload Profil Responden (opsional) (xlsx/xls)", type=["xlsx", "xls"])
    file_q2 = st.sidebar.file_uploader("Upload Q² (opsional) (xlsx/xls)", type=["xlsx", "xls"])

    if file_main is None:
        st.info("Upload file SmartPLS dulu ya.")
        st.stop()

# =========================
# LOAD SMARTPLS
# =========================
report = None
outer_matrix_df = rel_overview_df = fornell_df = pd.DataFrame()
r2_smart = {"R2": np.nan, "AdjR2": np.nan}
f2_df = pd.DataFrame()
q2_df = pd.DataFrame()

try:
    sheets = read_excel_all_sheets(file_main)
    if sheets:
        report = max(sheets.values(), key=lambda d: d.shape[0] * d.shape[1])
        report_col = guess_report_col(report)

        outer_matrix_df = extract_outer_loadings_matrix(report, report_col)
        rel_overview_df = extract_reliability_validity_overview(report, report_col)
        fornell_df = extract_fornell_larcker_matrix(report, report_col)
        r2_smart = extract_rsquare_overview(report, report_col)
        f2_df = extract_fsquare_matrix(report, report_col)
        q2_df = extract_q2_redundancy(report, report_col)
except Exception as e:
    st.error(f"Gagal baca file SmartPLS: {e}")

# =========================
# LOAD Q2.xlsx (opsional) -> override q2_df kalau kebaca
# =========================
try:
    if file_q2 is not None and str(file_q2).strip() != "":
        q2_sheets = read_excel_all_sheets(file_q2)
        if q2_sheets:
            q2_report = max(q2_sheets.values(), key=lambda d: d.shape[0] * d.shape[1])
            q2_col = guess_report_col(q2_report)
            q2_from_file = extract_q2_redundancy(q2_report, q2_col)
            if q2_from_file is not None and not q2_from_file.empty:
                q2_df = q2_from_file
except Exception as e:
    st.warning(f"Gagal baca Q2.xlsx: {e}")

# =========================
# LOAD HIPOTESIS (BOOTSTRAP)
# =========================
hyp_df = pd.DataFrame()
boot_df = None

try:
    if file_boot is not None and str(file_boot).strip() != "":
        boot_sheets = read_excel_all_sheets(file_boot)
        if "complete" in boot_sheets:
            boot_df = boot_sheets["complete"]
        else:
            boot_df = max(boot_sheets.values(), key=lambda d: d.shape[0] * d.shape[1]) if boot_sheets else None

        if boot_df is not None:
            hyp_df = extract_hypothesis_table_bootstrap_complete(boot_df)
except Exception as e:
    st.warning(f"Gagal parse Hipotesis: {e}")

# =========================
# LOAD SPSS (OUTPUT.xls)
# =========================
spss_ms = {"R2": np.nan, "AdjR2": np.nan}
spss_an = {"F": np.nan, "SigF": np.nan}
spss_coef = pd.DataFrame()
spss_vif = pd.DataFrame()

try:
    if file_out is not None and str(file_out).strip() != "":
        frames = load_spss_output_xls(str(file_out))
        spss_ms = extract_spss_model_summary(frames)
        spss_an = extract_spss_anova(frames)
        spss_coef, spss_vif = extract_spss_coefficients_and_vif(frames)
except Exception as e:
    st.error(f"Gagal membaca file OUTPUT.xls: {e}")

# =========================
# LOAD PROFIL (opsional)
# =========================
profil_df = None
try:
    if file_prof is not None and str(file_prof).strip() != "":
        profil_df = load_profil_responden(file_prof)
except Exception:
    profil_df = None

# =========================
# TABS
# =========================
tab1, tabP, tab2, tab3, tabD = st.tabs(
    ["Overview", "Profil Responden", "SEM-PLS (SmartPLS)", "Regresi (SPSS OUTPUT)", "Debug"]
)

# =========================
# TAB 1 - OVERVIEW
# =========================
with tab1:
    st.markdown("## Overview")

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(
        f"<div class='kpi'><b>Responden</b><br><span style='font-size:28px'>{int(n_resp)}</span></div>",
        unsafe_allow_html=True,
    )
    c2.markdown(
        f"<div class='kpi'><b>R² BI (SmartPLS)</b><br><span style='font-size:28px'>{fmt(r2_smart.get('R2'))}</span></div>",
        unsafe_allow_html=True,
    )
    c3.markdown(
        f"<div class='kpi'><b>Adj R² BI (SmartPLS)</b><br><span style='font-size:28px'>{fmt(r2_smart.get('AdjR2'))}</span></div>",
        unsafe_allow_html=True,
    )
    c4.markdown(
        f"<div class='kpi'><b>R² BI (SPSS)</b><br><span style='font-size:28px'>{fmt(spss_ms.get('R2'))}</span></div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.markdown("### Hasil Hipotesis (Bootstrap SmartPLS)")
    st.markdown(
        "<div class='muted'>Path coefficients (O, M, STDEV, T, P) — diambil dari Hipotesis.xlsx.</div>",
        unsafe_allow_html=True,
    )

    if hyp_df is None or hyp_df.empty:
        st.warning("Tabel hipotesis belum kebaca. Pastikan Hipotesis.xlsx berisi tabel path coefficients.")
    else:
        st.table(format_numeric_cols(hyp_df, 3))
        download_button_df(hyp_df, "hipotesis_bootstrap.csv", "⬇️ Download Hipotesis", key="dl_hyp_overview")

# =========================
# TAB PROFIL
# =========================
with tabP:
    st.markdown("## Profil Responden")
    st.markdown(
        "<div class='muted'>Data diambil dari <b>Profil Responden.xlsx</b> (opsional).</div>",
        unsafe_allow_html=True,
    )
    if profil_df is None or (isinstance(profil_df, pd.DataFrame) and profil_df.empty):
        st.info("Belum ada data profil responden.")
    else:
        st.dataframe(safe_df(profil_df), use_container_width=True, hide_index=True)
        download_button_df(safe_df(profil_df), "profil_responden.csv", "⬇️ Download Profil Responden", key="dl_profil")

# =========================
# TAB 2 - SEM-PLS
# =========================
with tab2:
    st.markdown("## SEM-PLS (SmartPLS)")
    A, B = st.tabs(["A) Outer Model", "B) Inner Model"])

    with A:
        st.markdown("### Gambar Model SmartPLS (opsional)")
        if mode == "Pakai file di folder ini (default)":
            p = Path(str(file_img))
            if p.exists():
                st.image(str(p), use_container_width=True)
            else:
                st.info("Gambar model tidak ditemukan (opsional).")
        else:
            if file_img is not None:
                st.image(file_img, use_container_width=True)
            else:
                st.info("Belum upload gambar (opsional).")

        st.markdown("### Outer loadings - Matrix")
        if outer_matrix_df is None or outer_matrix_df.empty:
            st.warning("Outer loadings belum kebaca.")
        else:
            st.dataframe(safe_df(round_numeric_df(outer_matrix_df, 3)), use_container_width=True, hide_index=True)

        st.markdown("### Construct reliability and validity - Overview")
        if rel_overview_df is None or rel_overview_df.empty:
            st.warning("Reliability & validity belum kebaca.")
        else:
            st.dataframe(safe_df(round_numeric_df(rel_overview_df, 3)), use_container_width=True, hide_index=True)

        st.markdown("### Discriminant validity - Fornell-Larcker criterion")
        if fornell_df is None or fornell_df.empty:
            st.warning("Fornell-Larcker belum kebaca.")
        else:
            st.dataframe(safe_df(round_numeric_df(fornell_df, 3)), use_container_width=True, hide_index=True)

    with B:
        st.markdown("### R-square - Overview")
        c1, c2 = st.columns(2)
        c1.metric("R-square (BI)", fmt(r2_smart.get("R2")))
        c2.metric("R-square adjusted (BI)", fmt(r2_smart.get("AdjR2")))

        st.markdown("### f-square - Matrix")
        if f2_df is None or f2_df.empty:
            st.warning("f-square belum kebaca.")
        else:
            st.dataframe(safe_df(round_numeric_df(f2_df, 3)), use_container_width=True, hide_index=True)

        st.markdown("### Construct cross-validated redundancy (Q²)")
        if q2_df is None or q2_df.empty:
            st.warning("Q² belum kebaca. (Coba isi Q2.xlsx atau pastiin SmartPLS export ada bagian Q²)")
        else:
            st.dataframe(safe_df(round_numeric_df(q2_df, 3)), use_container_width=True, hide_index=True)

# =========================
# TAB 3 - REGRESI (SPSS OUTPUT)
# =========================
with tab3:
    st.markdown("## Regresi (SPSS OUTPUT)")
    st.markdown(
        "<div class='muted'>Diambil dari OUTPUT.xls SPSS: R²/Adj R², F & Sig(F), Coefficients, Tolerance & VIF.</div>",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("R²", fmt(spss_ms.get("R2")))
    c2.metric("Adj R²", fmt(spss_ms.get("AdjR2")))
    c3.metric("F", fmt(spss_an.get("F")))
    c4.metric("Sig(F)", fmt(spss_an.get("SigF"), nd=4))

    st.markdown("### Uji Asumsi (VIF)")
    vif_show = clean_regression_rows(spss_vif)
    if vif_show is not None and not vif_show.empty:
        vif_show = vif_show[vif_show["Variabel"] != "(Constant)"].reset_index(drop=True)

    if vif_show is None or vif_show.empty:
        st.warning("VIF belum kebaca. Pastikan OUTPUT.xls ada tabel Coefficients yang memuat VIF.")
    else:
        vif_show = vif_show.copy()
        if "VIF" in vif_show.columns:
            vif_show["VIF"] = pd.to_numeric(vif_show["VIF"], errors="coerce").map(
                lambda x: "-" if pd.isna(x) else f"{x:.3f}"
            )
        st.table(vif_show[["Variabel", "VIF"]])

    st.markdown("### Coefficients")
    coef_show = clean_regression_rows(spss_coef)

    if coef_show is None or coef_show.empty:
        st.warning("Tabel Coefficients belum kebaca dari OUTPUT.xls.")
        st.info("Cek: OUTPUT.xls benar-benar hasil REGRESI (ada tabel 'Coefficients').")
    else:
        coef_show = coef_show.copy()
        want = ["Variabel", "B", "Beta", "t", "Sig", "Tolerance", "VIF"]
        want = [c for c in want if c in coef_show.columns]
        coef_show = coef_show[want]

        for c in ["B", "Beta", "t", "Tolerance", "VIF"]:
            if c in coef_show.columns:
                coef_show[c] = pd.to_numeric(coef_show[c], errors="coerce").map(
                    lambda x: "-" if pd.isna(x) else f"{x:.3f}"
                )
        if "Sig" in coef_show.columns:
            coef_show["Sig"] = pd.to_numeric(coef_show["Sig"], errors="coerce").map(
                lambda x: "-" if pd.isna(x) else f"{x:.3f}"
            )

        order = {"(Constant)": 0, "PE": 1, "EE": 2, "SI": 3, "FC": 4}
        coef_show["__ord"] = coef_show["Variabel"].map(lambda v: order.get(str(v).strip(), 99))
        coef_show = coef_show.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)

        st.table(coef_show)

# =========================
# TAB DEBUG
# =========================
with tabD:
    st.markdown("## Debug")
    st.markdown("<div class='muted'>Cek data mentah kalau parsing ada yang miss.</div>", unsafe_allow_html=True)

    st.markdown("### SmartPLS (raw head)")
    if report is None:
        st.info("SmartPLS belum kebaca.")
    else:
        st.write("report shape:", report.shape)
        st.dataframe(report.head(120), use_container_width=True)

    st.markdown("### Hipotesis.xlsx (raw head)")
    if boot_df is None:
        st.info("Boot file belum kebaca.")
    else:
        st.write("boot_df shape:", boot_df.shape)
        st.dataframe(boot_df.head(120), use_container_width=True)

