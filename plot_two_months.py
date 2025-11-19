# final_plot_aug_sep.py
"""
Script final:
- Baca 2 file Excel (Agustus & September)
- Exclude baris biaya (PENGELUARAN, OPERASIONAL, BELANJA, TOTAL, dll.)
- Otomatis detect apakah kolom numeric adalah revenue (Rp) atau quantity (unit)
- Jika revenue & ada sheet price list, hitung qty = revenue / price
- Plot 2 chart vertikal (August top, September bottom)
- Save CSV per bulan dan PNG gabungan
- Tampilkan contoh perhitungan
"""

import os, re, sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# ------------------ UBAH PATH JIKA PERLU ------------------
AUG_PATH = "Agustusss Recap.xlsx"
SEP_PATH = "Copy of September Recap new.xlsx"

# kemungkinan nama sheet price list
POSSIBLE_PRICE_SHEETS = ["HARGA RATA-RATA MDL", "HARGA RATA-RATA", "HARGA RATA", "PRICE LIST", "HARGA"]

# kata2 biaya yang akan DIHAPUS (case-insensitive substring)
BLOCKWORDS = [
    "PENGELUARAN", "OPERASIONAL", "OPRASIONAL", "BELANJA", "BELANJA BAR",
    "TOTAL", "GRAND TOTAL", "GAJI", "MODAL", "BIAYA", "KELUAR", "PEMBAYARAN",
    "PENGAMBILAN", "HUTANG", "KREDIT"
]

# ---------- helper ----------
def normalize_name(s):
    if pd.isna(s):
        return ""
    s = str(s).upper().strip()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s

def find_header_row(raw_df, keywords=("ITEM","DESCRIPTION","NO")):
    max_check = min(20, raw_df.shape[0])
    for i in range(max_check):
        row = raw_df.iloc[i].astype(str).str.upper().fillna("")
        for kw in keywords:
            if row.str.contains(kw).any():
                return i
    return 3

def detect_price_sheet(path):
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        for s in xls.sheet_names:
            up = s.upper()
            for cand in POSSIBLE_PRICE_SHEETS:
                if cand in up:
                    return s
    except Exception:
        return None
    return None

def read_price_map(path):
    """Return dict {Item_norm: UnitPrice} or {} if not found.
       Robust: handle multi-column matches and pick best single column.
    """
    sheet = detect_price_sheet(path)
    if not sheet:
        return {}
    try:
        dfp = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except Exception:
        return {}

    # --- detect candidate name & price columns ---
    name_col = None
    price_col = None

    # 1) first try columns whose header text contains keywords
    for c in dfp.columns:
        c_up = str(c).upper()
        if name_col is None and any(k in c_up for k in ("ITEM","NAME","DESCRIPTION")):
            name_col = c
        if price_col is None and any(k in c_up for k in ("HARGA","PRICE","RP")):
            price_col = c
        if name_col is not None and price_col is not None:
            break

    # 2) fallback heuristics: first non-numeric-like col for name, first numeric-like col for price
    if name_col is None or price_col is None:
        # identify numeric columns
        numeric_cols = [c for c in dfp.columns if np.issubdtype(dfp[c].dtype, np.number)]
        non_numeric_cols = [c for c in dfp.columns if not np.issubdtype(dfp[c].dtype, np.number)]
        if name_col is None and non_numeric_cols:
            name_col = non_numeric_cols[0]
        if price_col is None and numeric_cols:
            price_col = numeric_cols[-1]  # pick last numeric column as likely price

    # If either still None, give up
    if name_col is None or price_col is None:
        return {}

    # If somehow name_col or price_col are lists/dataframes (rare), normalize to single column name
    if isinstance(name_col, (list, tuple, pd.Index)):
        name_col = name_col[0]
    if isinstance(price_col, (list, tuple, pd.Index)):
        price_col = price_col[0]

    # If dfp[name_col] is DataFrame (multi-col selection), take first column
    name_series = dfp[name_col]
    if isinstance(name_series, pd.DataFrame):
        name_series = name_series.iloc[:, 0]

    price_series = dfp[price_col]
    if isinstance(price_series, pd.DataFrame):
        price_series = price_series.iloc[:, 0]

    # Build cleaned df
    tmp = pd.DataFrame({
        'name_raw': name_series.astype(str).fillna(""),
        'price_raw': pd.to_numeric(price_series, errors='coerce')
    })

    # Drop rows without name or without numeric price
    tmp = tmp.dropna(subset=['name_raw'])
    tmp = tmp.dropna(subset=['price_raw'])

    if tmp.empty:
        return {}

    tmp['Item'] = tmp['name_raw'].apply(normalize_name)
    tmp['UnitPrice'] = tmp['price_raw'].astype(float)

    # remove duplicates keep first (or you can take mean)
    tmp = tmp.drop_duplicates(subset=['Item'], keep='first')

    price_map = dict(zip(tmp['Item'], tmp['UnitPrice']))
    # debug print few examples
    sample = list(price_map.items())[:6]
    if sample:
        print("[read_price_map] contoh harga ditemukan:", sample)
    return price_map


def extract_items_and_numeric(path):
    """Return: (out_df (Item_norm, values_sum), original_df, item_col, numeric_cols, detection) 
       detection: 'qty' or 'revenue' or 'unknown'"""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    header_row = find_header_row(raw)
    df = pd.read_excel(path, header=header_row, engine="openpyxl")
    df = df.dropna(axis=1, how='all').dropna(axis=0, how='all')

    # find item col
    item_col = None
    for c in df.columns:
        try:
            if isinstance(c, str) and ("ITEM" in c.upper() or "DESCRIPTION" in c.upper()):
                item_col = c
                break
        except Exception:
            pass
    if item_col is None:
        item_col = df.columns[1] if len(df.columns)>1 else df.columns[0]

    # detect numeric columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if not numeric_cols:
        # fallback convertible columns
        cand = []
        for c in df.columns:
            if c == item_col: continue
            s = pd.to_numeric(df[c], errors='coerce')
            if s.count() > 0:
                cand.append(c)
        numeric_cols = cand

    # compute per-row sum of numeric cols
    if numeric_cols:
        df['SumNumeric'] = df[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
    else:
        df['SumNumeric'] = 0.0

    # quick detection: if mean of sums is large (>=1000) treat as revenue (Rp), else quantity
    mean_val = df['SumNumeric'].replace(0, np.nan).mean()
    if pd.isna(mean_val):
        detection = 'unknown'
    else:
        detection = 'revenue' if mean_val >= 1000 else 'qty'

    # build normalized items with SumNumeric
    res = pd.DataFrame({
        'Item_raw': df[item_col].astype(str).fillna(""),
        'SumNumeric': df['SumNumeric']
    })
    res['Item'] = res['Item_raw'].apply(normalize_name)

    # remove blockwords (cost rows)
    def is_cost(s):
        if not s: return False
        for w in BLOCKWORDS:
            if w in s:
                return True
        return False
    mask_cost = res['Item'].apply(is_cost)
    if mask_cost.any():
        print(f"[{os.path.basename(path)}] Baris yang dibuang karena label biaya (contoh max 8):")
        print(res.loc[mask_cost, ['Item_raw','Item','SumNumeric']].head(8).to_string(index=False))
    res = res[~mask_cost]
    res = res[res['Item'] != ""].copy()

    # aggregate duplicates (per file)
    out = res.groupby('Item', as_index=False)['SumNumeric'].sum().rename(columns={'SumNumeric':'ValueSum'})
    out = out.sort_values('ValueSum', ascending=False).reset_index(drop=True)
    return out, df, item_col, numeric_cols, detection

# ---------- main ----------
def main():
    # load price map (try from either file)
    price_map = read_price_map(AUG_PATH) or read_price_map(SEP_PATH)
    if price_map:
        print("[INFO] Harga per unit ditemukan dari price sheet (contoh 5):")
        sample_prices = list(price_map.items())[:5]
        for k,v in sample_prices:
            print(f"  {k} -> {v:,.0f}")
    else:
        print("[INFO] Tidak ditemukan price sheet otomatis. Jika datamu adalah revenue dan kamu butuh qty, sediakan sheet harga atau default price map.")

    # process both months
    aug_out, aug_df_raw, aug_item_col, aug_numeric_cols, aug_det = extract_items_and_numeric(AUG_PATH)
    sep_out, sep_df_raw, sep_item_col, sep_numeric_cols, sep_det = extract_items_and_numeric(SEP_PATH)

    print(f"\nDeteksi Agustus: {aug_det}. Deteksi September: {sep_det}.")

    # If detection == 'revenue' and price_map available -> compute qty = revenue / price
    def compute_quantity(out_df):
        df = out_df.copy()
        df['UnitPrice'] = df['Item'].map(price_map) if price_map else np.nan
        if price_map:
            # estimate qty (float), then round to nearest integer (or floor)
            df['EstimatedQty'] = df.apply(lambda r: (r['ValueSum'] / r['UnitPrice']) if pd.notna(r['UnitPrice']) and r['UnitPrice']>0 else np.nan, axis=1)
            df['EstimatedQtyInt'] = df['EstimatedQty'].apply(lambda x: int(np.floor(x)) if pd.notna(x) else np.nan)
            return df
        else:
            return df

    # Prepare final tables for plotting: either Qty or ValueSum depending on detection
    def prepare_for_plot(out_df, detection):
        if detection == 'qty':
            df = out_df.rename(columns={'ValueSum':'Qty'}).copy()
            df['PlotValue'] = df['Qty']
            df['Label'] = df['Item']
            return df[['Item','PlotValue','Label','Qty']]
        elif detection == 'revenue':
            if price_map:
                df = compute_quantity(out_df)
                # if EstimatedQtyInt available use that as PlotValue, else fallback to ValueSum
                if 'EstimatedQtyInt' in df.columns:
                    df['PlotValue'] = df['EstimatedQtyInt'].fillna(0)
                    df['Label'] = df['Item']
                    return df[['Item','PlotValue','Label','EstimatedQtyInt','ValueSum','UnitPrice']]
                else:
                    df = out_df.rename(columns={'ValueSum':'Revenue'})
                    df['PlotValue'] = df['Revenue']
                    df['Label'] = df['Item']
                    return df[['Item','PlotValue','Label','Revenue']]
            else:
                # no price map: we cannot get qty -> fallback to showing revenue
                df = out_df.rename(columns={'ValueSum':'Revenue'}).copy()
                df['PlotValue'] = df['Revenue']
                df['Label'] = df['Item']
                return df[['Item','PlotValue','Label','Revenue']]
        else:
            df = out_df.rename(columns={'ValueSum':'Value'}).copy()
            df['PlotValue'] = df['Value']
            df['Label'] = df['Item']
            return df[['Item','PlotValue','Label','Value']]

    aug_plot_df = prepare_for_plot(aug_out, aug_det)
    sep_plot_df = prepare_for_plot(sep_out, sep_det)

    # save CSVs with useful columns
    aug_plot_df.to_csv("totals_august_final.csv", index=False)
    sep_plot_df.to_csv("totals_september_final.csv", index=False)
    print("\n[SAVED] totals_august_final.csv, totals_september_final.csv")

    # Plot two vertical charts (top = August, bottom = September)
    # optionally show top N for readability
    TOP_N = None   # set e.g. 20 to show top 20 items
    def plot_vertical(df_top, df_bottom, title_top, title_bottom, out_png="charts_aug_sep_vertical.png", top_n=None):
        d1 = df_top.head(top_n) if top_n else df_top
        d2 = df_bottom.head(top_n) if top_n else df_bottom

        fig, axes = plt.subplots(nrows=2, ncols=1, figsize=(14,10), constrained_layout=True)
        axes[0].bar(range(len(d1)), d1['PlotValue'])
        axes[0].set_xticks(range(len(d1))); axes[0].set_xticklabels(d1['Label'], rotation=45, ha='right')
        axes[0].set_title(title_top)
        axes[0].set_ylabel('Estimated Qty' if 'Qty' in d1.columns or 'EstimatedQtyInt' in d1.columns else 'Value')

        axes[1].bar(range(len(d2)), d2['PlotValue'])
        axes[1].set_xticks(range(len(d2))); axes[1].set_xticklabels(d2['Label'], rotation=45, ha='right')
        axes[1].set_title(title_bottom)
        axes[1].set_ylabel('Estimated Qty' if 'Qty' in d2.columns or 'EstimatedQtyInt' in d2.columns else 'Value')

        plt.savefig(out_png, dpi=300)
        print(f"[SAVED] {out_png}")
        plt.show()

    plot_vertical(aug_plot_df, sep_plot_df,
                  "August — Qty (or Revenue if price missing)", "September — Qty (or Revenue if price missing)",
                  out_png="charts_aug_sep_vertical.png", top_n=TOP_N)

    # Print explanation examples for top 3 items of each month
    def explain_examples(original_df, item_col, numeric_cols, out_plot_df, detection, month_name):
        print(f"\n=== EXPLANATION SAMPLE — {month_name} (top 3) ===")
        sample_items = out_plot_df.head(3)['Item'].tolist()
        for it in sample_items:
            print(f"\nItem (normalized): {it}")
            # find rows in original corresponding to this item
            mask = original_df[item_col].astype(str).fillna("").apply(normalize_name) == it
            rows = original_df.loc[mask]
            if rows.empty:
                print("  (no direct matching row found in original file)")
                continue
            for idx, row in rows.iterrows():
                vals = []
                total = 0
                for c in numeric_cols:
                    v = pd.to_numeric(row.get(c, np.nan), errors='coerce')
                    if pd.isna(v): v = 0
                    vals.append((c, v))
                    total += v
                # print breakdown
                print(f"  Row index {idx}, Item raw: {row[item_col]}")
                for c,v in vals:
                    print(f"    {c} = {v:,.0f}")
                print(f"  Sum = {total:,.0f}")
                # if revenue detected and price_map available, show estimated qty
                if detection == 'revenue' and price_map:
                    unit = price_map.get(it, None)
                    if unit:
                        est_qty = total / unit
                        print(f"  -> Estimated qty = Sum / UnitPrice = {total:,.0f} / {unit:,.0f} = {est_qty:.2f} (round -> {int(np.floor(est_qty))})")
        print("=== END SAMPLE ===\n")

    explain_examples(aug_df_raw, aug_item_col, aug_numeric_cols, aug_plot_df, aug_det, "August")
    explain_examples(sep_df_raw, sep_item_col, sep_numeric_cols, sep_plot_df, sep_det, "September")

if __name__ == "__main__":
    main()
