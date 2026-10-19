#!/usr/bin/env python3
"""Generic data cleaner: merge, trim, fix types and dates, dedupe, report missing values.

Usage:
    python data_clean.py INPUT [INPUT ...] [-o cleaned.xlsx] [--key COL [COL ...]]
                         [--dayfirst] [--no-merge] [--fill-numeric {none,zero,median}]

INPUT can be files (.csv .tsv .txt .xlsx .xls) or folders. Every sheet of every
workbook is read. Steps, in order:
  1. Read all inputs (auto-detects CSV delimiter and encoding: utf-8 / utf-8-sig / cp1252 / latin-1)
  2. Normalise headers (trim, collapse spaces, match case-insensitively across files)
  3. Trim whitespace in text, collapse double spaces, turn blank / "N/A" / "null" / "-" into missing
  4. Drop rows and columns that are completely empty
  5. Detect numeric columns ($, commas, %, (negatives), stray spaces) and convert them
  6. Detect date columns and convert them to real dates (mixed formats are OK)
  7. Remove exact duplicate rows, or duplicates on --key columns (keeps the first)
  8. Missing-values report per column
Outputs: cleaned .xlsx (Cleaned, Removed Duplicates, Missing Values, Column Types,
Change Log sheets) plus a <output>_summary.md in plain English.
"""
import argparse
import csv
import io
import re
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

NULLS = {"", "na", "n/a", "nan", "null", "none", "-", "--", "?", "#n/a", "nil", "missing", "undefined"}
READABLE = {".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".xls"}
DATE_HINT = re.compile(r"date|time|dob|day|created|updated|joined|period|month", re.I)
NUM_CLEAN = re.compile(r"[\s,$€£₹¥]|(?<=\d)%$")
ID_SUFFIX = re.compile(r"(Id|ID|Code|No|Num|Number|Ref)$|invoice|order[\s_]?(no|id|num)")  # camelCase IDs, e.g. InvoiceNo, StockCode, CustomerID
ID_HINT = re.compile(r"(^|[\s_])(id|code|zip|postal|pin|phone|mobile|tel|sku|account|acct)([\s_]|$)", re.I)


# ---------------- reading ----------------
def read_text_file(path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        sep = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|").delimiter
    except csv.Error:
        sep = "\t" if path.suffix.lower() == ".tsv" else ","
    df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False, skip_blank_lines=True)
    return [(path.name, df, enc, sep)]


def read_any(path):
    ext = path.suffix.lower()
    if ext in (".csv", ".tsv", ".txt"):
        return read_text_file(path)
    sheets = pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False)
    return [(f"{path.name} [{name}]", df, "", "") for name, df in sheets.items() if not df.empty]


def collect(inputs):
    files = []
    for i in inputs:
        p = Path(i)
        if p.is_dir():
            files += sorted(f for f in p.rglob("*") if f.suffix.lower() in READABLE and not f.name.startswith("~$"))
        elif p.exists():
            files.append(p)
        else:
            sys.exit(f"Not found: {i}")
    return files


# ---------------- cleaning helpers ----------------
def norm_header(h):
    h = re.sub(r"\s+", " ", str(h)).strip()
    return h if h and not h.lower().startswith("unnamed:") else ""


def to_number(s):
    """Convert a text series to numbers; returns (numbers, share_parsed_of_non_null)."""
    t = s.dropna().astype(str).str.strip()
    if t.empty:
        return None, 0.0
    neg = t.str.match(r"^\(.*\)$")
    pct = t.str.endswith("%")
    c = t.str.replace(r"^\((.*)\)$", r"\1", regex=True).map(lambda x: NUM_CLEAN.sub("", x))
    n = pd.to_numeric(c, errors="coerce").astype(float)
    n[neg] = -n[neg]
    n[pct] = n[pct] / 100
    ok = n.notna().mean()
    out = pd.Series(pd.NA, index=s.index, dtype="Float64")
    out.loc[n.index] = n
    return out, ok


def to_date(s, dayfirst):
    try:
        return _to_date(s, dayfirst)
    except (pd.errors.OutOfBoundsDatetime, OverflowError, ValueError):
        # text that only looks date-like (e.g. product names with numbers) - treat as not a date column
        return None, 0.0


def _to_date(s, dayfirst):
    t = s.dropna().astype(str).str.strip()
    if t.empty:
        return None, 0.0
    # Excel serial numbers (e.g. 45231) in a date-named column
    if t.str.fullmatch(r"\d{5}(\.\d+)?").mean() > 0.9:
        d = pd.to_datetime(pd.to_numeric(t), unit="D", origin="1899-12-30", errors="coerce")
    else:
        # year-first values (2026-01-05, 2026/01/12) are never day-first, so parse them separately
        iso = t.str.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}")
        d = pd.Series(pd.NaT, index=t.index, dtype="datetime64[ns]")
        if iso.any():
            d[iso] = pd.to_datetime(t[iso], errors="coerce", format="mixed")
        if (~iso).any():
            d[~iso] = pd.to_datetime(t[~iso], errors="coerce", dayfirst=dayfirst, format="mixed")
    ok = d.notna().mean()
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    out.loc[d.index] = d
    return out, ok


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-o", "--output", default="cleaned.xlsx")
    ap.add_argument("--key", nargs="+", help="columns that identify a unique record (dedupe on these)")
    ap.add_argument("--dayfirst", action="store_true", help="read 03/04/2026 as 3 April (default: March 4)")
    ap.add_argument("--no-merge", action="store_true", help="clean each file/sheet separately")
    ap.add_argument("--fill-numeric", choices=["none", "zero", "median"], default="none",
                    help="fill missing numbers (default: leave blank, just report)")
    ap.add_argument("--num-threshold", type=float, default=0.9, help="share of values that must parse as numbers")
    ap.add_argument("--date-threshold", type=float, default=0.8, help="share of values that must parse as dates")
    a = ap.parse_args()

    log = []
    frames = []
    for f in collect(a.inputs):
        for label, df, enc, sep in read_any(f):
            df.columns = [norm_header(c) for c in df.columns]
            blank = [i for i, c in enumerate(df.columns) if c == ""]
            if blank:
                df = df.loc[:, [c != "" for c in df.columns]] if df.iloc[:, blank].replace("", pd.NA).isna().all().all() else df
                df.columns = [c or f"Column {i+1}" for i, c in enumerate(df.columns)]
            frames.append((label, df))
            log.append(f"Read {label}: {len(df)} rows x {df.shape[1]} cols"
                       + (f" (encoding {enc}, delimiter {sep!r})" if enc else ""))
    if not frames:
        sys.exit("No readable files.")

    groups = [(lab, [(lab, d)]) for lab, d in frames] if a.no_merge else [("Cleaned", frames)]
    out = Path(a.output)
    writer = pd.ExcelWriter(out, engine="openpyxl")
    summary = [f"# Data cleaning summary\n", f"Files/sheets read: {len(frames)}\n"]

    for gname, parts in groups:
        # ---- merge with case-insensitive header matching ----
        # canonical header = the first spelling that isn't all-lowercase (e.g. "Order Date" over "order date")
        canon = {}
        for _, d in parts:
            for c in d.columns:
                k = c.lower()
                if k not in canon or (canon[k] == canon[k].lower() and c != c.lower()):
                    canon[k] = c
        merged = []
        for lab, d in parts:
            d = d.copy()
            d.columns = [canon[c.lower()] for c in d.columns]
            d = d.loc[:, ~pd.Index(d.columns).duplicated()]
            d.insert(0, "Source File", lab)
            merged.append(d)
        df = pd.concat(merged, ignore_index=True, sort=False)
        n0 = len(df)
        stats = {"rows_in": n0}
        if len(parts) > 1:
            sets = [set(map(str.lower, d.columns)) - {"source file"} for _, d in parts]
            only_some = sorted(set.union(*sets) - set.intersection(*sets))
            log.append(f"[{gname}] merged {len(parts)} inputs -> {n0} rows, {df.shape[1]-1} columns")
            if only_some:
                log.append(f"[{gname}] columns missing from some files (left blank there): {', '.join(only_some)}")

        # ---- trim / nulls ----
        cols = [c for c in df.columns if c != "Source File"]
        trimmed = 0
        for c in cols:
            s = df[c].astype("string")
            new = s.str.replace("\u00a0", " ", regex=False).str.replace(r"\s+", " ", regex=True).str.strip()
            trimmed += int(((s != new) & s.notna()).sum())
            new = new.mask(new.str.lower().isin(NULLS))
            df[c] = new
        log.append(f"[{gname}] trimmed / collapsed whitespace in {trimmed} cells; standardised blank-like values to empty")
        stats["trimmed"] = trimmed

        # ---- empty rows / cols ----
        empty_rows = df[cols].isna().all(axis=1)
        df = df[~empty_rows]
        empty_cols = [c for c in cols if df[c].isna().all()]
        df = df.drop(columns=empty_cols)
        cols = [c for c in cols if c not in empty_cols]
        stats["empty_rows"] = int(empty_rows.sum())
        log.append(f"[{gname}] dropped {int(empty_rows.sum())} empty rows"
                   + (f" and empty columns: {', '.join(empty_cols)}" if empty_cols else ""))

        # ---- types ----
        types = []
        for c in cols:
            s = df[c]
            nonnull = s.notna().sum()
            kind, detail = "text", ""
            looks_id = bool(ID_HINT.search(c)) or bool(ID_SUFFIX.search(c)) or s.dropna().astype(str).str.match(r"^0\d+$").mean() > 0.3
            if nonnull and not looks_id:
                if DATE_HINT.search(c) or s.dropna().astype(str).str.contains(r"[/\-.]\d{2,4}|\d{4}-\d{2}|[A-Za-z]{3}").mean() > 0.6:
                    dser, ok = to_date(s, a.dayfirst)
                    numeric_only = s.dropna().astype(str).str.fullmatch(r"-?[\d,.]+").mean() > 0.9 and not DATE_HINT.search(c)
                    if dser is not None and ok >= a.date_threshold and not numeric_only:
                        bad = int(s.notna().sum() - dser.notna().sum())
                        df[c] = dser
                        kind, detail = "date", f"{bad} unparseable value(s) set to blank" if bad else ""
                if kind == "text":
                    nser, ok = to_number(s)
                    if nser is not None and ok >= a.num_threshold:
                        bad = int(s.notna().sum() - nser.notna().sum())
                        if (nser.dropna() % 1 == 0).all():
                            nser = nser.astype("Int64")
                        df[c] = nser
                        kind, detail = "number", f"{bad} non-numeric value(s) set to blank" if bad else ""
            elif looks_id:
                detail = "kept as text (ID/code/phone; leading zeros preserved)"
            types.append({"Column": c, "Detected Type": kind, "Notes": detail})
            if detail and kind != "text":
                log.append(f"[{gname}] {c}: {detail}")
        stats["types"] = types

        # ---- duplicates ----
        key = None
        if a.key:
            lower = {c.lower(): c for c in cols}
            key = [lower.get(k.lower()) for k in a.key]
            if None in key:
                sys.exit(f"--key column not found. Available: {', '.join(cols)}")
        cmp = df[key or cols].copy()
        for c in cmp.columns:
            if cmp[c].dtype == "string":
                cmp[c] = cmp[c].str.lower()
        dup_mask = cmp.duplicated(keep="first")
        dups = df[dup_mask]
        df = df[~dup_mask].reset_index(drop=True)
        stats["dups"] = len(dups)
        log.append(f"[{gname}] removed {len(dups)} duplicate rows "
                   + (f"(same {', '.join(key)})" if key else "(identical across all columns, case-insensitive)"))

        # ---- fill numeric (optional) ----
        if a.fill_numeric != "none":
            for t in types:
                if t["Detected Type"] == "number":
                    c = t["Column"]; nmiss = int(df[c].isna().sum())
                    if nmiss:
                        val = 0 if a.fill_numeric == "zero" else df[c].median()
                        df[c] = df[c].fillna(val)
                        log.append(f"[{gname}] {c}: filled {nmiss} missing with {a.fill_numeric} ({val})")

        # ---- missing report ----
        miss = pd.DataFrame({"Column": cols,
                             "Missing": [int(df[c].isna().sum()) for c in cols],
                             "Filled": [int(df[c].notna().sum()) for c in cols]})
        miss["% Missing"] = (miss["Missing"] / max(len(df), 1) * 100).round(1)
        miss = miss.sort_values("% Missing", ascending=False)
        stats["rows_out"] = len(df)

        # ---- write ----
        if a.no_merge:
            base = re.sub(r"[\[\]:*?/\\]", "", gname).strip()[:21]
            names = (base, f"{base} dups", f"{base} missing", f"{base} types")
        else:
            names = ("Cleaned", "Removed Duplicates", "Missing Values", "Column Types")
        df.to_excel(writer, sheet_name=names[0], index=False)
        dups.to_excel(writer, sheet_name=names[1], index=False)
        miss.to_excel(writer, sheet_name=names[2], index=False)
        pd.DataFrame(types).to_excel(writer, sheet_name=names[3], index=False)

        summary += [f"\n## {gname}\n",
                    f"- Rows in: {n0} -> rows out: {stats['rows_out']}",
                    f"- Empty rows dropped: {stats['empty_rows']}",
                    f"- Duplicate rows removed: {stats['dups']}" + (f" (key: {', '.join(key)})" if key else ""),
                    f"- Cells with extra spaces fixed: {stats['trimmed']}",
                    "- Column types: " + "; ".join(f"{t['Column']} = {t['Detected Type']}" for t in types),
                    "- Columns with missing values: " + (", ".join(f"{r.Column} ({r.Missing}, {r._4}%)"
                        for r in miss.itertuples() if r.Missing) or "none")]

    pd.DataFrame({"Step": log}).to_excel(writer, sheet_name="Change Log", index=False)

    # formatting: widths, date format, header style
    from openpyxl.styles import Font, PatternFill
    for ws in writer.book.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="305496")
        for col in ws.columns:
            width = max((len(str(c.value)) for c in list(col)[:200] if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)
            for c in col[1:]:
                if hasattr(c.value, "year"):
                    c.number_format = "yyyy-mm-dd"
        ws.freeze_panes = "A2"
    writer.close()

    summary.append("\n## Change log\n")
    summary += [f"- {l}" for l in log]
    sm = out.with_name(out.stem + "_summary.md")
    sm.write_text("\n".join(summary) + "\n", encoding="utf-8")
    print("\n".join(summary))
    print(f"\nSaved {out} and {sm}")


if __name__ == "__main__":
    main()
