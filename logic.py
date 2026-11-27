
import csv
import io
import re

import numpy as np
import pandas as pd

# Columns that identify one item in one store.
KEYS = ["Store_norm", "Article_norm", "Color_norm", "Size_norm"]

# --------------------------------------------------------------------------
# Mode settings
# --------------------------------------------------------------------------

MODES = {
    "Outward": {
        "date_label": "Outward On",
        "date_aliases": {"outwardon", "outwarddate", "outwardeddate"},
        "no_record_reason": "No Outward",
        "uses_safety": True,
    },
    "Inward": {
        "date_label": "inwardedon",
        "date_aliases": {"inwardedon", "inwardon", "inwarddate", "inwardeddate"},
        "no_record_reason": "No Inward",
        "uses_safety": False,
    },
}

# --------------------------------------------------------------------------
# Required columns. Values are the accepted header names after header_key()
# (lower case, letters and digits only), so "Outward On", "outward_on" and
# "OUTWARDON" are all the same header.
# --------------------------------------------------------------------------

ITEM_COLUMNS = {
    "Article": {"article"},
    "Color": {"color", "colour"},
    "Size": {"size"},
}
STORE_COLUMN = {"Store": {"pos", "storename", "store"}}
STOCK_COLUMNS = {**ITEM_COLUMNS, "Qty": {"qty", "quantity"}}

COLUMN_LABELS = {"Store": "Pos / store name"}

TEMPLATES = {
    "Outward": (
        ["Pos", "Article", "Color", "Size", "Outward On"],
        ["MH PUNE", "19529", "DEEP TEAL", "M", "16-03-2026"],
    ),
    "Inward": (
        ["store name", "article", "color", "size", "inwardedon"],
        ["MH PUNE", "20024", "BLUE SAPPHIRE", "M", "15-04-2026"],
    ),
    "Current_Stock": (
        ["Article", "Color", "Size", "Qty", "Pos"],
        ["13111", "AFRICAN VIOLET", "M", "1", "MH PUNE"],
    ),
    "Sales": (
        ["Store", "Article", "Color", "S", "M", "L", "XL"],
        ["MH PUNE", "19529", "DEEP TEAL", "1", "0", "2", "0"],
    ),
}

# Whole sheets that are optional: every header gets the yellow fill.
OPTIONAL_TEMPLATE_SHEETS = {
    "Sales": "Optional, RTV report only. One row per Store, Article, Color and one column per size. Other columns are ignored.",
}

# Template columns that are optional; they get a yellow header.
OPTIONAL_TEMPLATE_COLUMNS = {
    "Current_Stock": {"Pos": "Optional. Add only when all stores are in one file."},
}


# Age slabs for the Stock Age report: 0-30, 31-60, 61-90, 91-180, 180+
SLAB_LIMITS = [30, 60, 90, 180]
NOT_IN_RECORD = "Not in record"


def slab_labels():
    labels, start = [], 0
    for limit in SLAB_LIMITS:
        labels.append(f"{start}-{limit}")
        start = limit + 1
    labels.append(f"{SLAB_LIMITS[-1]}+")
    return labels


# --------------------------------------------------------------------------
# File reading and headers
# --------------------------------------------------------------------------

def read_any_file(file_like, name):
    """Read an .xlsx or .csv file as text, so nothing is changed on load."""
    ext = str(name).lower().rsplit(".", 1)[-1]
    if ext == "csv":
        try:
            df = pd.read_csv(file_like, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file_like.seek(0)
            df = pd.read_csv(file_like, dtype=str, encoding="latin-1")
    elif ext in ("xlsx", "xlsm"):
        df = pd.read_excel(file_like, dtype=str)
    else:
        raise ValueError("only .xlsx and .csv files are supported")
    return df


def header_key(name):
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def standardize_columns(df, wanted):
    """
    Find the wanted columns in df whatever their case or spacing.
    Returns (frame with the canonical column names, list of missing names).
    """
    positions = {}
    for pos, col in enumerate(df.columns):
        key = header_key(col)
        for canon, aliases in wanted.items():
            if canon not in positions and key in aliases:
                positions[canon] = pos
                break

    missing = [c for c in wanted if c not in positions]
    if missing:
        return None, missing

    out = df.iloc[:, [positions[c] for c in wanted]].copy()
    out.columns = list(wanted)
    return out, []


# --------------------------------------------------------------------------
# Cleaning helpers
# --------------------------------------------------------------------------

def clean_text(series):
    s = series.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    return s.mask(s.str.lower().isin(["nan", "none", "null"]), "")


def clean_article(series):
    """Keep the article as text; only drop a trailing '.0' added by Excel."""
    return clean_text(series).str.replace(r"^(\d+)\.0+$", r"\1", regex=True)


_DATE_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%y",
    "%d/%m/%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%y",
    "%Y/%m/%d",
    "%d.%m.%Y",
    "%d.%m.%y",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%d-%m-%Y %H:%M",
    "%d/%m/%Y %H:%M",
    "%Y-%m-%d %H:%M:%S.%f",
    "%d-%b-%Y",
    "%d-%b-%y",
    "%d %b %Y",
    "%d-%m-%Y %I:%M:%S %p",
    "%d/%m/%Y %I:%M:%S %p",
]


def _naive(parsed):
    if getattr(parsed.dt, "tz", None) is not None:
        return parsed.dt.tz_localize(None)
    return parsed


def parse_dates(series):
    """
    Convert text dates to datetimes without ever raising an error.
    Day comes first in ambiguous dates (15-04-26 is 15 April 2026).
    Excel serial numbers are accepted. Anything unreadable becomes NaT.
    """
    text = series.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    pending = text != ""
    pieces = []

    for fmt in _DATE_FORMATS:
        if not pending.any():
            break
        parsed = pd.to_datetime(text[pending], format=fmt, errors="coerce")
        ok = parsed.notna()
        if ok.any():
            pieces.append(_naive(parsed[ok]))
            pending.loc[ok[ok].index] = False

    if pending.any():
        numbers = pd.to_numeric(text[pending], errors="coerce")
        is_serial = numbers.between(20000, 80000)
        if is_serial.any():
            pieces.append(pd.to_datetime(numbers[is_serial], unit="D", origin="1899-12-30"))
            pending.loc[is_serial[is_serial].index] = False

    if pending.any():
        rest = text[pending]
        try:
            parsed = pd.to_datetime(rest, format="mixed", dayfirst=True, errors="coerce")
        except (TypeError, ValueError):
            parsed = pd.to_datetime(rest, dayfirst=True, errors="coerce")
        ok = parsed.notna()
        if ok.any():
            pieces.append(_naive(parsed[ok]))

    if not pieces:
        return pd.Series(pd.NaT, index=text.index, dtype="datetime64[ns]")
    return pd.to_datetime(pd.concat(pieces).reindex(text.index))


def norm_key(text):
    return re.sub(r"\s+", " ", str(text).strip()).upper()


# --------------------------------------------------------------------------
# Movement file (Outward or Inward)
# --------------------------------------------------------------------------

def process_movement(raw, mode):
    """
    Returns {"ok": True, "stores", "history_days", "min_date", "max_date",
    "summary"} or {"ok": False, "error"}.
    summary has one row per Store + Article + Color + Size with the first
    date, last date and number of records.
    """
    cfg = MODES[mode]
    wanted = {**STORE_COLUMN, **ITEM_COLUMNS, "Date": cfg["date_aliases"]}
    std, missing = standardize_columns(raw, wanted)
    if missing:
        names = [cfg["date_label"] if m == "Date" else COLUMN_LABELS.get(m, m) for m in missing]
        return {"ok": False, "error": f"Missing column(s) in the {mode} file: {', '.join(names)}"}

    store = clean_text(std["Store"])
    frame = pd.DataFrame(
        {
            "Store_display": store,
            "Store_norm": store.str.upper(),
            "Article_norm": clean_article(std["Article"]).str.upper(),
            "Color_norm": clean_text(std["Color"]).str.upper(),
            "Size_norm": clean_text(std["Size"]).str.upper(),
            "Date": parse_dates(std["Date"]).dt.normalize(),
        }
    )
    frame = frame[(frame["Store_norm"] != "") & (frame["Article_norm"] != "")]

    stores = (
        frame.drop_duplicates("Store_norm")[["Store_norm", "Store_display"]]
        .sort_values("Store_display")["Store_display"]
        .tolist()
    )
    if not stores:
        return {"ok": False, "error": f"No stores found in the {mode} file."}

    dated = frame.dropna(subset=["Date"])
    if dated.empty:
        return {"ok": False, "error": f"No readable dates found in '{cfg['date_label']}'."}

    min_date = dated["Date"].min()
    max_date = dated["Date"].max()

    summary = (
        dated.groupby(KEYS, sort=False)
        .agg(First=("Date", "min"), Last=("Date", "max"), Count=("Date", "size"))
        .reset_index()
    )
    return {
        "ok": True,
        "stores": stores,
        "history_days": int((max_date - min_date).days),
        "min_date": min_date,
        "max_date": max_date,
        "summary": summary,
    }


# --------------------------------------------------------------------------
# Current Stock files
# --------------------------------------------------------------------------

def _store_column_position(raw):
    for pos, col in enumerate(raw.columns):
        if header_key(col) in STORE_COLUMN["Store"]:
            return pos
    return None


def process_stock(raw):
    """
    Clean one Current Stock file and total the quantity per item.
    If the file has a Pos / store name column with names in it, it is a
    merged file: the result keeps the store name of every row.
    """
    std, missing = standardize_columns(raw, STOCK_COLUMNS)
    if missing:
        return {"ok": False, "error": f"missing column(s): {', '.join(missing)}"}

    qty = pd.to_numeric(std["Qty"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    frame = pd.DataFrame(
        {
            "Article": clean_article(std["Article"]),
            "Color": clean_text(std["Color"]),
            "Size": clean_text(std["Size"]),
            "Qty": qty.fillna(0),
        }
    )

    store = None
    store_pos = _store_column_position(raw)
    if store_pos is not None:
        names = clean_text(raw.iloc[:, store_pos])
        if (names != "").any():
            store = names
            frame["Source_name"] = names

    frame = frame[frame["Article"] != ""].copy()
    blank_rows = 0
    if store is not None:
        blank_rows = int((frame["Source_name"] == "").sum())
        frame = frame[frame["Source_name"] != ""].copy()
        frame["Source_key"] = frame["Source_name"].str.replace(r"[\s_\-]+", "", regex=True).str.upper()

    frame["Article_norm"] = frame["Article"].str.upper()
    frame["Color_norm"] = frame["Color"].str.upper()
    frame["Size_norm"] = frame["Size"].str.upper()

    group_cols = ["Article_norm", "Color_norm", "Size_norm"]
    agg = dict(Article=("Article", "first"), Color=("Color", "first"), Size=("Size", "first"), Qty=("Qty", "sum"))
    if store is None:
        return {"ok": True, "merged": False, "data": frame.groupby(group_cols, as_index=False).agg(**agg)}

    agg["Source_name"] = ("Source_name", "first")
    data = frame.groupby(["Source_key"] + group_cols, as_index=False).agg(**agg)
    return {"ok": True, "merged": True, "data": data, "blank_rows": blank_rows}


def source_stores(data):
    """(key, name) of every store name found in a merged file, sorted by name."""
    pairs = data.drop_duplicates("Source_key")[["Source_key", "Source_name"]]
    return list(pairs.sort_values("Source_name").itertuples(index=False, name=None))


def attach_store(data, store_display):
    """Tag a per-store file's data with its store."""
    out = data.copy()
    out["Store_display"] = store_display
    out["Store_norm"] = norm_key(store_display)
    return out


def attach_stores(data, mapping):
    """Tag a merged file's data. mapping: Source_key -> movement store name."""
    out = data[data["Source_key"].isin(mapping)].copy()
    out["Store_display"] = out["Source_key"].map(mapping)
    out["Store_norm"] = out["Store_display"].map(norm_key)
    return out.drop(columns=["Source_key", "Source_name"])


def _name_key(text):
    return re.sub(r"[\s_\-]+", "", str(text)).upper()


def match_store(name, stores):
    """Exact match on a store name; case, spaces, _ and - are ignored."""
    key = _name_key(name)
    if not key:
        return None
    for store in stores:
        if _name_key(store) == key:
            return store
    return None


def suggest_store(filename, stores):
    """Match a file name (without its extension) to a store name."""
    return match_store(re.sub(r"\.[A-Za-z0-9]+$", "", filename), stores)


# --------------------------------------------------------------------------
# Sale file (optional, RTV report only)
# One row per Store + Article + Color, one column per size. There is no date:
# the file covers whatever period it was exported for. Any size with a sale
# quantity above 0 keeps that item out of the RTV list.
# --------------------------------------------------------------------------

_SALE_ID_KEYS = STORE_COLUMN["Store"] | ITEM_COLUMNS["Article"] | ITEM_COLUMNS["Color"]

# Columns that are not sizes. Every other column is read as a size.
_SALE_META_KEYS = _SALE_ID_KEYS | {
    "", "nan", "section", "subsection", "subsubsection", "brand", "total", "soh", "mrp", "cp",
    "clearancediscount", "otherofferdiscount", "giftvoucher", "spsdiscount",
    "totalsalesalecreditnotewallet", "exchangeqty", "exchangeamount",
}


def _sale_header_row(raw, max_rows=10):
    """The report may have a title row above the headers; find the real header row."""
    for i in range(min(max_rows, len(raw))):
        keys = {header_key(v) for v in raw.iloc[i].tolist()}
        if keys & STORE_COLUMN["Store"] and keys & ITEM_COLUMNS["Article"] and keys & ITEM_COLUMNS["Color"]:
            return i
    return 0


def read_sale_file(file_like, name):
    ext = str(name).lower().rsplit(".", 1)[-1]
    if ext == "csv":
        content = file_like.read()
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("latin-1")
        rows = list(csv.reader(io.StringIO(text)))
        width = max((len(r) for r in rows), default=0)
        raw = pd.DataFrame([r + [""] * (width - len(r)) for r in rows], dtype=str)
    elif ext in ("xlsx", "xlsm"):
        raw = pd.read_excel(file_like, dtype=str, header=None)
    else:
        raise ValueError("only .xlsx and .csv files are supported")

    header = _sale_header_row(raw)
    df = raw.iloc[header + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in raw.iloc[header]]
    return df


def process_sale(df):
    """Returns {"ok": True, "data"} with one row per Store, Article, Color, Size that sold."""
    pos = {}
    for i, col in enumerate(df.columns):
        key = header_key(col)
        for name, aliases in (("Store", STORE_COLUMN["Store"]), ("Article", ITEM_COLUMNS["Article"]),
                              ("Color", ITEM_COLUMNS["Color"])):
            if name not in pos and key in aliases:
                pos[name] = i
    missing = [n for n in ("Store", "Article", "Color") if n not in pos]
    if missing:
        return {"ok": False, "error": f"Sale file is missing: {', '.join(missing)}"}

    size_pos = [
        i for i, col in enumerate(df.columns)
        if i not in pos.values()
        and header_key(col) not in _SALE_META_KEYS
        and not str(col).lower().startswith("unnamed")
    ]
    if not size_pos:
        return {"ok": False, "error": "Sale file has no size columns."}

    base = pd.DataFrame(
        {
            "Store_name": clean_text(df.iloc[:, pos["Store"]]),
            "Article_norm": clean_article(df.iloc[:, pos["Article"]]).str.upper(),
            "Color_norm": clean_text(df.iloc[:, pos["Color"]]).str.upper(),
        }
    )
    sizes = df.iloc[:, size_pos].apply(
        lambda c: pd.to_numeric(c.astype(str).str.replace(",", "", regex=False), errors="coerce")
    )
    size_names = {str(n): clean_text(pd.Series([df.columns[i]])).iloc[0].upper() for n, i in enumerate(size_pos)}
    sizes.columns = [str(n) for n in range(len(size_pos))]

    long = pd.concat([base, sizes], axis=1).melt(
        id_vars=list(base.columns), var_name="Size_norm", value_name="Sale_Qty"
    )
    long = long[(long["Sale_Qty"] > 0) & (long["Store_name"] != "") & (long["Article_norm"] != "")]
    long = long[~long["Store_name"].str.replace(r"\s+", "", regex=True).str.lower().isin(["total", "grandtotal"])]
    long["Size_norm"] = long["Size_norm"].map(size_names)

    data = long.groupby(["Store_name", "Article_norm", "Color_norm", "Size_norm"], as_index=False)["Sale_Qty"].sum()
    return {"ok": True, "data": data}


def match_sales(sale, stores):
    """
    Match the sale file's store names to the movement file's stores.
    Returns (sold keys, unmatched store names, article count, item count).
    """
    lookup = {name: match_store(name, stores) for name in sale["Store_name"].drop_duplicates()}
    matched = sale["Store_name"].map(lookup)
    unmatched = sorted(name for name, store in lookup.items() if store is None)
    kept = sale[matched.notna()].copy()
    kept["Store_norm"] = matched[matched.notna()].map(norm_key)
    sold = kept[KEYS].drop_duplicates().reset_index(drop=True)
    return sold, unmatched, int(kept["Article_norm"].nunique()), len(sold)


# --------------------------------------------------------------------------
# RTV report
# --------------------------------------------------------------------------

def cutoff_days(rtv_age, safety_days):
    return int(rtv_age) + int(safety_days)


def report_columns(mode):
    return [
        "Store Name", "Article", "Color", "Size", "Current Stock",
        "RTV Days Range", "RTV Qty", "Reason",
        f"First {mode}", f"Last {mode}", f"No. of {mode}",
    ]


def build_report(stock, summary, report_date, cutoff, mode, sold=None):
    """
    An item is returned when its last movement is at least `cutoff` days old,
    or when the movement file has no record of it at all.
    Items in `sold` (from the Sale file) are kept out even if they qualify.
    Returns (report, number of items kept out because they are selling).
    """
    cfg = MODES[mode]
    columns = report_columns(mode)

    stock = stock[stock["Qty"] > 0]
    merged = stock.merge(summary, on=KEYS, how="left")

    no_record = merged["Last"].isna()
    age = (pd.Timestamp(report_date).normalize() - merged["Last"]).dt.days
    merged = merged[no_record | (age >= cutoff)].reset_index(drop=True)

    excluded = 0
    if sold is not None and not sold.empty and not merged.empty:
        flag = sold[KEYS].drop_duplicates().assign(_sold=True)
        selling = merged.merge(flag, on=KEYS, how="left")["_sold"].eq(True).to_numpy()
        excluded = int(selling.sum())
        merged = merged[~selling].reset_index(drop=True)
    if merged.empty:
        return pd.DataFrame(columns=columns), excluded

    no_record = merged["Last"].isna()
    qty = merged["Qty"]
    if (qty % 1 == 0).all():
        qty = qty.astype("int64")

    report = pd.DataFrame(
        {
            "Store Name": merged["Store_display"],
            "Article": merged["Article"],
            "Color": merged["Color"],
            "Size": merged["Size"],
            "Current Stock": qty,
            "RTV Days Range": f"{cutoff}+ Days",
            "RTV Qty": qty,
            "Reason": np.where(no_record, cfg["no_record_reason"], "Old Stock"),
            f"First {mode}": merged["First"],
            f"Last {mode}": merged["Last"],
            f"No. of {mode}": merged["Count"].fillna(0).astype("int64"),
        }
    )
    return report[columns], excluded


# --------------------------------------------------------------------------
# Stock Age report
# --------------------------------------------------------------------------

def age_columns(mode):
    return [
        "Store Name", "Article", "Color", "Size", "Current Stock",
        "Age (Days)", "Age Slab", f"Last {mode}",
    ]


def build_age_report(stock, summary, report_date, history_days, mode):
    """
    Age of all stock in hand: days from the last movement to the report date.
    Items with no record get the days covered by the movement file with a
    plus sign (for example 25+) and the slab "Not in record".
    """
    columns = age_columns(mode)
    stock = stock[stock["Qty"] > 0]
    merged = stock.merge(summary[KEYS + ["Last"]], on=KEYS, how="left")
    if merged.empty:
        return pd.DataFrame(columns=columns)

    no_record = merged["Last"].isna()
    age = (pd.Timestamp(report_date).normalize() - merged["Last"]).dt.days

    merged["_no_record"] = no_record
    merged["_age"] = age
    merged = merged.sort_values(
        ["Store_display", "_no_record", "_age", "Article_norm", "Color_norm", "Size_norm"],
        ascending=[True, False, False, True, True, True],
        kind="stable",
    ).reset_index(drop=True)
    no_record = merged["_no_record"]
    age = merged["_age"]

    bins = [-np.inf] + SLAB_LIMITS + [np.inf]
    slab = pd.cut(age.fillna(0), bins=bins, labels=slab_labels()).astype(object)
    slab[no_record] = NOT_IN_RECORD

    age_shown = age.fillna(0).astype("int64").astype(object)
    age_shown[no_record] = f"{history_days}+"

    qty = merged["Qty"]
    if (qty % 1 == 0).all():
        qty = qty.astype("int64")

    report = pd.DataFrame(
        {
            "Store Name": merged["Store_display"],
            "Article": merged["Article"],
            "Color": merged["Color"],
            "Size": merged["Size"],
            "Current Stock": qty,
            "Age (Days)": age_shown,
            "Age Slab": slab,
            f"Last {mode}": merged["Last"],
        }
    )
    return report[columns]


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

def _style_header(ws, widths):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def build_excel_bytes(report, sheet_name="RTV_Report"):
    """An Excel workbook with one sheet holding the report. Dates are DD-MM-YYYY only."""
    out = report.copy()
    date_cols = [c for c in out.columns if pd.api.types.is_datetime64_any_dtype(out[c])]
    for col in date_cols:
        out[col] = out[col].dt.date  # drop the time part

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        sample = out.head(1000)
        widths = []
        for col in out.columns:
            longest = max([len(str(col))] + [len(str(v)) for v in sample[col]])
            widths.append(min(max(longest + 3, 12), 40))
        _style_header(ws, widths)

        for col in date_cols:
            position = list(out.columns).index(col) + 1
            for row in ws.iter_rows(min_row=2, min_col=position, max_col=position):
                row[0].number_format = "DD-MM-YYYY"

        ws.auto_filter.ref = ws.dimensions
    return buffer.getvalue()


def build_template_bytes():
    """A workbook with one sheet per upload file: required headers and an example row."""
    from openpyxl.comments import Comment
    from openpyxl.styles import Font, PatternFill

    yellow = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet, (headers, example) in TEMPLATES.items():
            pd.DataFrame([example], columns=headers).to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            _style_header(ws, [max(len(h) + 6, 16) for h in headers])
            sheet_note = OPTIONAL_TEMPLATE_SHEETS.get(sheet)
            for pos, header in enumerate(headers, start=1):
                note = sheet_note if (sheet_note and pos == 1) else OPTIONAL_TEMPLATE_COLUMNS.get(sheet, {}).get(header)
                if sheet_note or note:
                    cell = ws.cell(row=1, column=pos)
                    cell.fill = yellow
                    cell.font = Font(bold=True, color="000000")
                    if note:
                        cell.comment = Comment(note, "RTV Report")
    return buffer.getvalue()
