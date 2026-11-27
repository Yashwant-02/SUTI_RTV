"""
app.py
------
Streamlit interface for the RTV report tool.
Business rules live in logic.py; this file only handles the screen.
"""

import io
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import logic

BLANK = "- Select store -"
SKIP = "Skip (leave out)"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PREVIEW_LIMIT = 1000
DEFAULT_SAFETY_DAYS = 10
DEFAULT_SALE_DAYS = 30
RTV, AGE = "RTV Report", "Stock Age"

st.set_page_config(page_title="RTV Report", layout="wide")
st.markdown(
    """
    <style>
        .block-container {padding-top: 2.2rem;}
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------
# Cached work. Streamlit re-runs this script on every click, so file reading
# is cached by file content and only repeats when a file actually changes.
# --------------------------------------------------------------------------

@st.cache_data(show_spinner="Reading file...", max_entries=3)
def load_movement(data, name, mode):
    try:
        return logic.process_movement(logic.read_any_file(io.BytesIO(data), name), mode)
    except Exception as exc:
        return {"ok": False, "error": f"Could not read {name}: {exc}"}


@st.cache_data(show_spinner="Reading file...", max_entries=60)
def load_stock(data, name):
    try:
        return logic.process_stock(logic.read_any_file(io.BytesIO(data), name))
    except Exception as exc:
        return {"ok": False, "error": f"could not be read ({exc})"}


@st.cache_data(show_spinner="Reading file...", max_entries=2)
def load_sale(data, name):
    try:
        return logic.process_sale(logic.read_sale_file(io.BytesIO(data), name))
    except Exception as exc:
        return {"ok": False, "error": f"Could not read {name}: {exc}"}


@st.cache_data(show_spinner=False, max_entries=2)
def make_excel(report, sheet_name):
    return logic.build_excel_bytes(report, sheet_name)


@st.cache_data(show_spinner=False)
def make_templates():
    return logic.build_template_bytes()


# --------------------------------------------------------------------------
# Store mapping list (used for per-store files and for a merged file)
# --------------------------------------------------------------------------

def map_to_stores(items, stores, allow_skip, run_id):
    """
    items: list of (id, label, guess). Shows each item with a store dropdown
    in a scrollable box, pending items first.
    Returns {id: store name, SKIP or BLANK}.
    """
    store_set = set(stores)
    valid = store_set | ({SKIP} if allow_skip else set())
    taken, entries = set(), []

    for item_id, label, guess in items:
        key = f"map_{run_id}_{item_id}"
        chosen = st.session_state[key] if key in st.session_state else (guess or BLANK)
        if chosen != BLANK and (chosen not in valid or chosen in taken):
            chosen = BLANK
        st.session_state[key] = chosen
        if chosen in store_set:
            taken.add(chosen)
        entries.append({"id": item_id, "label": label, "key": key, "chosen": chosen})

    entries.sort(key=lambda e: e["chosen"] != BLANK)  # pending first

    done = sum(e["chosen"] != BLANK for e in entries)
    st.caption(f"{done} assigned, {len(entries) - done} pending")
    with st.container(height=min(60 + 52 * len(entries), 420), border=True):
        for e in entries:
            c_name, c_pick, c_state = st.columns([3, 3, 1.2], vertical_alignment="center")
            c_name.write(e["label"])
            options = [BLANK] + ([SKIP] if allow_skip else [])
            options += [s for s in stores if s == e["chosen"] or s not in taken]
            c_pick.selectbox("Store", options, key=e["key"], label_visibility="collapsed")
            if e["chosen"] == BLANK:
                c_state.markdown(":orange[Pending]")
            elif e["chosen"] == SKIP:
                c_state.markdown(":gray[Skipped]")
            else:
                c_state.markdown(":green[Assigned]")

    return {e["id"]: e["chosen"] for e in entries}


# --------------------------------------------------------------------------
# Sidebar
# --------------------------------------------------------------------------

st.session_state.setdefault("run_id", 0)
run_id = st.session_state.run_id

with st.sidebar:
    st.markdown("### RTV Report")
    report_type = st.radio("Report", [RTV, AGE], horizontal=True)
    mode = st.radio("Movement file", list(logic.MODES), horizontal=True)
    settings_box = st.container()
    st.divider()
    st.download_button("Templates", make_templates(), "RTV_Upload_Templates.xlsx", EXCEL_MIME)
    if st.button("Reset"):
        st.session_state.run_id += 1
        st.cache_data.clear()
        st.rerun()

cfg = logic.MODES[mode]
use_sale = report_type == RTV


# --------------------------------------------------------------------------
# Title, guide, uploads
# --------------------------------------------------------------------------

st.title("RTV Report")

guide_path = Path(__file__).with_name("Upload_Guide.txt")
with st.expander("Guide"):
    if guide_path.exists():
        st.text(guide_path.read_text(encoding="utf-8"))
    else:
        st.write("Upload_Guide.txt not found.")

boxes = st.columns(3 if use_sale else 2)
with boxes[0].container(border=True):
    st.markdown(f"**1. {mode} file**")
    move_file = st.file_uploader(
        f"{mode} file", type=["xlsx", "csv"], key=f"move_{mode}_{run_id}", label_visibility="collapsed"
    )
with boxes[1].container(border=True):
    st.markdown("**2. Current Stock**")
    stock_files = st.file_uploader(
        "Current Stock",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"stock_{run_id}",
        label_visibility="collapsed",
    )
sale_file = None
if use_sale:
    with boxes[2].container(border=True):
        st.markdown("**3. Sale (optional)**")
        sale_file = st.file_uploader(
            "Sale file", type=["xlsx", "csv"], key=f"sale_{run_id}", label_visibility="collapsed"
        )

if move_file is None:
    st.info(f"Upload the {mode} file.")
    st.stop()

movement = load_movement(move_file.getvalue(), move_file.name, mode)
if not movement["ok"]:
    st.error(movement["error"])
    st.stop()

stores = movement["stores"]
history = movement["history_days"]
last_date = movement["max_date"].date()

with settings_box:
    st.markdown("**Settings**")
    report_date = st.date_input(
        "Report date",
        value=max(date.today(), last_date),
        min_value=last_date,
        format="DD/MM/YYYY",
    )
    st.caption(f"{mode} history: {history} days")

if not stock_files:
    st.caption("Upload the Current Stock file(s).")
    st.stop()


# --------------------------------------------------------------------------
# Stores
# --------------------------------------------------------------------------

st.subheader("Stores")

loaded = [(f, load_stock(f.getvalue(), f.name)) for f in stock_files]
bad = [f"{f.name}: {r['error']}" for f, r in loaded if not r["ok"]]
for message in bad:
    st.error(message)
if bad:
    st.stop()

merged_mode = any(r["merged"] for _, r in loaded)
if merged_mode and len(loaded) > 1:
    st.error("Upload either one file with a Pos column, or one file per store without it. Not both.")
    st.stop()

frames, assigned_stores, pending = [], set(), 0

if merged_mode:
    _, result = loaded[0]
    data = result["data"]
    found = logic.source_stores(data)
    items = [(key, name, logic.match_store(name, stores)) for key, name in found]
    choice = map_to_stores(items, stores, allow_skip=True, run_id=run_id)
    pending = sum(c == BLANK for c in choice.values())
    mapping = {k: v for k, v in choice.items() if v not in (BLANK, SKIP)}
    assigned_stores = set(mapping.values())
    if mapping:
        frames.append(logic.attach_stores(data, mapping))
    if result["blank_rows"]:
        st.caption(f"{result['blank_rows']} row(s) without a store left out.")
else:
    items = []
    for i, (f, _) in enumerate(loaded):
        file_id = getattr(f, "file_id", None) or f"{i}_{f.name}"
        items.append((file_id, f.name, logic.suggest_store(f.name, stores)))
    choice = map_to_stores(items, stores, allow_skip=False, run_id=run_id)
    pending = sum(c == BLANK for c in choice.values())
    for (f, result), (file_id, _, _) in zip(loaded, items):
        if choice[file_id] != BLANK:
            frames.append(logic.attach_store(result["data"], choice[file_id]))
            assigned_stores.add(choice[file_id])

missing = [s for s in stores if s not in assigned_stores]
if missing:
    st.warning("No stock for: " + ", ".join(missing))


# --------------------------------------------------------------------------
# Sale file (RTV only)
# --------------------------------------------------------------------------

sold, sale_error = None, None
if sale_file is not None:
    sale = load_sale(sale_file.getvalue(), sale_file.name)
    if not sale["ok"]:
        sale_error = sale["error"]
    else:
        sold, unmatched, n_articles, n_items = logic.match_sales(sale["data"], stores)
        st.caption(f"Sale file: {n_articles} article(s) ({n_items} store/color/size items) selling are kept out of RTV.")
        if unmatched:
            st.warning("Sale stores not matched: " + ", ".join(unmatched))


# --------------------------------------------------------------------------
# Reports
# --------------------------------------------------------------------------

def show_table(df):
    preview = df.head(PREVIEW_LIMIT).copy()
    for col in preview.columns:
        if pd.api.types.is_datetime64_any_dtype(preview[col]):
            preview[col] = preview[col].dt.strftime("%d-%m-%Y").fillna("")
    if "Age (Days)" in preview.columns:
        preview["Age (Days)"] = preview["Age (Days)"].astype(str)
    st.dataframe(preview, hide_index=True, height=420)
    if len(df) > PREVIEW_LIMIT:
        st.caption(f"Showing the first {PREVIEW_LIMIT:,} of {len(df):,} rows. Excel has all rows.")


def show_rtv(stock_all):
    default_age = max(history - (DEFAULT_SAFETY_DAYS if cfg["uses_safety"] else 0), 1)
    cols = st.columns(4 if cfg["uses_safety"] else 3)
    rtv_age = cols[0].number_input("RTV age (days)", min_value=1, value=int(default_age), step=1)
    safety_days, nxt = 0, 1
    if cfg["uses_safety"]:
        safety_days = cols[1].number_input("Safety days", min_value=0, value=DEFAULT_SAFETY_DAYS, step=1)
        nxt = 2
    sale_days = cols[nxt].number_input(
        "Sale cutoff days", min_value=1, value=DEFAULT_SALE_DAYS, step=1,
        help="Reminder only. The Sale file should cover this many days.",
    )
    cutoff = logic.cutoff_days(rtv_age, safety_days)
    cols[-1].metric("Cutoff", f"{cutoff} days")
    if sale_file is not None:
        st.caption(f"Reminder: the Sale file should cover the last {sale_days} days.")

    if sale_error:
        st.error(sale_error)
        return
    if cutoff > history:
        st.error(f"Not enough history: file has {history} days, cutoff is {cutoff}.")
        return

    report, excluded = logic.build_report(stock_all, movement["summary"], report_date, cutoff, mode, sold=sold)
    if report.empty:
        st.info("Nothing to return.")
        if excluded:
            st.caption(f"{excluded:,} item(s) kept out because they are selling.")
        return

    widths = [1, 1, 1, 1, 1.3] if sale_file is not None else [1, 1, 1, 1.3]
    m = st.columns(widths, vertical_alignment="bottom")
    m[0].metric("Stores", report["Store Name"].nunique())
    m[1].metric("Records", f"{len(report):,}")
    m[2].metric("RTV qty", f"{report['RTV Qty'].sum():,.0f}")
    if sale_file is not None:
        m[3].metric("Excluded (selling)", f"{excluded:,}")
    m[-1].download_button(
        "Download Excel", make_excel(report, "RTV_Report"), "RTV_Report.xlsx",
        EXCEL_MIME, type="primary", key="download_rtv",
    )
    show_table(report)


def show_age(stock_all):
    report = logic.build_age_report(stock_all, movement["summary"], report_date, history, mode)
    if report.empty:
        st.info("No stock found.")
        return

    m1, m2, m3, m4 = st.columns([1, 1, 1, 1.3], vertical_alignment="bottom")
    m1.metric("Stores", report["Store Name"].nunique())
    m2.metric("Records", f"{len(report):,}")
    m3.metric("Stock qty", f"{report['Current Stock'].sum():,.0f}")
    m4.download_button(
        "Download Excel", make_excel(report, "Stock_Age"), "Stock_Age_Report.xlsx",
        EXCEL_MIME, type="primary", key="download_age",
    )
    st.caption(f"{history}+ = not in the {mode.lower()} file.")
    show_table(report)


st.subheader(report_type)

if pending:
    st.caption("Finish the store mapping to see the report.")
    st.stop()
if not frames:
    st.caption("No stock is assigned to a store.")
    st.stop()

stock_all = pd.concat(frames, ignore_index=True)
if report_type == RTV:
    show_rtv(stock_all)
else:
    show_age(stock_all)