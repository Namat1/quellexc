import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

st.set_page_config(page_title="Transportgruppen Editor", layout="wide")

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background: #f5f7fa; }
    .block-container { padding: 1.5rem 2rem; }
    h1 { color: #1F4E79; font-size: 1.6rem; }
    h3 { color: #1F4E79; font-size: 1.1rem; margin-bottom: 0.3rem; }

    .kunde-card {
        background: white;
        border-radius: 10px;
        padding: 1rem 1.4rem;
        margin-bottom: 1rem;
        border-left: 5px solid #1F4E79;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .kunde-info { color: #444; font-size: 0.88rem; margin-bottom: 0.6rem; }

    .tag-header {
        font-weight: 700;
        font-size: 0.85rem;
        padding: 4px 8px;
        border-radius: 5px;
        text-align: center;
        margin-bottom: 4px;
    }
    .slot-box {
        background: #f0f4f8;
        border-radius: 6px;
        padding: 6px 8px;
        margin-bottom: 6px;
        font-size: 0.8rem;
    }
    .slot-empty { color: #aaa; font-style: italic; }

    div[data-testid="stSelectbox"] label { font-size: 0.8rem; color: #555; }
    div[data-testid="stTextInput"] label { font-size: 0.8rem; color: #555; }

    .stButton > button {
        background: #1F4E79;
        color: white;
        border-radius: 6px;
        border: none;
        padding: 0.4rem 1.2rem;
        font-size: 0.88rem;
    }
    .stButton > button:hover { background: #2e6da4; }

    .tag-mo  { background:#FFF2CC; color:#7D6608; }
    .tag-die { background:#E2EFDA; color:#1E5C25; }
    .tag-mit { background:#DDEBF7; color:#1F4E79; }
    .tag-don { background:#FCE4D6; color:#843C0C; }
    .tag-fre { background:#EAD1DC; color:#6D1A36; }
    .tag-sam { background:#D9D2E9; color:#351C75; }

    .pill {
        display:inline-block;
        padding:2px 8px;
        border-radius:10px;
        font-size:0.75rem;
        font-weight:600;
        margin-right:3px;
    }
</style>
""", unsafe_allow_html=True)

TAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
TAG_CSS = ["tag-mo", "tag-die", "tag-mit", "tag-don", "tag-fre", "tag-sam"]
TAG_SHORT = ["Mo", "Di", "Mi", "Do", "Fr", "Sa"]

ZEITEN = ["", "06:00 Uhr", "07:00 Uhr", "08:00 Uhr", "09:00 Uhr", "10:00 Uhr",
          "11:00 Uhr", "12:00 Uhr", "13:00 Uhr", "14:00 Uhr", "15:00 Uhr",
          "16:00 Uhr", "17:00 Uhr", "18:00 Uhr", "19:00 Uhr", "20:00 Uhr", "21:00 Uhr"]

FIXED_START = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort", "Mo", "Die", "Mitt", "Don", "Fr", "Sam"]
FIXED_END   = ["Tag Mo", "Tag Die", "Tag Mitt", "Tag Don", "Tag Fr", "Tag Sa", "Fax", "Fachberater", "Ostkunden"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def clean(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()

def parse_slot_columns(df_cols):
    """Gibt Liste von Gruppen zurück: [(idx_zeit, idx_sort, idx_tag), ...]"""
    groups = []
    i = 0
    cols = list(df_cols)
    while i < len(cols):
        c = cols[i]
        if re.match(r'.+_Zeit$', c):
            if i+2 < len(cols) and re.match(r'.+_Sort$', cols[i+1]) and re.match(r'.+_Tag$', cols[i+2]):
                groups.append((i, i+1, i+2))
                i += 3
                continue
        i += 1
    return groups

def load_excel(file):
    xls = pd.ExcelFile(file)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name, dtype=str)
        df = df.fillna("")
        sheets[name] = df
    return sheets

def get_day_of_group(col_name):
    """Aus 'Montag_Zeit' → 'Montag'"""
    for tag in TAGE:
        if col_name.startswith(tag):
            return tag
    return ""

def row_to_slots(row, groups, df_cols):
    """Gibt Dict: { 'Montag': [(zeit, sort, liefertag), ...], ... }"""
    result = {t: [] for t in TAGE}
    cols = list(df_cols)
    for (iz, is_, it) in groups:
        tag = get_day_of_group(cols[iz])
        if tag:
            zeit    = clean(row.iloc[iz])
            sort_   = clean(row.iloc[is_])
            ltag    = clean(row.iloc[it])
            result[tag].append([zeit, sort_, ltag])
    return result

def slots_to_row(slots, groups, df_cols, original_row):
    """Schreibt editierte Slots zurück in eine Zeile"""
    row = original_row.copy()
    cols = list(df_cols)
    # Pro Tag: Zähler für welchen Slot wir gerade befüllen
    tag_counter = {t: 0 for t in TAGE}
    for (iz, is_, it) in groups:
        tag = get_day_of_group(cols[iz])
        if tag:
            idx = tag_counter[tag]
            slot_list = slots.get(tag, [])
            if idx < len(slot_list):
                row.iloc[iz] = slot_list[idx][0]
                row.iloc[is_] = slot_list[idx][1]
                row.iloc[it]  = slot_list[idx][2]
            else:
                row.iloc[iz] = ""
                row.iloc[is_] = ""
                row.iloc[it]  = ""
            tag_counter[tag] += 1
    return row

def df_to_excel_bytes(sheets_dict):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()

# ── Session State ─────────────────────────────────────────────────────────────
if "sheets" not in st.session_state:
    st.session_state.sheets = {}
if "edits" not in st.session_state:
    st.session_state.edits = {}   # {(sheet, row_idx): modified_row}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Datei laden")
    uploaded = st.file_uploader("Transportgruppen Excel hochladen", type=["xlsx", "xls"])
    if uploaded:
        st.session_state.sheets = load_excel(uploaded)
        st.success(f"{len(st.session_state.sheets)} Blätter geladen")

    if st.session_state.sheets:
        sheet_name = st.selectbox("📋 Blatt auswählen", list(st.session_state.sheets.keys()))
    else:
        sheet_name = None

    st.markdown("---")
    if st.session_state.sheets and sheet_name:
        df = st.session_state.sheets[sheet_name]
        search = st.text_input("🔍 Suche (Name, Nr, Ort)", "")
        st.markdown(f"**{len(df)} Kunden** auf diesem Blatt")

    st.markdown("---")
    if st.session_state.sheets:
        st.markdown("### 💾 Export")
        if st.button("Excel herunterladen"):
            # Edits zurückschreiben
            final = {}
            for sname, df in st.session_state.sheets.items():
                df2 = df.copy()
                for (s, ridx), row in st.session_state.edits.items():
                    if s == sname:
                        df2.iloc[ridx] = row
                final[sname] = df2
            xls_bytes = df_to_excel_bytes(final)
            st.download_button(
                "⬇️ Download",
                data=xls_bytes,
                file_name="Transportgruppen_editiert.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("# 🚛 Transportgruppen Editor")

if not st.session_state.sheets:
    st.info("👈 Bitte links die Excel-Datei hochladen.")
    st.stop()

df_orig = st.session_state.sheets[sheet_name]
df = df_orig.copy()

# Apply saved edits to display
for (s, ridx), row in st.session_state.edits.items():
    if s == sheet_name:
        df.iloc[ridx] = row

# Find slot column groups
groups = parse_slot_columns(df.columns)

# Filter rows
mask = pd.Series([True] * len(df))
if search:
    q = search.lower()
    mask = df.apply(lambda row: any(q in str(v).lower() for v in row[:10]), axis=1)

filtered_idx = df[mask].index.tolist()
st.markdown(f"**{len(filtered_idx)} Kunden** gefunden")

if not filtered_idx:
    st.warning("Keine Kunden gefunden.")
    st.stop()

# ── Pagination ────────────────────────────────────────────────────────────────
PAGE_SIZE = 10
total_pages = max(1, (len(filtered_idx) + PAGE_SIZE - 1) // PAGE_SIZE)
col_pg1, col_pg2, col_pg3 = st.columns([1, 2, 1])
with col_pg1:
    if st.button("◀ Zurück"):
        st.session_state["page"] = max(0, st.session_state.get("page", 0) - 1)
with col_pg3:
    if st.button("Weiter ▶"):
        st.session_state["page"] = min(total_pages - 1, st.session_state.get("page", 0) + 1)
with col_pg2:
    page = st.session_state.get("page", 0)
    st.markdown(f"<div style='text-align:center;padding-top:6px'>Seite {page+1} / {total_pages}</div>", unsafe_allow_html=True)

page_idx = filtered_idx[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]

# ── Render each customer ──────────────────────────────────────────────────────
for ridx in page_idx:
    row = df.iloc[ridx]
    nr      = clean(row.get("Nr", ""))
    sap     = clean(row.get("SAP-Nr.", ""))
    name    = clean(row.get("Name", ""))
    strasse = clean(row.get("Strasse", ""))
    plz     = clean(row.get("Plz", ""))
    ort     = clean(row.get("Ort", ""))
    fax     = clean(row.get("Fax", ""))
    berater = clean(row.get("Fachberater", ""))

    slots = row_to_slots(row, groups, df.columns)

    with st.expander(f"**{nr}** · {name.strip()} · {plz} {ort.strip()}", expanded=False):
        st.markdown(f"""
        <div class='kunde-info'>
            <b>SAP:</b> {sap} &nbsp;|&nbsp;
            <b>Adresse:</b> {strasse.strip()}, {plz} {ort.strip()} &nbsp;|&nbsp;
            <b>Fax:</b> {fax} &nbsp;|&nbsp;
            <b>Fachberater:</b> {berater}
        </div>
        """, unsafe_allow_html=True)

        # Show day columns
        day_cols = st.columns(6)
        edited_slots = {}

        for d_idx, (tag, css) in enumerate(zip(TAGE, TAG_CSS)):
            with day_cols[d_idx]:
                st.markdown(f"<div class='tag-header {css}'>{tag}</div>", unsafe_allow_html=True)
                tag_slots = slots.get(tag, [])
                # Filter non-empty slots
                active = [s for s in tag_slots if s[0] or s[1]]
                empty_count = len(tag_slots) - len(active)

                edited_day = []
                for s_idx, (zeit, sort_, ltag) in enumerate(tag_slots):
                    if not (zeit or sort_):
                        edited_day.append(["", "", ""])
                        continue
                    uid = f"{sheet_name}_{ridx}_{tag}_{s_idx}"
                    new_zeit = st.selectbox(
                        f"Zeit #{s_idx+1}",
                        ZEITEN,
                        index=ZEITEN.index(zeit) if zeit in ZEITEN else 0,
                        key=f"z_{uid}"
                    )
                    new_sort = st.text_input(
                        f"Sortiment #{s_idx+1}",
                        value=sort_,
                        key=f"s_{uid}"
                    )
                    new_ltag = st.selectbox(
                        f"Liefertag #{s_idx+1}",
                        [""] + TAGE,
                        index=([""] + TAGE).index(ltag) if ltag in TAGE else 0,
                        key=f"t_{uid}"
                    )
                    edited_day.append([new_zeit, new_sort, new_ltag])

                edited_slots[tag] = edited_day

        # Save button per customer
        col_save, col_reset = st.columns([1, 1])
        with col_save:
            if st.button(f"💾 Speichern", key=f"save_{sheet_name}_{ridx}"):
                new_row = slots_to_row(edited_slots, groups, df.columns, df.iloc[ridx])
                st.session_state.edits[(sheet_name, ridx)] = new_row
                st.success("Gespeichert! Export über die Sidebar.")
        with col_reset:
            if st.button(f"↺ Zurücksetzen", key=f"reset_{sheet_name}_{ridx}"):
                key = (sheet_name, ridx)
                if key in st.session_state.edits:
                    del st.session_state.edits[key]
                st.rerun()

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
edits_count = sum(1 for (s, _) in st.session_state.edits if s == sheet_name)
if edits_count:
    st.markdown(f"✏️ **{edits_count} Kunden** auf diesem Blatt wurden geändert (noch nicht exportiert)")
