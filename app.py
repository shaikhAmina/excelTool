import streamlit as st

# ── Page config (must be first Streamlit call) ───────
st.set_page_config(
    page_title="Excel Toolbox",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Tool UI imports ───────────────────────────────────
from utils.file_utils import merge_excel_ui, split_excel_ui, find_duplicates_ui
from utils.vlookup_tool import vlookup_tool_ui
from utils.converters import convert_to_number_ui, comma_string_formatter_ui
from utils.json_excel_tools import json_excel_tool_ui

# ─────────────────────────────────────────────────────
# TOOL REGISTRY
# ─────────────────────────────────────────────────────
TOOLS = [
    {
        "id": "merge",
        "label": "Merge Excel",
        "icon": "🧩",
        "category": "Excel Tools",
        "desc": "Combine multiple Excel files — or sheets — into one spreadsheet.",
        "fn": merge_excel_ui,
    },
    {
        "id": "split",
        "label": "Split Excel",
        "icon": "✂️",
        "category": "Excel Tools",
        "desc": "Break a large file into smaller chunks and download them as a ZIP.",
        "fn": split_excel_ui,
    },
    {
        "id": "vlookup",
        "label": "VLOOKUP",
        "icon": "🔗",
        "category": "Excel Tools",
        "desc": "Match and pull columns across two spreadsheets by a shared key.",
        "fn": vlookup_tool_ui,
    },
    {
        "id": "duplicates",
        "label": "Find Duplicates",
        "icon": "🔍",
        "category": "Data Cleaning",
        "desc": "Detect duplicate rows and download a yellow-highlighted Excel file.",
        "fn": find_duplicates_ui,
    },
    {
        "id": "json_excel",
        "label": "JSON ↔ Excel",
        "icon": "🔄",
        "category": "Conversion Tools",
        "desc": "Convert JSON to Excel, Excel to JSON, or pull live API data into a sheet.",
        "fn": json_excel_tool_ui,
    },
    {
        "id": "number_fmt",
        "label": "Number Formatter",
        "icon": "🔢",
        "category": "Data Cleaning",
        "desc": "Convert text-stored numbers to real numeric values so formulas work correctly.",
        "fn": convert_to_number_ui,
    },
    {
        "id": "string_fmt",
        "label": "String Formatter",
        "icon": "🧩",
        "category": "Conversion Tools",
        "desc": "Format a list of values as comma-separated, quoted, or SQL IN-clause strings.",
        "fn": comma_string_formatter_ui,
    },
]

CATEGORIES = ["Excel Tools", "Conversion Tools", "Data Cleaning"]

BADGE_CLASS = {
    "Excel Tools":      "badge-excel",
    "Conversion Tools": "badge-convert",
    "Data Cleaning":    "badge-clean",
}

# ─────────────────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────────────────
st.markdown(
    """
    <!-- Google Search Console verification -->
    <meta name="google-site-verification" content="u4O-8Kcd_N9jIufg-9dli6KzYCl9ykB80BgsLpTENW8" />

    <!-- SEO meta tags -->
    <meta name="description" content="Free online Excel tools — merge, split, VLOOKUP, find duplicates, convert JSON to Excel, format numbers and strings. No sign-up required." />
    <meta name="keywords" content="excel merge tool, split excel, vlookup online, find duplicates excel, json to excel, excel to json, free excel tools, data cleaning" />
    <meta property="og:title" content="Excel Toolbox — Free Excel & Data Tools" />
    <meta property="og:description" content="Merge, split, VLOOKUP, find duplicates, convert JSON to Excel and more. Free, no sign-up, files never leave your browser." />
    <meta property="og:url" content="https://exceltoolbox.streamlit.app/" />
    <meta property="og:type" content="website" />
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    #MainMenu, footer, header { visibility: hidden; }

    .tb-navbar {
        display: flex; align-items: center; gap: 12px;
        padding: 14px 0 10px 0;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 28px;
    }
    .tb-logo { font-size: 1.55rem; font-weight: 800; color: #1a7f5a; margin: 0; }
    .tb-logo span { color: #2d3748; }
    .tb-tagline { font-size: 0.85rem; color: #718096; margin-left: 4px; }

    .tb-badge {
        display: inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: 0.72rem; font-weight: 600; letter-spacing: 0.02em; margin-bottom: 6px;
    }
    .badge-excel   { background: #ebf8f1; color: #276749; }
    .badge-convert { background: #ebf4ff; color: #2b6cb0; }
    .badge-clean   { background: #fef3c7; color: #92400e; }

    .tb-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 20px 18px 16px 18px; height: 100%;
        transition: box-shadow 0.18s, border-color 0.18s;
        display: flex; flex-direction: column; gap: 8px;
    }
    .tb-card:hover { box-shadow: 0 4px 18px rgba(0,0,0,0.09); border-color: #a0aec0; }
    .tb-card-icon  { font-size: 1.8rem; line-height: 1; }
    .tb-card-title { font-size: 1.05rem; font-weight: 700; color: #1a202c; margin: 0; }
    .tb-card-desc  { font-size: 0.85rem; color: #4a5568; flex: 1; margin: 0; }

    .tb-section {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 0.08em;
        text-transform: uppercase; color: #718096;
        margin: 28px 0 12px 0; padding-bottom: 4px;
        border-bottom: 1px solid #e2e8f0;
    }

    .tb-breadcrumb { font-size: 0.82rem; color: #718096; margin-bottom: 6px; }
    .tb-breadcrumb b { color: #1a202c; }

    div[data-testid="stTextInput"] > div > div > input {
        border-radius: 8px !important; border: 1.5px solid #cbd5e0 !important;
        font-size: 0.95rem !important; padding: 10px 14px !important;
    }

    h4 { color: #2d3748 !important; margin-top: 22px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────
if "active_tool" not in st.session_state:
    st.session_state.active_tool = None   # None = homepage


def _go_home():
    st.session_state.active_tool = None


def _open_tool(tool_id: str):
    st.session_state.active_tool = tool_id


# ─────────────────────────────────────────────────────
# GRID RENDERER
# ─────────────────────────────────────────────────────
def _render_grid(tools: list):
    """Render tool cards in rows of 4."""
    num = len(tools)
    if num == 0:
        return
    # Work in rows of 4
    for row_start in range(0, num, 4):
        row_tools = tools[row_start: row_start + 4]
        cols = st.columns(len(row_tools))
        for col, tool in zip(cols, row_tools):
            badge_cls = BADGE_CLASS.get(tool["category"], "badge-excel")
            with col:
                st.markdown(
                    f"""
                    <div class="tb-card">
                        <div class="tb-card-icon">{tool["icon"]}</div>
                        <span class="tb-badge {badge_cls}">{tool["category"]}</span>
                        <p class="tb-card-title">{tool["label"]}</p>
                        <p class="tb-card-desc">{tool["desc"]}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Open Tool →", key=f"open_{tool['id']}_{row_start}"):
                    _open_tool(tool["id"])
                    st.rerun()


# ─────────────────────────────────────────────────────
# TOP NAV  (always visible)
# ─────────────────────────────────────────────────────
nav_left, nav_right = st.columns([6, 1])
with nav_left:
    st.markdown(
        '<div class="tb-navbar">'
        '<p class="tb-logo">🧰 Excel<span>Toolbox</span></p>'
        '<span class="tb-tagline">'
        "Free tools for developers, analysts, and everyone who works with data"
        "</span>"
        "</div>",
        unsafe_allow_html=True,
    )
with nav_right:
    if st.session_state.active_tool is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🏠 Home", key="nav_home_btn"):
            _go_home()
            st.rerun()

# ─────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────
active = st.session_state.active_tool

# ══════════════════ HOMEPAGE ══════════════════════════
if active is None:

    # Hero
    st.markdown(
        """
        <div style="text-align:center; padding: 8px 0 32px 0;">
            <h1 style="font-size:2.2rem; font-weight:800; color:#1a202c; margin-bottom:10px;">
                Simple tools for your Excel & data problems
            </h1>
            <p style="font-size:1.05rem; color:#4a5568; max-width:560px; margin:0 auto;">
                No sign-up. No upload limits. Files never leave your browser session.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Search bar
    search_col, _ = st.columns([2, 3])
    with search_col:
        query = st.text_input(
            "Search tools",
            placeholder="🔍  Search tools — try 'merge', 'json', 'duplicate'…",
            label_visibility="collapsed",
            key="search_query",
        )

    # Filter
    def _matches(tool: dict, q: str) -> bool:
        q = q.lower().strip()
        return (
            q in tool["label"].lower()
            or q in tool["desc"].lower()
            or q in tool["category"].lower()
            or q in tool["id"].lower()
        )

    visible = [t for t in TOOLS if _matches(t, query)] if query.strip() else TOOLS

    if query.strip() and not visible:
        st.warning(
            f"No tools match **\"{query}\"**. "
            "Try searching for 'merge', 'split', 'json', 'duplicate', or 'number'."
        )

    elif query.strip():
        st.markdown(
            f'<div class="tb-section">Search results for &ldquo;{query}&rdquo;</div>',
            unsafe_allow_html=True,
        )
        _render_grid(visible)

    else:
        for cat in CATEGORIES:
            cat_tools = [t for t in visible if t["category"] == cat]
            if not cat_tools:
                continue
            st.markdown(f'<div class="tb-section">{cat}</div>', unsafe_allow_html=True)
            _render_grid(cat_tools)

    # Privacy notice
    st.markdown("---")
    st.markdown(
        """
        <div style="
            background:#f7fafc; border:1px solid #e2e8f0; border-radius:10px;
            padding:14px 18px; font-size:0.87rem; color:#4a5568;
        ">
            🔒 <strong>Privacy</strong> — Every file you upload is processed entirely within your
            active browser session using in-memory computation. Nothing is written to disk or
            sent to any server. All data is automatically discarded when you close or refresh the tab.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Feedback
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("💬 Need another Excel tool? Tell us!", expanded=False):
        idea = st.text_area(
            "Describe the tool you need",
            placeholder="e.g. 'A tool that removes blank rows and trims whitespace from all cells'",
            height=100,
            key="home_feedback_text",
        )
        if st.button("Submit idea", key="home_feedback_btn"):
            if idea.strip():
                from utils.feedback import save_feedback
                saved = save_feedback(idea)
                if saved:
                    st.success("Thanks! Your idea has been saved. 🙌")
                else:
                    st.success("Thanks for the suggestion! We'll consider it for the next release. 🙌")
            else:
                st.warning("Please describe your idea before submitting.")

# ══════════════════ TOOL VIEW ═════════════════════════
else:
    tool_map = {t["id"]: t for t in TOOLS}
    tool = tool_map.get(active)

    if tool is None:
        st.error("Tool not found. Returning to homepage…")
        _go_home()
        st.rerun()

    # Breadcrumb
    st.markdown(
        f'<div class="tb-breadcrumb">'
        f'🏠 Home &rsaquo; {tool["category"]} &rsaquo; '
        f'<b>{tool["icon"]} {tool["label"]}</b>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Render the tool
    tool["fn"]()
