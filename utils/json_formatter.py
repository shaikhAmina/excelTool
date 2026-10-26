"""
JSON Formatter + Minifier
─────────────────────────
Paste → Format → collapsible Postman-style tree using native HTML <details>.
Click any arrow to collapse/expand any object or array block.
"""

import streamlit as st
import streamlit.components.v1 as components
import ujson as json
import json as _stdlib_json
import html as _html
import pandas as pd

from utils.helpers import privacy_notice, feedback_widget, df_to_excel_bytes


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _size_label(s: str) -> str:
    b = len(s.encode("utf-8"))
    if b < 1024:
        return f"{b} B"
    if b < 1_048_576:
        return f"{b / 1024:.1f} KB"
    return f"{b / 1_048_576:.2f} MB"


def _to_excel_bytes(obj) -> bytes | None:
    try:
        if isinstance(obj, dict):
            list_val = next((v for v in obj.values() if isinstance(v, list)), None)
            data = list_val if list_val is not None else [obj]
        elif isinstance(obj, list):
            data = obj
        else:
            return None
        df = pd.json_normalize(data)
        return df_to_excel_bytes(df, sheet_name="Data")
    except Exception:
        return None


# ──────────────────────────────────────────────────────────────────────────────
# Pure-HTML collapsible tree builder  (<details>/<summary> — no JS needed)
# ──────────────────────────────────────────────────────────────────────────────

def _val(v) -> str:
    """Render a scalar value as a coloured span."""
    if v is None:
        return '<span class="null">null</span>'
    if isinstance(v, bool):
        return f'<span class="bool">{"true" if v else "false"}</span>'
    if isinstance(v, (int, float)):
        return f'<span class="num">{v}</span>'
    if isinstance(v, str):
        return f'<span class="str">{_html.escape(_stdlib_json.dumps(v))}</span>'
    return _html.escape(str(v))


def _build(node, comma: str = "") -> str:
    """Recursively build HTML for any JSON node."""
    if isinstance(node, dict):
        if not node:
            return f'<span class="brace">{{}}</span><span class="punct">{comma}</span>'
        count = len(node)
        hint = f"{count} key{'s' if count != 1 else ''}"
        inner = ""
        items = list(node.items())
        for i, (k, v) in enumerate(items):
            c = "," if i < len(items) - 1 else ""
            key_html = f'<span class="key">{_html.escape(_stdlib_json.dumps(k))}</span><span class="punct">: </span>'
            inner += f'<div class="line">{key_html}{_build(v, c)}</div>'
        return (
            f'<details open>'
            f'<summary><span class="brace">{{</span>'
            f'<span class="hint"> {hint} </span>'
            f'<span class="brace ellipsis">…}}</span></summary>'
            f'<div class="block">{inner}</div>'
            f'<div class="line"><span class="brace">}}</span>'
            f'<span class="punct">{comma}</span></div>'
            f'</details>'
        )

    if isinstance(node, list):
        if not node:
            return f'<span class="brace">[]</span><span class="punct">{comma}</span>'
        count = len(node)
        hint = f"{count} item{'s' if count != 1 else ''}"
        inner = ""
        for i, item in enumerate(node):
            c = "," if i < count - 1 else ""
            inner += f'<div class="line">{_build(item, c)}</div>'
        return (
            f'<details open>'
            f'<summary><span class="brace">[</span>'
            f'<span class="hint"> {hint} </span>'
            f'<span class="brace ellipsis">…]</span></summary>'
            f'<div class="block">{inner}</div>'
            f'<div class="line"><span class="brace">]</span>'
            f'<span class="punct">{comma}</span></div>'
            f'</details>'
        )

    return f'{_val(node)}<span class="punct">{comma}</span>'


_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  background: #1e1e1e;
  color: #d4d4d4;
  font-family: 'Fira Code', 'Cascadia Code', 'Consolas', 'Courier New', monospace;
  font-size: 13px;
  line-height: 1.75;
  padding: 16px 20px;
}

/* ── collapsible block ── */
details { display: inline-block; vertical-align: top; }
details[open] > summary .ellipsis { display: none; }
details:not([open]) > summary .hint { display: none; }
details:not([open]) > .block { display: none; }
details:not([open]) > div.line { display: none; }

summary {
  display: inline-block;   /* keep nested opening braces beside their keys */
  cursor: pointer;
  list-style: none;
  outline: none;
  white-space: nowrap;
}
summary::-webkit-details-marker { display: none; }

/* custom arrow via ::before */
summary::before {
  content: "▼ ";
  font-size: 10px;
  color: #888;
  display: inline-block;
  width: 16px;
}
details:not([open]) > summary::before { content: "▶ "; }
summary:hover::before { color: #ccc; }

/* indent children */
.block { padding-left: 22px; }
.line  { display: block; white-space: pre; }

/* colours */
.key   { color: #9cdcfe; }
.str   { color: #ce9178; }
.num   { color: #b5cea8; }
.bool  { color: #569cd6; }
.null  { color: #569cd6; }
.brace { color: #ffd700; font-weight: bold; }
.punct { color: #d4d4d4; }
.hint  { color: #888; font-size: 11px; font-style: italic; }
"""


def _render_tree(obj, height_px: int = 500):
    tree_html = _build(obj)
    full = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>{_CSS}</style>
</head><body>
<div class="line">{tree_html}</div>
</body></html>"""
    components.html(full, height=height_px, scrolling=True)


# ──────────────────────────────────────────────────────────────────────────────
# Main UI
# ──────────────────────────────────────────────────────────────────────────────

def json_formatter_ui():
    st.markdown("### 🎨 JSON Formatter & Minifier")
    st.caption("Paste any JSON → hit **Format JSON** → collapsible tree. Click ▶ to collapse any block.")
    privacy_notice()
    st.markdown("---")

    # ── Input ──────────────────────────────────────────────────────────────────
    raw = st.text_area(
        "Paste your JSON here",
        height=200,
        placeholder="Paste any JSON here — flat, nested, array of objects, deeply nested…",
        key="jf_raw",
        label_visibility="collapsed",
    )

    btn_col, opt_col1, opt_col2 = st.columns([2, 2, 2])
    with btn_col:
        go = st.button("▶️  Format JSON", type="primary", key="jf_go", use_container_width=True)
    with opt_col1:
        indent = st.select_slider("Indent", options=[2, 4], value=2, key="jf_indent", label_visibility="collapsed")
    with opt_col2:
        minify = st.checkbox("Minify instead", value=False, key="jf_minify")

    # Persist across reruns (download clicks rerun the page)
    if go and raw.strip():
        st.session_state["jf_committed"] = raw.strip()

    committed = st.session_state.get("jf_committed", "")

    if committed and committed != raw.strip() and raw.strip():
        st.info("✏️ Input changed — press **▶️ Format JSON** again to update.")

    if not committed:
        st.caption("👆 Paste your JSON above and press **▶️ Format JSON**.")
        feedback_widget()
        return

    # ── Parse ──────────────────────────────────────────────────────────────────
    try:
        obj = json.loads(committed)
    except Exception as e:
        st.error(
            f"⚠️ **Invalid JSON** — {str(e)[:180]}\n\n"
            "Common fixes: use double quotes, remove trailing commas, check all brackets are closed."
        )
        feedback_widget()
        return

    # ── Flat string for download ────────────────────────────────────────────────
    if minify:
        output = json.dumps(obj, ensure_ascii=False)
        label = "minified"
    else:
        output = _stdlib_json.dumps(obj, indent=indent, ensure_ascii=False)
        label = "formatted"

    lines = output.count("\n") + 1

    # ── Collapsible tree preview ───────────────────────────────────────────────
    st.markdown("#### Preview")
    st.caption(f"📄 {lines:,} lines · {_size_label(output)} · click ▶ / ▼ to collapse or expand any block")

    box_h = min(max(280, lines * 20), 660)
    _render_tree(obj, height_px=box_h)

    # ── Download ───────────────────────────────────────────────────────────────
    st.markdown("#### Download")
    fname = "output"

    dcol1, dcol2, dcol3 = st.columns(3)

    with dcol1:
        st.download_button(
            f"⬇️ {label.capitalize()} JSON",
            data=output,
            file_name=f"{fname}_{label}.json",
            mime="application/json",
            type="primary",
            key="jf_dl1",
        )

    with dcol2:
        other_label = "formatted" if minify else "minified"
        other_out = (
            _stdlib_json.dumps(obj, indent=indent, ensure_ascii=False)
            if minify
            else json.dumps(obj, ensure_ascii=False)
        )
        st.download_button(
            f"⬇️ {other_label.capitalize()} JSON",
            data=other_out,
            file_name=f"{fname}_{other_label}.json",
            mime="application/json",
            key="jf_dl2",
        )

    with dcol3:
        xl = _to_excel_bytes(obj)
        if xl:
            st.download_button(
                "⬇️ Download as Excel",
                data=xl,
                file_name=f"{fname}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="jf_dl3",
            )
        else:
            st.button(
                "⬇️ Download as Excel",
                disabled=True,
                key="jf_dl3_off",
                help="Excel export needs the JSON to be an array of objects.",
            )

    feedback_widget()
