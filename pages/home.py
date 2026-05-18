import io
import json
import re
import time
from typing import Any

import requests
from json_repair import repair_json
import streamlit as st
from docx import Document
from docxtpl import DocxTemplate
from pydantic import BaseModel, Field, ValidationError
from tenacity import retry, stop_after_attempt, wait_exponential

from app_config import PROMPT_FILE, STRAICO_ENDPOINT, STRAICO_MODEL, TEMPLATE_FILE
from auth import enforce_workspace_auth
from debug_store import add_run
from field_config import CHECKBOX_CHAR, UNCHECKED_CHAR, SECTIONS
from streamlit_quill import st_quill


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class BookFormData(BaseModel):
    title: str | None = ""
    subtitle: str | None = ""
    series: str | None = ""
    author_name: str | None = ""
    author_bio: str | None = ""
    target_audience: str | None = ""
    synopsis: str | None = ""
    isbn: str | None = ""
    publication_date: str | None = ""
    language: str | None = ""
    category: str | None = ""
    keywords: list[str] = Field(default_factory=list)
    page_count: str | None = ""
    trim_size: str | None = ""
    price: str | None = ""
    contributor: str | None = ""
    editor_notes: str | None = ""
    extras: dict[str, Any] = Field(default_factory=dict)


# ---------------------------------------------------------------------------
# Document helpers
# ---------------------------------------------------------------------------

def read_prompt_text() -> str:
    if not PROMPT_FILE.exists():
        raise FileNotFoundError(
            "找不到專案根目錄的 prompt.txt，請先建立後再執行擷取。"
        )
    return PROMPT_FILE.read_text(encoding="utf-8").strip()


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    lines: list[str] = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    lines.append(text)

    for section in doc.sections:
        for para in section.header.paragraphs + section.footer.paragraphs:
            text = para.text.strip()
            if text:
                lines.append(text)

    return "\n".join(lines)


def flatten_extracted_data(payload: dict[str, Any]) -> dict[str, str]:
    flat: dict[str, str] = {}

    def walk(obj: Any, prefix: str = "") -> None:
        if isinstance(obj, dict):
            for key, value in obj.items():
                walk(value, f"{prefix}.{key}" if prefix else key)
        elif isinstance(obj, list):
            flat[prefix] = (
                ", ".join(obj)
                if all(isinstance(i, str) for i in obj)
                else json.dumps(obj, ensure_ascii=False)
            )
        elif obj is None:
            flat[prefix] = ""
        else:
            flat[prefix] = str(obj)

    walk(payload)
    return flat


# ---------------------------------------------------------------------------
# Straico helpers
# ---------------------------------------------------------------------------

def parse_json_from_text(raw_text: str) -> dict[str, Any]:
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw_text, re.DOTALL)
    candidate = fenced.group(1) if fenced else raw_text

    # 1. Try strict parse first
    try:
        parsed = json.loads(candidate)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    # 2. LLM output often has unescaped quotes inside string values;
    #    json_repair fixes these automatically.
    try:
        repaired = repair_json(candidate, return_objects=True)
        if isinstance(repaired, dict) and repaired:
            return repaired
    except Exception:
        pass

    # 3. Last resort: brace-extract then repair
    brace = re.search(r"\{[\s\S]*\}", candidate)
    if brace:
        try:
            repaired = repair_json(brace.group(0), return_objects=True)
            if isinstance(repaired, dict) and repaired:
                return repaired
        except Exception:
            pass

    raise ValueError("無法從 Straico 回應中解析 JSON 物件。")


def normalize_extracted_payload(payload: dict[str, Any]) -> dict[str, Any]:
    known_keys = {
        "title", "subtitle", "series", "author_name", "author_bio",
        "target_audience", "synopsis", "isbn", "publication_date",
        "language", "category", "keywords", "page_count", "trim_size",
        "price", "contributor", "editor_notes", "extras",
    }
    extras: dict[str, Any] = payload.get("extras") or {}
    for key, value in payload.items():
        if key not in known_keys:
            extras[key] = value
    data = dict(payload)
    data["extras"] = extras if isinstance(extras, dict) else {}
    return BookFormData.model_validate(data).model_dump()


def _collect_text_candidates(obj: Any) -> list[str]:
    if isinstance(obj, str):
        return [obj]
    if isinstance(obj, dict):
        return [c for v in obj.values() for c in _collect_text_candidates(v)]
    if isinstance(obj, list):
        return [c for item in obj for c in _collect_text_candidates(item)]
    return []


def compose_straico_prompt(prompt_text: str, document_text: str) -> str:
    return (
        f"{prompt_text}\n\n"
        "INPUT DOCUMENT:\n"
        f"{document_text}\n\n"
        "Return only one valid JSON object."
    )


@retry(wait=wait_exponential(multiplier=1, min=1, max=8), stop=stop_after_attempt(3), reraise=True)
def call_straico(api_key: str, model: str, composed_prompt: str) -> tuple[dict[str, Any], str]:
    """Call Straico and return (parsed_json, raw_response_text)."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "x-api-key": api_key,
        "Content-Type": "application/json",
    }
    payload_variants = [
        {"model": model, "prompt": composed_prompt, "temperature": 0},
        {"models": [model], "prompt": composed_prompt, "temperature": 0},
        {"model": model, "message": composed_prompt, "temperature": 0},
    ]

    last_error: Exception | None = None
    for payload in payload_variants:
        resp = requests.post(
            STRAICO_ENDPOINT, headers=headers, json=payload, timeout=90)
        raw_text = resp.text
        if resp.status_code >= 400:
            last_error = RuntimeError(
                f"Straico 請求失敗，狀態碼 {resp.status_code}。\n{raw_text}"
            )
            continue

        data = resp.json()
        for candidate in _collect_text_candidates(data):
            try:
                return parse_json_from_text(candidate), raw_text
            except Exception:
                continue

        try:
            return parse_json_from_text(json.dumps(data, ensure_ascii=False)), raw_text
        except Exception as exc:
            last_error = exc

    raise last_error or RuntimeError("Straico 請求因不明原因失敗。")


# ---------------------------------------------------------------------------
# Form value helpers (config-driven)
# ---------------------------------------------------------------------------

def _parse_trim_dimension(trim_size: str, index: int) -> str:
    """Parse width (index=0) or height (index=1) from e.g. '150mm X 210mm'."""
    parts = re.split(r"\s*[xX×]\s*", trim_size.strip())
    if len(parts) > index:
        return re.sub(r"[^\d.]", "", parts[index])
    return ""


_DATE_FORMATS = [
    "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d",
    "%d-%m-%Y", "%d-%m-%y", "%B %d, %Y", "%b %d, %Y",
    "%d %B %Y", "%d %b %Y", "%Y年%m月%d日",
]


def _format_date_ddmmyy(value: str) -> str:
    """Try to parse *value* and return dd/mm/yy; return original on failure."""
    from datetime import datetime
    v = value.strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return v


def _parse_to_date(value: str):
    """Parse *value* with known formats and return a datetime.date, or None."""
    from datetime import datetime
    v = value.strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(v, fmt).date()
        except ValueError:
            continue
    return None


_CM_PLACEHOLDERS = {"width", "height", "thickness"}


def _mm_to_cm(value: str) -> str:
    """Convert a numeric mm string to cm (divide by 10), e.g. '150' → '15'."""
    try:
        cm = float(value) / 10
        return str(int(cm)) if cm == int(cm) else f"{cm:.1f}"
    except (ValueError, TypeError):
        return value


def build_initial_form_values(extracted_flat: dict[str, str]) -> dict[str, str]:
    """Map extracted flat JSON keys → template placeholder names via field_config."""
    trim_size = extracted_flat.get("trim_size", "")
    values: dict[str, str] = {}
    for section in SECTIONS:
        for field in section["fields"]:
            ph = field["placeholder"]
            if not ph:
                continue
            src = field.get("source_key", "")
            if src == "#trim_width":
                values[ph] = _parse_trim_dimension(trim_size, 0)
            elif src == "#trim_height":
                values[ph] = _parse_trim_dimension(trim_size, 1)
            elif src and src in extracted_flat:
                raw = extracted_flat[src]
                if ph == "publishDate":
                    raw = _format_date_ddmmyy(raw)
                values[ph] = raw
            else:
                values[ph] = ""
    return values


def _clear_form_widget_state() -> None:
    """Remove widget-key session entries so the form re-renders with fresh values."""
    for section in SECTIONS:
        for field in section["fields"]:
            if not field["placeholder"]:
                continue
            key = f"form_{field['placeholder']}"
            if key in st.session_state:
                del st.session_state[key]


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------

def render_docx_template(template_path: str, context: dict[str, str]) -> bytes:
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    buf = io.BytesIO()
    tpl.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

def main() -> None:
    enforce_workspace_auth()

    st.title("📝 新書表單產生器")
    st.caption("上傳 Word 文件，讓 AI 自動擷取書籍資料並填入範本。")

    with st.sidebar:
        st.subheader("帳號")
        st.caption(f"已登入：{getattr(st.user, 'email', '未知帳號')}")
        st.caption(f"模型：{STRAICO_MODEL}")
        if st.button("登出", key="sidebar_logout"):
            st.logout()
            st.stop()

    for key, default in [
        ("form_values", {}),
        ("last_output", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    if not TEMPLATE_FILE.exists():
        st.error("專案根目錄找不到 template.docx。")
        st.stop()

    try:
        prompt_text = read_prompt_text()
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    api_key = st.secrets.get("straico", {}).get("api_key", "")
    if not api_key:
        st.error("在 .streamlit/secrets.toml 中找不到 [straico].api_key")
        st.stop()

    uploaded_file = st.file_uploader("上傳來源 DOCX", type=["docx"])
    col_extract, col_reset = st.columns([2, 1])
    with col_extract:
        extract_clicked = st.button(
            "擷取資料並預填表單", type="primary", use_container_width=True
        )
    with col_reset:
        reset_clicked = st.button("重設", use_container_width=True)

    if reset_clicked:
        _clear_form_widget_state()
        st.session_state.form_values = {}
        st.session_state.last_output = None
        st.rerun()

    if extract_clicked:
        if uploaded_file is None:
            st.warning("請先上傳 DOCX 檔案。")
            st.stop()

        with st.spinner("正在讀取上傳的 DOCX..."):
            input_bytes = uploaded_file.getvalue()
            document_text = extract_text_from_docx(input_bytes)

        if not document_text.strip():
            st.error("無法從上傳 DOCX 擷取可讀文字。")
            st.stop()

        composed_prompt = compose_straico_prompt(prompt_text, document_text)
        start_time = time.monotonic()

        with st.spinner("正在從 Straico 擷取結構化 JSON..."):
            try:
                raw_json, raw_response = call_straico(
                    api_key, STRAICO_MODEL, composed_prompt
                )
                elapsed = time.monotonic() - start_time
                add_run(
                    file_name=uploaded_file.name,
                    model=STRAICO_MODEL,
                    raw_prompt=composed_prompt,
                    raw_response=raw_response,
                    elapsed_seconds=elapsed,
                    success=True,
                )
                normalized = normalize_extracted_payload(raw_json)
                extracted_flat = flatten_extracted_data(normalized)
                _clear_form_widget_state()
                st.session_state.form_values = build_initial_form_values(
                    extracted_flat)
                st.session_state.last_output = None
                st.success("擷取完成，您可在下方檢查並編輯欄位。")
            except ValidationError as exc:
                elapsed = time.monotonic() - start_time
                add_run(
                    file_name=uploaded_file.name,
                    model=STRAICO_MODEL,
                    raw_prompt=composed_prompt,
                    raw_response="",
                    elapsed_seconds=elapsed,
                    success=False,
                    error_message=str(exc),
                )
                st.error("Straico 回傳的資料格式不符合預期。")
                st.code(str(exc))
                st.stop()
            except Exception as exc:
                elapsed = time.monotonic() - start_time
                add_run(
                    file_name=getattr(uploaded_file, "name", "unknown"),
                    model=STRAICO_MODEL,
                    raw_prompt=composed_prompt,
                    raw_response="",
                    elapsed_seconds=elapsed,
                    success=False,
                    error_message=str(exc),
                )
                st.error(f"資料擷取失敗：{exc}")
                st.stop()

    if st.session_state.form_values:
        st.markdown("---")
        st.caption("所有欄位皆為選填，不需填寫的欄位可留空。打勾（✓）表示已選取。")

        # ── Render all sections in order (html fields inline at correct position) ──
        _CHECKBOX_COLS = 4
        fv = st.session_state.form_values  # shorthand

        for section in SECTIONS:
            fields = section["fields"]
            layout = section.get("layout", "grid")

            st.subheader(section["title"])

            if layout == "grid":
                # 2-column grid for non-checkbox fields; checkboxes (if any) after
                others = [f for f in fields if f["type"] != "checkbox"]
                checkboxes = [f for f in fields if f["type"] == "checkbox"]
                for i in range(0, len(others), 2):
                    pair = others[i:i + 2]
                    cols = st.columns(2)
                    for col, field in zip(cols, pair):
                        ph = field["placeholder"]
                        key = f"form_{ph}"
                        hint = field.get("hint") or None
                        opts = field.get("options")
                        with col:
                            if field["type"] == "html":
                                st.markdown(f"**{field['label']}**")
                                if key not in st.session_state:
                                    st_quill(value=fv.get(ph, ""),
                                             html=True, key=key)
                                else:
                                    st_quill(html=True, key=key)
                            elif field["type"] == "date":
                                if key not in st.session_state:
                                    st.date_input(field["label"], value=_parse_to_date(
                                        fv.get(ph, "")), key=key, format="DD/MM/YYYY")
                                else:
                                    st.date_input(
                                        field["label"], key=key, format="DD/MM/YYYY")
                            elif opts:
                                current = fv.get(ph, "")
                                idx = opts.index(
                                    current) if current in opts else 0
                                st.selectbox(
                                    field["label"], options=opts, index=idx, key=key)
                            elif field["type"] == "textarea":
                                if key not in st.session_state:
                                    st.text_area(field["label"], value=fv.get(
                                        ph, ""), key=key, placeholder=hint)
                                else:
                                    st.text_area(
                                        field["label"], key=key, placeholder=hint)
                            else:
                                if key not in st.session_state:
                                    st.text_input(field["label"], value=fv.get(
                                        ph, ""), key=key, placeholder=hint)
                                else:
                                    st.text_input(
                                        field["label"], key=key, placeholder=hint)
                for i in range(0, len(checkboxes), _CHECKBOX_COLS):
                    group = checkboxes[i:i + _CHECKBOX_COLS]
                    cols = st.columns(_CHECKBOX_COLS)
                    for j, field in enumerate(group):
                        ph = field["placeholder"]
                        key = f"form_{ph}"
                        with cols[j]:
                            if key not in st.session_state:
                                st.checkbox(field["label"], value=fv.get(
                                    ph, UNCHECKED_CHAR) == CHECKBOX_CHAR, key=key)
                            else:
                                st.checkbox(field["label"], key=key)
            else:
                # Render in config order: group consecutive checkboxes 4/row,
                # non-checkbox fields rendered individually at their config position
                idx = 0
                while idx < len(fields):
                    field = fields[idx]
                    if field["type"] == "checkbox":
                        # Collect run of consecutive checkboxes
                        run: list[dict] = []
                        while idx < len(fields) and fields[idx]["type"] == "checkbox":
                            run.append(fields[idx])
                            idx += 1
                        for j in range(0, len(run), _CHECKBOX_COLS):
                            group = run[j:j + _CHECKBOX_COLS]
                            cols = st.columns(_CHECKBOX_COLS)
                            for k, cb_field in enumerate(group):
                                ph = cb_field["placeholder"]
                                key = f"form_{ph}"
                                with cols[k]:
                                    if key not in st.session_state:
                                        st.checkbox(cb_field["label"], value=fv.get(
                                            ph, UNCHECKED_CHAR) == CHECKBOX_CHAR, key=key)
                                    else:
                                        st.checkbox(cb_field["label"], key=key)
                    else:
                        if field["type"] == "label":
                            st.markdown(f"**{field['label']}**")
                            idx += 1
                        else:
                            ph = field["placeholder"]
                            key = f"form_{ph}"
                            hint = field.get("hint") or None
                            opts = field.get("options")
                            if field["type"] == "html":
                                st.markdown(f"**{field['label']}**")
                                if key not in st.session_state:
                                    st_quill(value=fv.get(ph, ""),
                                             html=True, key=key)
                                else:
                                    st_quill(html=True, key=key)
                            elif field["type"] == "date":
                                if key not in st.session_state:
                                    st.date_input(field["label"], value=_parse_to_date(
                                        fv.get(ph, "")), key=key, format="DD/MM/YYYY")
                                else:
                                    st.date_input(
                                        field["label"], key=key, format="DD/MM/YYYY")
                            elif opts:
                                current = fv.get(ph, "")
                                sidx = opts.index(
                                    current) if current in opts else 0
                                st.selectbox(
                                    field["label"], options=opts, index=sidx, key=key)
                            elif field["type"] == "textarea":
                                if key not in st.session_state:
                                    st.text_area(field["label"], value=fv.get(
                                        ph, ""), key=key, placeholder=hint)
                                else:
                                    st.text_area(
                                        field["label"], key=key, placeholder=hint)
                            else:
                                if key not in st.session_state:
                                    st.text_input(field["label"], value=fv.get(
                                        ph, ""), key=key, placeholder=hint)
                                else:
                                    st.text_input(
                                        field["label"], key=key, placeholder=hint)
                            idx += 1

            st.markdown("---")

        generate_clicked = st.button(
            "產生最終文件", type="primary", use_container_width=True
        )

        if generate_clicked:
            # Collect all current values from session state
            all_updated: dict[str, str] = {}
            for section in SECTIONS:
                for field in section["fields"]:
                    ph = field["placeholder"]
                    if not ph:  # skip pseudo-fields (e.g. label markers)
                        continue
                    key = f"form_{ph}"
                    if field["type"] == "checkbox":
                        val = st.session_state.get(key, False)
                        all_updated[ph] = CHECKBOX_CHAR if val else UNCHECKED_CHAR
                    elif field["type"] == "date":
                        _dval = st.session_state.get(key)
                        if hasattr(_dval, "strftime"):
                            all_updated[ph] = _dval.strftime("%d/%m/%Y")
                        else:
                            all_updated[ph] = _format_date_ddmmyy(
                                str(_dval)) if _dval else ""
                    else:
                        all_updated[ph] = st.session_state.get(key, "")
            st.session_state.form_values = all_updated
            with st.spinner("正在將資料合併到 template.docx..."):
                docx_context = dict(all_updated)
                for ph in _CM_PLACEHOLDERS:
                    if docx_context.get(ph):
                        docx_context[ph] = _mm_to_cm(docx_context[ph])
                st.session_state.last_output = render_docx_template(
                    str(TEMPLATE_FILE), docx_context
                )
            st.success("文件已產生，可於下方下載。")

    if st.session_state.last_output:
        st.download_button(
            label="⬇️ 下載最終 DOCX",
            data=st.session_state.last_output,
            file_name="merged_output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )


if __name__ == "__main__":
    main()
