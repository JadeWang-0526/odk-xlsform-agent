"""Streamlit chat interface for the ODK XLSForm ADK agent."""

from __future__ import annotations

import asyncio
import os
import re
from pathlib import Path
from typing import Any

import streamlit as st

# Set API key from Streamlit secrets (for cloud) or fall back to env var
if "GOOGLE_API_KEY" in st.secrets:
    os.environ["GOOGLE_API_KEY"] = st.secrets["GOOGLE_API_KEY"]

from google.adk.runners import Runner
from google.adk.sessions import InMemorySessionService
from google.genai import types

from agent import root_agent

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="ODK XLSForm Agent", page_icon="📋", layout="wide")
st.title("ODK XLSForm Agent")
st.caption("Design and generate ODK XLSForm through chat.")

# ---------------------------------------------------------------------------
# ADK runner (cached so it persists across reruns)
# ---------------------------------------------------------------------------
@st.cache_resource
def get_runner_and_session_service():
    session_service = InMemorySessionService()
    runner = Runner(
        agent=root_agent,
        app_name="odk_xlsform_agent",
        session_service=session_service,
    )
    return runner, session_service


runner, session_service = get_runner_and_session_service()

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
if "messages" not in st.session_state:
    st.session_state.messages = []
if "session_id" not in st.session_state:
    st.session_state.session_id = "streamlit-session-0"
if "user_id" not in st.session_state:
    st.session_state.user_id = "streamlit-user"
if "session_counter" not in st.session_state:
    st.session_state.session_counter = 0
if "xlsx_files" not in st.session_state:
    st.session_state.xlsx_files = []
if "button_clicked" not in st.session_state:
    st.session_state.button_clicked = None

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Controls")
    if st.button("New Chat", use_container_width=True):
        st.session_state.session_counter += 1
        st.session_state.session_id = f"streamlit-session-{st.session_state.session_counter}"
        st.session_state.messages = []
        st.session_state.xlsx_files = []
        st.rerun()

    # Show download buttons for any generated xlsx files
    if st.session_state.xlsx_files:
        st.divider()
        st.subheader("Generated Files")
        for filepath in st.session_state.xlsx_files:
            p = Path(filepath)
            if p.exists():
                st.download_button(
                    label=f"Download {p.name}",
                    data=p.read_bytes(),
                    file_name=p.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _extract_xlsx_paths(text: str) -> list[str]:
    """Find .xlsx file paths mentioned in text."""
    return re.findall(r"[\w/\\._-]+\.xlsx", text)


def _extract_xlsx_paths_from_obj(obj: Any) -> list[str]:
    """Best-effort extractor that looks for .xlsx paths in arbitrary objects."""
    paths: list[str] = []
    seen_ids: set[int] = set()

    def _walk(value: Any) -> None:
        if value is None:
            return
        if id(value) in seen_ids:
            return
        seen_ids.add(id(value))
        if isinstance(value, (str, os.PathLike)):
            paths.extend(_extract_xlsx_paths(str(value)))
            return
        if isinstance(value, dict):
            for v in value.values():
                _walk(v)
            return
        if isinstance(value, (list, tuple, set)):
            for v in value:
                _walk(v)
            return
        try:
            paths.extend(_extract_xlsx_paths(str(value)))
        except Exception:
            pass

    _walk(obj)
    return paths


def _remember_xlsx_path(path: str) -> None:
    """Normalize and store a discovered XLSX path if it exists."""
    p = Path(path).expanduser().resolve()
    if p.exists():
        normalized = str(p)
        if normalized not in st.session_state.xlsx_files:
            st.session_state.xlsx_files.append(normalized)


def _summarize_result(tool_name: str, result: Any) -> str:
    """Build a short human-readable summary of a tool result."""
    if not isinstance(result, dict):
        return str(result)[:300]
    lines: list[str] = []
    if tool_name in ("write_xlsform", "save_xlsform_draft"):
        if "output_path" in result:
            lines.append(f"Saved to: `{result['output_path']}`")
        if "row_counts" in result:
            lines.append(f"Row counts: {result['row_counts']}")
        if "sheet_names" in result:
            lines.append(f"Sheets: {result['sheet_names']}")
        v = result.get("validation") or {}
        if v.get("valid") is True:
            w = v.get("warnings") or []
            lines.append(f"✅ Validation passed" + (f" ({len(w)} warning{'s' if len(w) != 1 else ''})" if w else ""))
        elif v.get("valid") is False:
            errs = v.get("errors") or []
            lines.append(f"❌ Validation FAILED — {len(errs)} error{'s' if len(errs) != 1 else ''}")
            for e in errs[:3]:
                lines.append(f"  • {e[:120]}")
    elif tool_name == "validate_xlsform":
        if result.get("valid") is True:
            w = result.get("warnings") or []
            lines.append(f"✅ Valid" + (f" — {len(w)} warning{'s' if len(w) != 1 else ''}" if w else ""))
            for w_msg in w[:3]:
                lines.append(f"  ⚠ {w_msg[:120]}")
        elif result.get("valid") is False:
            errs = result.get("errors") or []
            lines.append(f"❌ Invalid — {len(errs)} error{'s' if len(errs) != 1 else ''}")
            for e in errs[:5]:
                lines.append(f"  • {e[:120]}")
    elif tool_name == "load_xlsform":
        if "path" in result:
            lines.append(f"Loaded: `{result['path']}`")
        if "sheet_names" in result:
            lines.append(f"Sheets: {result['sheet_names']}")
        if "row_counts" in result:
            lines.append(f"Row counts: {result['row_counts']}")
    elif tool_name == "new_form_spec":
        sheets = result.get("sheets", {})
        lines.append(f"Created blank spec — sheets: {list(sheets.keys())}")
        for sn, sv in sheets.items():
            rows = sv.get("rows", [])
            lines.append(f"  {sn}: {len(rows)} starter rows")
    elif tool_name == "design_survey_outline":
        lines.append(f"Topic: {result.get('topic', '?')}")
        lines.append(f"Survey rows generated: {len(result.get('survey_rows', []))}")
        lines.append(f"Choice rows generated: {len(result.get('choices_rows', []))}")
        if result.get("languages"):
            lines.append(f"Languages: {result['languages']}")
    elif tool_name == "merge_form_spec":
        summary = result.get("summary", {})
        added = summary.get("added", {})
        skipped = summary.get("skipped", {})
        for sheet, names in added.items():
            lines.append(f"Added {len(names)} rows to '{sheet}'")
        for sheet, names in skipped.items():
            if names:
                lines.append(f"Skipped {len(names)} duplicates in '{sheet}'")
    elif tool_name == "add_calculations_and_conditions":
        lines.append(f"Added calculations: {result.get('added_calculations', [])}")
        lines.append(f"Updated targets: {result.get('updated_targets', [])}")
    else:
        # Generic: show top-level scalar/count fields
        for k, v in list(result.items())[:8]:
            if isinstance(v, list):
                lines.append(f"{k}: [{len(v)} items]")
            elif isinstance(v, dict):
                lines.append(f"{k}: {{{len(v)} keys}}")
            else:
                lines.append(f"{k}: {str(v)[:120]}")
    return "\n".join(lines) if lines else "(no summary available)"


def _args_summary(tool_name: str, args: dict) -> str:
    """Concise one-line summary of tool call arguments."""
    if not args:
        return "(no args)"
    # Show scalar args; abbreviate large structures
    parts: list[str] = []
    for k, v in args.items():
        if isinstance(v, (str, int, float, bool)):
            parts.append(f"{k}={repr(v)}")
        elif isinstance(v, list):
            parts.append(f"{k}=[{len(v)} items]")
        elif isinstance(v, dict):
            parts.append(f"{k}={{...}}")
        else:
            parts.append(f"{k}=...")
    return ", ".join(parts)


def _render_tool_steps(tool_steps: list[dict]) -> None:
    """Render tool call details inside an expander."""
    if not tool_steps:
        return
    label = f"🔧 {len(tool_steps)} tool call{'s' if len(tool_steps) != 1 else ''} processed"
    with st.expander(label, expanded=False):
        for i, step in enumerate(tool_steps):
            if i > 0:
                st.divider()
            name = step["name"]
            args = step.get("args") or {}
            result = step.get("result")

            st.markdown(f"**`{name}`** — {_args_summary(name, args)}")

            col_args, col_result = st.columns(2)
            with col_args:
                if args:
                    st.caption("Arguments")
                    # Show scalar args as text, large structures as JSON
                    simple = {k: v for k, v in args.items()
                              if isinstance(v, (str, int, float, bool, type(None)))}
                    complex_ = {k: v for k, v in args.items() if k not in simple}
                    if simple:
                        for k, v in simple.items():
                            st.text(f"{k}: {v}")
                    if complex_:
                        st.json(complex_, expanded=False)

            with col_result:
                if result is not None:
                    st.caption("Result summary")
                    st.text(_summarize_result(name, result))


# (pattern, yes_label, no_label)
# Patterns are checked in order — put more specific ones first.
_BINARY_QUESTION_PATTERNS: list[tuple[str, str, str]] = [

    # --- Translations ---
    (r"\bdo these translations look correct\b",
     "Yes, translations are correct — proceed to save",
     "No, I want to make further changes"),
    (r"\bdo the translations look (correct|right|good)\b",
     "Yes, proceed to save",
     "No, make changes first"),

    # --- Saving ---
    (r"\bshall i proceed to save\b",   "Yes, save the form now",  "No, make changes first"),
    (r"\bproceed to save\b",           "Yes, save the form now",  "No, make changes first"),
    (r"\bshall i save\b",              "Yes, save it",            "No, not yet"),
    (r"\bdo you want me to save\b",    "Yes, save it",            "No, not yet"),
    (r"\bwould you like (me )?to save\b", "Yes, save it",         "No, not yet"),

    # --- Conditions and calculations ---
    (r"\bdo you agree with these (conditions|calculations|suggestions|items)\b",
     "Yes, I agree with all of them",
     "No, I want to adjust some"),
    (r"\bwould you like to (add|remove|change) any\b",
     "Yes, everything looks good",
     "No, I want to adjust some"),
    (r"\badd, remove, or change any\b",
     "Yes, everything looks good",
     "No, I want to make adjustments"),

    # --- General approval of lists / outlines / tables ---
    (r"\bdoes this (outline|question list|list of questions|structure|table|form spec|summary) look (right|correct|good|ok)\b",
     "Yes, looks good",
     "No, make changes"),
    (r"\bdo(es)? (this|the|these) (questions?|rows?|sections?|fields?|choices?) look (right|correct|good|ok)\b",
     "Yes, looks good",
     "No, make changes"),
    (r"\bwhat should we (add|remove|change)\b",
     "Looks good, nothing to change",
     "I have some changes"),
    (r"\bdoes (everything|this) look (right|correct|good|ok)\b",
     "Yes, looks good",
     "No, make changes"),

    # --- Language list confirmation (STEP 3) ---
    (r"\b(is the|does the) language (list|selection|order) (look )?(correct|right|ok)\b",
     "Yes, the language list is correct",
     "No, I want to change it"),

    # --- Auto-generate vs manual (STEP 5) ---
    (r"\bauto.?generat\b",
     "Auto-generate a starter outline",
     "I'll describe the questions manually"),

    # --- Generic proceed / continue ---
    (r"\bshall i proceed\b",           "Yes, proceed",   "No, wait"),
    (r"\bshall i (go ahead|continue|move on)\b", "Yes, continue", "No, pause here"),
    (r"\bshall i (add|apply|include|attach|merge)\b", "Yes, go ahead", "No, not yet"),
    (r"\bwould you like (me )?to (add|apply|include|attach|merge|continue|proceed)\b",
     "Yes, please",
     "No, not yet"),

    # --- Generic agreement / correctness ---
    (r"\bdo you agree\b",              "Yes, I agree",   "No, let me adjust"),
    (r"\bis this (correct|right|good|ok)\b", "Yes, correct", "No, make changes"),
    (r"\bare (these|the) (settings|details|fields|values) (correct|right|ok)\b",
     "Yes, all correct",
     "No, I want to change some"),
]


def _detect_binary_question(text: str) -> dict | None:
    """Return {"yes": label, "no": label} if the text contains a binary decision question."""
    lower = text.lower()
    for pattern, yes_label, no_label in _BINARY_QUESTION_PATTERNS:
        if re.search(pattern, lower):
            return {"yes": yes_label, "no": no_label}
    return None


def _render_assistant_message(msg: dict) -> None:
    """Render one assistant message (text + optional tool steps)."""
    content = msg.get("content", "")
    if content:
        st.markdown(content)
    _render_tool_steps(msg.get("tool_steps") or [])


# ---------------------------------------------------------------------------
# Display chat history
# ---------------------------------------------------------------------------
for i, msg in enumerate(st.session_state.messages):
    is_last = i == len(st.session_state.messages) - 1
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            _render_assistant_message(msg)
            # Show yes/no buttons only for the last assistant message,
            # and only when no button click is already being processed.
            if is_last and not st.session_state.button_clicked:
                buttons = _detect_binary_question(msg.get("content", ""))
                if buttons:
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(buttons["yes"], key="btn_yes", type="primary", use_container_width=True):
                            st.session_state.button_clicked = buttons["yes"]
                            st.rerun()
                    with col2:
                        if st.button(buttons["no"], key="btn_no", use_container_width=True):
                            st.session_state.button_clicked = buttons["no"]
                            st.rerun()
        else:
            st.markdown(msg["content"])

# ---------------------------------------------------------------------------
# Agent runner
# ---------------------------------------------------------------------------
async def _ensure_session() -> None:
    """Create the ADK session if it doesn't exist yet."""
    existing = await session_service.get_session(
        app_name="odk_xlsform_agent",
        user_id=st.session_state.user_id,
        session_id=st.session_state.session_id,
    )
    if existing is None:
        await session_service.create_session(
            app_name="odk_xlsform_agent",
            user_id=st.session_state.user_id,
            session_id=st.session_state.session_id,
        )


async def _run_agent(user_input: str) -> tuple[str, list[dict]]:
    """
    Run the agent and return (response_text, tool_steps).

    tool_steps is a list of dicts:
        {"name": str, "args": dict, "result": dict | None}
    """
    await _ensure_session()
    content = types.Content(
        role="user",
        parts=[types.Part(text=user_input)],
    )

    response_text = ""
    tool_steps: list[dict] = []

    async for event in runner.run_async(
        user_id=st.session_state.user_id,
        session_id=st.session_state.session_id,
        new_message=content,
    ):
        if not event.content or not event.content.parts:
            continue

        for part in event.content.parts:
            # --- tool call ---
            fc = getattr(part, "function_call", None)
            if fc is not None:
                tool_steps.append({
                    "name": fc.name,
                    "args": dict(fc.args) if fc.args else {},
                    "result": None,
                })

            # --- tool response ---
            fr = getattr(part, "function_response", None)
            if fr is not None:
                raw = getattr(fr, "response", None) or {}
                result = dict(raw) if isinstance(raw, dict) else raw
                # Scan for generated xlsx files
                for path in _extract_xlsx_paths_from_obj(result):
                    _remember_xlsx_path(path)
                # Match to the most recent unmatched call with the same name
                for step in reversed(tool_steps):
                    if step["name"] == fr.name and step["result"] is None:
                        step["result"] = result
                        break

            # --- text response (collect from ALL events, not just final) ---
            # This ensures we never miss model text that arrives in non-final
            # streaming events before the "done" signal.
            text = getattr(part, "text", None)
            if text:
                response_text += text

    # Scan response text for xlsx paths too
    for path in _extract_xlsx_paths(response_text):
        _remember_xlsx_path(path)

    # Fallback 1: synthesise a summary from tool results when model emitted no text
    if not response_text.strip() and tool_steps:
        lines: list[str] = []
        for step in tool_steps:
            summary = _summarize_result(step["name"], step.get("result") or {})
            lines.append(f"**`{step['name']}`** completed:\n{summary}")
        response_text = "\n\n".join(lines)

    # Fallback 2: mention the most-recently saved file
    if not response_text.strip() and st.session_state.xlsx_files:
        latest = st.session_state.xlsx_files[-1]
        response_text = f"Saved draft: `{Path(latest).name}`"

    # Fallback 3: absolute last resort — nudge the model once, then use a
    # placeholder.  This prevents an empty assistant message from being stored
    # in session state, which would cause Gemini to return empty on every
    # subsequent turn (the "stuck" loop).
    if not response_text.strip():
        nudge = types.Content(
            role="user",
            parts=[types.Part(text="Please continue and respond to my previous message.")],
        )
        try:
            async for event in runner.run_async(
                user_id=st.session_state.user_id,
                session_id=st.session_state.session_id,
                new_message=nudge,
            ):
                if not event.content or not event.content.parts:
                    continue
                for part in event.content.parts:
                    text = getattr(part, "text", None)
                    if text:
                        response_text += text
        except Exception:
            pass

    # Absolute placeholder if everything above failed
    if not response_text.strip():
        response_text = "_(Understood. What would you like to do next?)_"

    return response_text, tool_steps


# ---------------------------------------------------------------------------
# Chat input  (button clicks take priority over typed input)
# ---------------------------------------------------------------------------
user_input: str | None = None

# Consume a pending button click
if st.session_state.button_clicked:
    user_input = st.session_state.button_clicked
    st.session_state.button_clicked = None

# Fall back to the text input box
if not user_input:
    user_input = st.chat_input("Describe the survey you'd like to create…")

if user_input:
    # Show user message
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    # Run agent and render response
    with st.chat_message("assistant"):
        with st.spinner("Thinking…"):
            response_text, tool_steps = asyncio.run(_run_agent(user_input))

        msg = {"role": "assistant", "content": response_text, "tool_steps": tool_steps}
        _render_assistant_message(msg)

    # Only store the message when it has visible content — an empty assistant
    # message in session state causes Gemini to return empty on the next turn.
    if response_text.strip() or tool_steps:
        st.session_state.messages.append(msg)

    # Rerun to show updated sidebar download buttons or fresh buttons
    st.rerun()
