"""Streamlit chat interface for the ODK XLSForm ADK agent."""

from __future__ import annotations

import asyncio
import os
import re
from pathlib import Path

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
st.caption("Design and generate ODK XLSForm surveys through chat.")

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
# Display chat history
# ---------------------------------------------------------------------------
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _extract_xlsx_paths(text: str) -> list[str]:
    """Find .xlsx file paths mentioned in text."""
    return re.findall(r"[\w/\\._-]+\.xlsx", text)


def _extract_xlsx_paths_from_obj(obj) -> list[str]:
    """Best-effort extractor that looks for .xlsx paths in arbitrary tool results."""
    paths: list[str] = []
    seen_ids: set[int] = set()

    def _walk(value) -> None:
        if value is None:
            return
        if id(value) in seen_ids:
            return
        seen_ids.add(id(value))

        # Strings / Paths
        if isinstance(value, (str, os.PathLike)):
            paths.extend(_extract_xlsx_paths(str(value)))
            return

        # Collections
        if isinstance(value, dict):
            for v in value.values():
                _walk(v)
            return
        if isinstance(value, (list, tuple, set)):
            for v in value:
                _walk(v)
            return

        # Attributes commonly used in tool results
        for attr in (
            "output_path",
            "path",
            "file_path",
            "filepath",
            "result",
            "text",
            "content",
        ):
            if hasattr(value, attr):
                try:
                    _walk(getattr(value, attr))
                except Exception:
                    pass

        # Fallback to string representation
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


async def _ensure_session():
    """Create the ADK session if it doesn't exist yet."""
    session_id = st.session_state.session_id
    user_id = st.session_state.user_id
    existing = await session_service.get_session(
        app_name="odk_xlsform_agent",
        user_id=user_id,
        session_id=session_id,
    )
    if existing is None:
        await session_service.create_session(
            app_name="odk_xlsform_agent",
            user_id=user_id,
            session_id=session_id,
        )


async def _run_agent(user_input: str) -> str:
    """Send a message to the ADK agent and return the final response text."""
    await _ensure_session()
    content = types.Content(
        role="user",
        parts=[types.Part(text=user_input)],
    )
    response_text = ""
    tool_messages: list[str] = []
    async for event in runner.run_async(
        user_id=st.session_state.user_id,
        session_id=st.session_state.session_id,
        new_message=content,
    ):
        # Collect final response
        if event.is_final_response() and event.content and event.content.parts:
            for part in event.content.parts:
                if part.text:
                    response_text += part.text

        # Detect xlsx files from tool responses
        if hasattr(event, "actions") and event.actions:
            tool_results = event.actions.tool_results if hasattr(event.actions, "tool_results") else []
            for result in tool_results:
                for path in _extract_xlsx_paths_from_obj(result):
                    _remember_xlsx_path(path)
                # Capture any text/result content so the user sees a meaningful reply
                msg = None
                if hasattr(result, "text") and result.text:
                    msg = str(result.text)
                elif hasattr(result, "result") and result.result not in (None, ""):
                    msg = str(result.result)
                elif hasattr(result, "output") and result.output not in (None, ""):
                    msg = str(result.output)
                else:
                    # Fallback to repr if nothing else is available
                    try:
                        msg = str(result)
                    except Exception:
                        msg = None
                if msg and msg.strip():
                    tool_messages.append(msg.strip())

    # Also scan the response text for xlsx paths
    for path in _extract_xlsx_paths(response_text):
        _remember_xlsx_path(path)

    # If the agent produced no text, surface tool output or saved file info to avoid empty replies
    if not response_text.strip():
        if tool_messages:
            response_text = "\n".join(tool_messages)
        elif st.session_state.xlsx_files:
            latest = st.session_state.xlsx_files[-1]
            response_text = f"Saved draft: {Path(latest).name}"

    return response_text

# ---------------------------------------------------------------------------
# Chat input
# ---------------------------------------------------------------------------
if user_input := st.chat_input("Describe the survey you'd like to create…"):
    # Show user message
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    # Get agent response
    with st.chat_message("assistant"):
        with st.spinner("Thinking…"):
            response = asyncio.run(_run_agent(user_input))
        st.markdown(response)

    st.session_state.messages.append({"role": "assistant", "content": response})

    # Rerun to refresh sidebar download buttons
    if st.session_state.xlsx_files:
        st.rerun()
