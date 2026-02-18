

from __future__ import annotations

import copy
import datetime
import re
import tempfile
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from google.adk.agents import Agent
from openpyxl import Workbook, load_workbook

# Preferred column orderings to keep XLSForm sheets tidy.
_PREFERRED_COLUMN_ORDER: Dict[str, List[str]] = {
    "survey": [
        "type",
        "name",
        "label",
        "hint",
        "required",
        "relevant",
        "appearance",
        "choice_filter",
        "constraint",
        "constraint_message",
        "calculation",
        "repeat_count",
        "default",
        "autoplay",
    ],
    "choices": [
        "list_name",
        "name",
        "label",
        "choice_filter",
    ],
    "settings": [
        "form_title",
        "form_id",
        "default_language",
        "version",
        "public_key",
        "submission_url",
        "instance_name",
    ],
}

_TEMPLATE_URL = (
    "https://github.com/getodk/xlsform-template/raw/main/ODK%20XLSForm%20Template.xlsx"
)
_TEMPLATE_CACHE_PATH = Path(tempfile.gettempdir()) / "odk_xlsform_template.xlsx"

# Common language names mapped to their BCP 47 / IANA subtags.
_CANONICAL_LANGUAGE_CODES: Dict[str, str] = {
    "Arabic": "ar",
    "Bengali": "bn",
    "Bulgarian": "bg",
    "Chinese": "zh",
    "Croatian": "hr",
    "Czech": "cs",
    "Danish": "da",
    "Dutch": "nl",
    "English": "en",
    "Finnish": "fi",
    "French": "fr",
    "German": "de",
    "Greek": "el",
    "Hebrew": "he",
    "Hindi": "hi",
    "Hungarian": "hu",
    "Indonesian": "id",
    "Italian": "it",
    "Japanese": "ja",
    "Korean": "ko",
    "Malay": "ms",
    "Norwegian": "no",
    "Persian": "fa",
    "Polish": "pl",
    "Portuguese": "pt",
    "Romanian": "ro",
    "Russian": "ru",
    "Serbian": "sr",
    "Somali": "so",
    "Spanish": "es",
    "Swahili": "sw",
    "Swedish": "sv",
    "Thai": "th",
    "Turkish": "tr",
    "Ukrainian": "uk",
    "Urdu": "ur",
    "Vietnamese": "vi",
}
_LANGUAGE_CODE_MAP: Dict[str, str] = {name.lower(): code for name, code in _CANONICAL_LANGUAGE_CODES.items()}
_LANGUAGE_CODE_MAP.update(
    {
        "farsi": "fa",
        "filipino": "fil",
        "kiswahili": "sw",
        "mandarin": "zh",
        "pashto": "ps",
        "swahili": "sw",
        "tagalog": "tl",
    }
)
_CODE_TO_NAME: Dict[str, str] = {}
for name, code in _CANONICAL_LANGUAGE_CODES.items():
    _CODE_TO_NAME.setdefault(code, name)


def _get_template_path() -> Optional[Path]:
    """Download the official ODK XLSForm template, caching it in the system temp dir."""
    if _TEMPLATE_CACHE_PATH.exists():
        return _TEMPLATE_CACHE_PATH
    try:
        urllib.request.urlretrieve(_TEMPLATE_URL, _TEMPLATE_CACHE_PATH)
        return _TEMPLATE_CACHE_PATH
    except Exception:
        return None


def _safe_form_id(title: str) -> str:
    """Convert an arbitrary title into a safe XLSForm id."""
    cleaned = "".join(ch.lower() if ch.isalnum() else "_" for ch in title).strip("_")
    cleaned = "_".join(part for part in cleaned.split("_") if part)
    return cleaned or "odk_form"


def _normalize_language_tag(language: str) -> Dict[str, str]:
    """Return name/code/header fields for an input language string."""
    lang = (language or "").strip()
    if not lang:
        return {"header": "", "code": "", "name": ""}

    paren_match = re.match(r"^(.*?)(?:\s*\(([^)]+)\))\s*$", lang)
    if paren_match:
        name, code = paren_match.groups()
        name = name.strip()
        code = code.strip()
        normalized_code = code.lower()
        name = name or _CODE_TO_NAME.get(normalized_code, normalized_code)
        return {"header": f"{name} ({normalized_code})", "code": normalized_code, "name": name}

    key = lang.lower()
    if key in _LANGUAGE_CODE_MAP:
        code = _LANGUAGE_CODE_MAP[key]
        name = _CODE_TO_NAME.get(code, lang.strip() or code)
        return {"header": f"{name} ({code})", "code": code, "name": name}

    if re.fullmatch(r"[a-zA-Z]{2,3}(?:-[a-zA-Z0-9]{2,8})*", lang):
        code = lang.lower()
        name = _CODE_TO_NAME.get(code, lang)
        return {"header": f"{name} ({code})", "code": code, "name": name}

    safe_code = "".join(ch.lower() for ch in lang if ch.isalnum() or ch == "-") or "und"
    name = lang
    return {"header": f"{name} ({safe_code})", "code": safe_code, "name": name}


def _normalize_languages(languages: Sequence[str]) -> List[Dict[str, str]]:
    """Normalize a list of language names to include IANA codes."""
    normalized: List[Dict[str, str]] = []
    seen_headers: set[str] = set()
    for lang in languages or []:
        entry = _normalize_language_tag(lang)
        header = entry.get("header")
        if header and header not in seen_headers:
            seen_headers.add(header)
            normalized.append(entry)
    return normalized


def _language_headers_from_columns(columns: Sequence[str]) -> List[str]:
    """Extract unique language headers from language-bearing columns."""
    headers: List[str] = []
    for col in columns or []:
        m = re.match(r"^(label|hint|constraint_message)::(.+)$", str(col) or "")
        if m:
            header = m.group(2).strip()
            if header and header not in headers:
                headers.append(header)
    return headers


def _normalize_language_column_name(col: str) -> str:
    """Normalize language-bearing column names to include the BCP47 code."""
    m = re.match(r"^(label|hint|constraint_message)::\s*(.+)$", col or "")
    if not m:
        return col
    prefix, lang_part = m.groups()
    normalized = _normalize_language_tag(lang_part)
    header = normalized.get("header")
    return f"{prefix}::{header}" if header else col


def _normalize_language_columns_and_rows(
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> tuple[List[str], List[Dict[str, Any]]]:
    """
    Ensure any label/hint/constraint_message language columns include the code,
    updating both the column list and row keys in place-safe copies.
    """
    all_keys = set(columns)
    for row in rows:
        all_keys.update(row.keys())

    mapping: Dict[str, str] = {}
    for key in all_keys:
        new_key = _normalize_language_column_name(key)
        mapping[key] = new_key or key

    updated_rows: List[Dict[str, Any]] = []
    for row in rows:
        new_row: Dict[str, Any] = {}
        for key, val in row.items():
            new_key = mapping.get(key, key)
            if new_key in new_row:
                # Prefer the non-empty value when a collision occurs.
                if new_row[new_key] in (None, "", " ") and val not in (None, "", " "):
                    new_row[new_key] = val
            else:
                new_row[new_key] = val
        updated_rows.append(new_row)

    updated_columns: List[str] = []
    for col in columns:
        new_col = mapping.get(col, col)
        if new_col not in updated_columns:
            updated_columns.append(new_col)

    # Include any normalized keys present in rows but missing from columns.
    for key in all_keys:
        new_key = mapping.get(key, key)
        if new_key not in updated_columns:
            updated_columns.append(new_key)

    lang_headers = _language_headers_from_columns(updated_columns)
    if lang_headers:
        user_facing_fields = ["label", "hint", "constraint_message", "required_message", "guidance_hint"]

        for field in user_facing_fields:
            lang_cols = [f"{field}::{lang}" for lang in lang_headers]

            # Ensure language-specific columns exist.
            for col in lang_cols:
                if col not in updated_columns:
                    updated_columns.append(col)

            for row in updated_rows:
                # Pick a fallback value from any existing language value or the base field.
                fallback = None
                for candidate_key in lang_cols + [field]:
                    val = row.get(candidate_key)
                    if val not in (None, "", " "):
                        fallback = val
                        break

                for col in lang_cols:
                    if row.get(col) in (None, "", " ") and fallback not in (None, "", " "):
                        row[col] = fallback

                # Drop the non-language column to avoid a "default" language in ODK.
                if field in row:
                    row.pop(field, None)

            # Remove the non-language column from the column list as well.
            updated_columns = [c for c in updated_columns if c != field]

    return updated_columns, updated_rows


def _row_has_content(values: Sequence[Any]) -> bool:
    return any(val not in (None, "", " ") for val in values)


def _infer_columns(rows: List[Dict[str, Any]], preferred: Optional[List[str]]) -> List[str]:
    seen: List[str] = []
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.append(key)
    if preferred:
        ordered: List[str] = [col for col in preferred if col in seen]
        ordered.extend(col for col in seen if col not in ordered)
        return ordered
    return seen


def _normalize_form_spec(form_spec: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """
    Accepts either {"sheets": {...}} or a direct mapping of sheet_name -> sheet_spec.
    Each sheet_spec should contain optional "columns" and a "rows" list of dicts.
    """
    if "sheets" in form_spec and isinstance(form_spec["sheets"], dict):
        sheets = form_spec["sheets"]
    else:
        sheets = {k: v for k, v in form_spec.items() if isinstance(v, dict)}

    normalized: Dict[str, Dict[str, Any]] = {}
    for name, sheet in sheets.items():
        columns = sheet.get("columns") or sheet.get("headers") or []
        rows = sheet.get("rows") or []
        normalized[name] = {"columns": list(columns), "rows": list(rows)}
    return normalized


def _columns_with_data(columns: List[str], rows: List[Dict[str, Any]]) -> List[str]:
    """Return only the columns that have at least one non-empty value in the rows."""
    if not rows:
        return columns
    populated: List[str] = []
    for col in columns:
        for row in rows:
            val = row.get(col)
            if val not in (None, "", " "):
                populated.append(col)
                break
    return populated


def _write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    columns = _columns_with_data(columns, rows)
    if columns:
        ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])


def _copy_sheet_values(target_wb: Workbook, source_wb, sheet_name: str) -> None:
    if sheet_name not in source_wb.sheetnames:
        return
    source_ws = source_wb[sheet_name]
    target_ws = target_wb.create_sheet(sheet_name)
    for row in source_ws.iter_rows(values_only=True):
        target_ws.append(list(row))


def load_description_document(
    file_path: str,
    *,
    max_chars: int = 15000,
    encoding: str = "utf-8",
) -> Dict[str, Any]:
    """
    Read a long description file (txt/markdown/docx exported as text) and return a truncated preview.

    This helps seed question design from lengthy user-provided specs.
    """
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Description file not found: {path}")
    raw = path.read_text(encoding=encoding, errors="ignore")
    preview = raw[:max_chars]
    return {
        "path": str(path),
        "length": len(raw),
        "truncated": len(raw) > len(preview),
        "preview": preview,
    }


def design_survey_outline(
    topic: str,
    *,
    objectives: Optional[List[str]] = None,
    include_demographics: bool = True,
    languages: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Propose a starter set of questions and choice lists for a survey.

    Returns a structure ready to merge into a form spec: survey rows and choices rows.

    When *languages* is provided (e.g. ``["English", "French"]``), the output
    uses ``label::English (en)``, ``label::French (fr)``, etc. instead of a plain
    ``label`` column so that ODK renders a multi-language form.  All labels
    are initially set to the English text as placeholders.  **The agent must
    translate every ``label::<Language (code)>`` value into the correct language
    before writing the XLSX.**

    A language-selector question (``select_one form_language``) is
    automatically prepended so that respondents can choose their preferred
    language at the start of the form.
    """
    objectives = objectives or []
    normalized_languages = _normalize_languages(languages or [])
    language_headers = [entry["header"] for entry in normalized_languages]

    # --- label / hint helpers ---
    # When multilingual, every label::<lang> is set to the English text.
    # The LLM agent is responsible for translating these before saving.
    def _label(text: str) -> Dict[str, str]:
        if not language_headers:
            return {"label": text}
        return {f"label::{lang}": text for lang in language_headers}

    def _hint(text: str) -> Dict[str, str]:
        if not language_headers:
            return {"hint": text}
        return {f"hint::{lang}": text for lang in language_headers}

    survey_rows: List[Dict[str, Any]] = []
    choices_rows: List[Dict[str, Any]] = []

    # --- language selector (only when multilingual) ---
    if normalized_languages:
        survey_rows.append(
            {
                "type": "select_one form_language",
                "name": "form_language",
                **_label("Select your preferred language"),
                "required": "yes",
            }
        )
        for entry in normalized_languages:
            lang_value = entry["code"] or _safe_form_id(entry["header"])
            choices_rows.append(
                {"list_name": "form_language", "name": lang_value, **_label(entry["name"])}
            )

    if include_demographics:
        survey_rows.extend(
            [
                {"type": "text", "name": "respondent_name", **_label("Interviewer: enter respondent name")},
                {"type": "integer", "name": "respondent_age", **_label("Respondent age"), "constraint": ". >= 0"},
                {"type": "select_one gender_list", "name": "respondent_gender", **_label("Respondent gender")},
            ]
        )
        choices_rows.extend(
            [
                {"list_name": "gender_list", "name": "female", **_label("Female")},
                {"list_name": "gender_list", "name": "male", **_label("Male")},
                {"list_name": "gender_list", "name": "other", **_label("Other / prefer not to say")},
            ]
        )

    survey_rows.extend(
        [
            {"type": "note", "name": "intro_topic", **_label(f"Survey: {topic}")},
            {
                "type": "select_one yes_no",
                "name": "consent",
                **_label("Do you consent to participate?"),
                "required": "yes",
                "constraint_message": "Consent is required to continue.",
                "relevant": "",
            },
        ]
    )

    for idx, obj in enumerate(objectives[:5], start=1):
        survey_rows.append(
            {
                "type": "text",
                "name": f"objective_{idx}",
                **_label(f"Question about: {obj}"),
                **_hint(f"Capture details related to {obj}"),
            }
        )

    survey_rows.extend(
        [
            {
                "type": "select_one frequency_list",
                "name": "usage_frequency",
                **_label("How often do you engage with this topic?"),
            },
            {
                "type": "select_one satisfaction_list",
                "name": "satisfaction_level",
                **_label("Overall satisfaction"),
                **_hint("Quick pulse check"),
            },
            {
                "type": "text",
                "name": "open_feedback",
                **_label("Any additional feedback?"),
            },
        ]
    )

    choices_rows.extend(
        [
            {"list_name": "yes_no", "name": "yes", **_label("Yes")},
            {"list_name": "yes_no", "name": "no", **_label("No")},
            {"list_name": "frequency_list", "name": "daily", **_label("Daily")},
            {"list_name": "frequency_list", "name": "weekly", **_label("Weekly")},
            {"list_name": "frequency_list", "name": "monthly", **_label("Monthly")},
            {"list_name": "frequency_list", "name": "rarely", **_label("Rarely")},
            {"list_name": "frequency_list", "name": "never", **_label("Never")},
            {"list_name": "satisfaction_list", "name": "very_satisfied", **_label("Very satisfied")},
            {"list_name": "satisfaction_list", "name": "satisfied", **_label("Satisfied")},
            {"list_name": "satisfaction_list", "name": "neutral", **_label("Neutral")},
            {"list_name": "satisfaction_list", "name": "dissatisfied", **_label("Dissatisfied")},
            {"list_name": "satisfaction_list", "name": "very_dissatisfied", **_label("Very dissatisfied")},
        ]
    )

    languages = language_headers
    return {
        "topic": topic,
        "objectives": objectives,
        "languages": languages,
        "survey_rows": survey_rows,
        "choices_rows": choices_rows,
        "notes": "Review wording, logic, and choice list names before writing XLSX."
        + (" Translations are auto-generated for common terms; review topic-specific labels." if languages else ""),
    }


def add_calculations_and_conditions(
    form_spec: Dict[str, Any],
    calculations: List[Dict[str, Any]],
    *,
    conditions: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Add calculate rows and attach relevance/constraint logic to existing questions.

    Each item in `calculations` should include:
      - name (required)
      - calculation (required)
      - label (optional but recommended)
      - relevant / constraint / constraint_message (optional)

    Each item in `conditions` should include:
      - target (question name to update)
      - relevant / constraint / constraint_message / required (any to apply)
    """
    if not calculations:
        raise ValueError("Provide at least one calculation definition.")

    spec = _normalize_form_spec(copy.deepcopy(form_spec))
    survey = spec.setdefault(
        "survey",
        {
            "columns": list(_PREFERRED_COLUMN_ORDER["survey"]),
            "rows": [],
        },
    )

    columns = survey.get("columns") or list(_PREFERRED_COLUMN_ORDER["survey"])
    added_calc_names: List[str] = []
    for calc in calculations:
        name = calc.get("name")
        expr = calc.get("calculation")
        if not name or not expr:
            raise ValueError("Each calculation needs both 'name' and 'calculation'.")
        row = {"type": "calculate", **calc}
        row.setdefault("label", f"Calculation for {name}")
        survey.setdefault("rows", []).append(row)
        added_calc_names.append(name)
        for key in row.keys():
            if key not in columns:
                columns.append(key)

    updated_targets: List[str] = []
    if conditions:
        targets = {r.get("name"): r for r in survey.get("rows", []) if r.get("name")}
        for cond in conditions:
            target = cond.get("target")
            if not target or target not in targets:
                continue
            row = targets[target]
            for key, val in cond.items():
                if key == "target":
                    continue
                row[key] = val
                if key not in columns:
                    columns.append(key)
            updated_targets.append(target)

    # Preserve preferred ordering but keep any new columns appended.
    preferred = _PREFERRED_COLUMN_ORDER["survey"]
    ordered = [c for c in preferred if c in columns]
    ordered.extend(c for c in columns if c not in ordered)
    survey["columns"] = ordered

    return {
        "form_spec": {"sheets": spec},
        "added_calculations": added_calc_names,
        "updated_targets": updated_targets,
        "survey_columns": survey["columns"],
    }


def merge_form_spec(
    base_spec: Dict[str, Any],
    addition_spec: Dict[str, Any],
    *,
    dedupe_by_name: bool = True,
) -> Dict[str, Any]:
    """
    Merge sheets/rows/columns from addition_spec into base_spec.

    - Preserves column order using preferred templates, then appends any new columns.
    - Optionally deduplicates rows by 'name' within each sheet.
    """
    base = _normalize_form_spec(copy.deepcopy(base_spec))
    extra = _normalize_form_spec(copy.deepcopy(addition_spec))

    summary: Dict[str, Dict[str, List[str]]] = {"added": {}, "skipped": {}}

    for sheet_name, extra_sheet in extra.items():
        base_sheet = base.setdefault(sheet_name, {"columns": [], "rows": []})
        base_rows = base_sheet.get("rows") or []
        existing_names = {r.get("name") for r in base_rows if r.get("name")}

        added: List[str] = []
        skipped: List[str] = []
        for row in extra_sheet.get("rows") or []:
            name = row.get("name")
            if dedupe_by_name and name and name in existing_names:
                skipped.append(name)
                continue
            base_rows.append(row)
            if name:
                existing_names.add(name)
                added.append(name)
        base_sheet["rows"] = base_rows

        columns = base_sheet.get("columns") or []
        col_set = set(columns)
        for col in extra_sheet.get("columns") or []:
            if col not in col_set:
                columns.append(col)
                col_set.add(col)

        preferred = _PREFERRED_COLUMN_ORDER.get(sheet_name.lower())
        if preferred:
            ordered = [c for c in preferred if c in col_set]
            ordered.extend(c for c in columns if c not in ordered)
            columns = ordered
        base_sheet["columns"] = columns

        summary["added"][sheet_name] = added
        summary["skipped"][sheet_name] = skipped

    return {
        "merged_spec": {"sheets": base},
        "summary": summary,
    }


def new_form_spec(
    form_title: str,
    *,
    form_id: Optional[str] = None,
    default_language: str = "English",
    languages: Optional[List[str]] = None,
    version: Optional[str] = None,
    submission_url: Optional[str] = None,
    public_key: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Build a starter form spec with blank survey/choices sheets and basic settings.

    When *languages* is provided (e.g. ``["English", "French"]``), the survey
    and choices column lists include ``label::<Language (code)>`` columns for each
    language (e.g., ``label::English (en)``), and the settings ``default_language``
    is set to the first language.  A language-selector question is added to the
    survey sheet.

    The returned structure matches what `write_xlsform(...)` expects.
    """
    form_id = form_id or _safe_form_id(form_title)
    version = version or datetime.datetime.now().strftime("%Y%m%d%H%M")
    normalized_languages = _normalize_languages(languages or [])
    language_headers = [entry["header"] for entry in normalized_languages]

    # Build column lists — only include label::<lang> when languages are given.
    survey_cols = list(_PREFERRED_COLUMN_ORDER["survey"])
    choices_cols = list(_PREFERRED_COLUMN_ORDER["choices"])
    if language_headers:
        default_language = language_headers[0]
        # Insert label::<lang> columns right after "label" (or replace it)
        for col_list in (survey_cols, choices_cols):
            idx = col_list.index("label")
            lang_cols = [f"label::{lang}" for lang in language_headers]
            col_list[idx:idx + 1] = lang_cols

    survey_rows: List[Dict[str, Any]] = []
    choices_rows: List[Dict[str, Any]] = []

    # Add language selector when multilingual
    if language_headers:
        def _label(text: str) -> Dict[str, str]:
            return {f"label::{lang}": text for lang in language_headers}

        survey_rows.append(
            {
                "type": "select_one form_language",
                "name": "form_language",
                **_label("Select your preferred language"),
                "required": "yes",
            }
        )
        for entry in normalized_languages:
            lang_value = entry["code"] or _safe_form_id(entry["header"])
            choices_rows.append(
                {"list_name": "form_language", "name": lang_value, **_label(entry["name"])}
            )

    return {
        "sheets": {
            "survey": {
                "columns": survey_cols,
                "rows": survey_rows,
            },
            "choices": {
                "columns": choices_cols,
                "rows": choices_rows,
            },
            "settings": {
                "columns": _PREFERRED_COLUMN_ORDER["settings"],
                "rows": [
                    {
                        "form_title": form_title,
                        "form_id": form_id,
                        "default_language": default_language,
                        "version": version,
                        "submission_url": submission_url or "",
                        "public_key": public_key or "",
                    }
                ],
            },
        }
    }


def load_xlsform(
    file_path: str,
    *,
    sheet_names: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Read an existing XLSForm and return a structured, JSON-friendly spec.
    Skips fully empty rows to keep the payload concise.
    """
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"XLSForm not found: {path}")

    wb = load_workbook(path, data_only=True)
    selected = sheet_names or wb.sheetnames

    sheets: Dict[str, Dict[str, Any]] = {}
    for sheet_name in selected:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            sheets[sheet_name] = {"columns": [], "rows": []}
            continue
        header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
        header = [col for col in header if col]  # drop blanks in the header row

        parsed_rows: List[Dict[str, Any]] = []
        for raw in rows[1:]:
            if not _row_has_content(raw):
                continue
            row_dict = {header[i]: raw[i] for i in range(min(len(header), len(raw)))}
            parsed_rows.append(row_dict)

        sheets[sheet_name] = {"columns": header, "rows": parsed_rows}

    return {
        "path": str(path),
        "sheets": sheets,
        "sheet_names": list(sheets.keys()),
        "row_counts": {k: len(v["rows"]) for k, v in sheets.items()},
    }


def write_xlsform(
    form_spec: Dict[str, Any],
    output_path: Optional[str] = None,
    *,
    base_form_path: Optional[str] = None,
    preserve_additional_sheets: bool = True,
) -> Dict[str, Any]:
    """
    Save a structured form_spec to an XLSX file.

    When no base_form_path is given the official ODK XLSForm Template is
    downloaded (and cached) so that every generated file inherits the
    template's structure and formatting.

    The form_spec can be either {"sheets": {...}} or a mapping of sheet_name -> sheet_spec.
    Each sheet_spec should contain "rows" (list of dicts) and optional "columns".
    """
    spec = _normalize_form_spec(form_spec)
    if not spec:
        raise ValueError("form_spec is empty; include at least the survey sheet.")

    preferred_copy: Dict[str, List[str]] = {
        name.lower(): order for name, order in _PREFERRED_COLUMN_ORDER.items()
    }

    # Determine the base workbook: an explicit path, the official ODK
    # template (downloaded & cached), or – as last resort – a blank workbook.
    if base_form_path:
        base_path = Path(base_form_path).expanduser().resolve()
        if not base_path.exists():
            raise FileNotFoundError(f"Base form not found: {base_path}")
    else:
        base_path = _get_template_path()

    if base_path and base_path.exists():
        wb = load_workbook(base_path)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    def _prepare_sheet(sheet_name: str, sheet_spec: Dict[str, Any]) -> tuple[List[str], List[Dict[str, Any]]]:
        columns = sheet_spec.get("columns") or []
        rows = sheet_spec.get("rows") or []
        columns = columns or _infer_columns(rows, preferred_copy.get(sheet_name.lower()))
        if not columns:
            columns = ["type", "name", "label"] if sheet_name.lower() == "survey" else ["name", "label"]
        preferred = preferred_copy.get(sheet_name.lower())
        if preferred:
            ordered = [c for c in preferred if c in columns]
            ordered.extend(c for c in columns if c not in ordered)
            columns = ordered
        columns, rows = _normalize_language_columns_and_rows(columns, rows)
        return columns, rows

    normalized_sheets: Dict[str, tuple[List[str], List[Dict[str, Any]]]] = {}
    survey_languages: List[str] = []
    for sheet_name, sheet_spec in spec.items():
        if sheet_name.lower() == "settings":
            continue
        columns, rows = _prepare_sheet(sheet_name, sheet_spec)
        if sheet_name.lower() == "survey":
            survey_languages = _language_headers_from_columns(columns)
        normalized_sheets[sheet_name] = (columns, rows)

    if "settings" in spec:
        settings_spec = spec["settings"]
        settings_rows = settings_spec.get("rows") or []
        if survey_languages:
            for row in settings_rows:
                if not row.get("default_language"):
                    row["default_language"] = survey_languages[0]
        settings_columns = settings_spec.get("columns") or _PREFERRED_COLUMN_ORDER["settings"]
        settings_columns, settings_rows = _normalize_language_columns_and_rows(settings_columns, settings_rows)
        normalized_sheets["settings"] = (settings_columns, settings_rows)

    written: List[str] = []
    for sheet_name, (columns, rows) in normalized_sheets.items():
        # Reuse the existing template sheet (clearing its rows) or create new.
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.delete_rows(1, ws.max_row)
        else:
            ws = wb.create_sheet(sheet_name)

        _write_sheet(ws, columns, rows)
        written.append(sheet_name)

    if not preserve_additional_sheets:
        for name in list(wb.sheetnames):
            if name not in written:
                del wb[name]

    default_output = "odk_form_draft.xlsx"
    if base_form_path:
        bp = Path(base_form_path).expanduser().resolve()
        default_output = bp.with_name(bp.stem + "_draft.xlsx").name
    output = Path(output_path or default_output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)

    return {
        "output_path": str(output),
        "sheet_names": wb.sheetnames,
        "row_counts": {name: len(spec.get(name, {}).get("rows", [])) for name in written},
    }

def save_xlsform_draft(
    form_spec: Dict[str, Any],
    output_path: Optional[str] = None,
    *,
    base_form_path: Optional[str] = None,
    preserve_additional_sheets: bool = True,
) -> Dict[str, Any]:
    """
    Convenience wrapper that always emits an XLSX, using a timestamped filename when none is provided.
    """
    if not output_path:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = "odk_form"
        if base_form_path:
            stem = Path(base_form_path).stem + "_draft"
        output_path = f"{stem}_{ts}.xlsx"
    return write_xlsform(
        form_spec,
        output_path=output_path,
        base_form_path=base_form_path,
        preserve_additional_sheets=preserve_additional_sheets,
    )


def compare_forms(
    reference_path: str,
    candidate_path: str,
    *,
    sheet_names: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Compare a generated form against a reference XLSForm and highlight gaps.

    Returns differences in sheets, columns, choice lists, and key row counts
    (e.g., calculate rows, relevant/constraint usage).
    """
    ref = load_xlsform(reference_path, sheet_names=sheet_names)
    cand = load_xlsform(candidate_path, sheet_names=sheet_names)

    def _cols(spec: Dict[str, Any], sheet: str) -> List[str]:
        return spec["sheets"].get(sheet, {}).get("columns", [])

    ref_sheets = set(ref["sheets"].keys())
    cand_sheets = set(cand["sheets"].keys())
    missing_sheets = sorted(ref_sheets - cand_sheets)
    extra_sheets = sorted(cand_sheets - ref_sheets)

    col_gaps: Dict[str, Dict[str, List[str]]] = {}
    for sheet in ref_sheets & cand_sheets:
        ref_cols = set(_cols(ref, sheet))
        cand_cols = set(_cols(cand, sheet))
        col_gaps[sheet] = {
            "missing_in_candidate": sorted(ref_cols - cand_cols),
            "extra_in_candidate": sorted(cand_cols - ref_cols),
        }

    def _choice_lists(spec: Dict[str, Any]) -> Dict[str, int]:
        choices = spec["sheets"].get("choices", {}).get("rows", [])
        counts: Dict[str, int] = {}
        for row in choices:
            ln = row.get("list_name")
            if not ln:
                continue
            counts[ln] = counts.get(ln, 0) + 1
        return counts

    ref_choice_counts = _choice_lists(ref)
    cand_choice_counts = _choice_lists(cand)
    missing_choice_lists = sorted(set(ref_choice_counts) - set(cand_choice_counts))
    extra_choice_lists = sorted(set(cand_choice_counts) - set(ref_choice_counts))

    def _survey_stats(spec: Dict[str, Any]) -> Dict[str, int]:
        rows = spec["sheets"].get("survey", {}).get("rows", [])
        return {
            "total": len(rows),
            "calculate_rows": sum(1 for r in rows if str(r.get("type", "")).lower() == "calculate"),
            "with_relevant": sum(1 for r in rows if r.get("relevant")),
            "with_constraint": sum(1 for r in rows if r.get("constraint")),
        }

    return {
        "reference_path": ref["path"],
        "candidate_path": cand["path"],
        "sheet_names_checked": list(ref_sheets | cand_sheets),
        "missing_sheets": missing_sheets,
        "extra_sheets": extra_sheets,
        "column_gaps": col_gaps,
        "choice_lists": {
            "reference": ref_choice_counts,
            "candidate": cand_choice_counts,
            "missing_in_candidate": missing_choice_lists,
            "extra_in_candidate": extra_choice_lists,
        },
        "survey_stats": {
            "reference": _survey_stats(ref),
            "candidate": _survey_stats(cand),
        },
    }
INSTRUCTION = """
You are odk_xlsform_agent. Help users design or edit ODK XLSForms step by step.

## Getting started
- Start by asking whether they want to create a new form or modify an existing one. If modifying, load it with `load_xlsform(...)` and summarize what you find before changing anything.
- **Always ask which languages the form should support.** Supported languages include English, French, German, Spanish, Portuguese, Arabic, Swahili, Chinese, etc. When the user picks languages, pass them as the `languages` parameter to `new_form_spec(...)` and `design_survey_outline(...)`.

## Multi-language support
- When languages are specified, use `label::<Language (code)>` columns (e.g. `label::English (en)`, `label::French (fr)`) instead of a plain `label` column. The same applies to `hint::<Language (code)>` and `constraint_message::<Language (code)>`.
- **Never include `label::<Language (code)>` columns if they will be empty.** Only include language columns that are actually populated. Empty language or media columns cause upload errors in ODK.
- The user can choose **any** languages they want — there is no fixed list. Pass the language names to `new_form_spec(languages=...)` and `design_survey_outline(languages=...)` to set up the correct `label::<Language (code)>` column structure (language names are normalized to include their IANA subtags).
- The tools create all `label::<Language (code)>` values as **English placeholders**. **You MUST translate every `label::<Language (code)>` value into the correct language before calling `write_xlsform` or `save_xlsform_draft`.** For example, if languages are English and French, set `label::English (en)` to the English text and `label::French (fr)` to the proper French translation. Do this for survey rows AND choices rows. You are an LLM — use your language abilities to produce accurate translations.
- Before saving or sharing any spec or XLSX, double-check that every language-bearing column includes the code suffix (e.g., `label::English (en)`) and normalize/fix it if missing.
- A language-selector question (`select_one form_language`) is automatically added at the start of the form when languages are specified.
- Set `default_language` in the settings sheet to the first language in the list.

## Designing questions
- Offer to draft an initial question set with `design_survey_outline(...)`; include types and choice lists so the user can confirm before adding to the form.
- If the user provides a long description or document path, ingest it with `load_description_document(...)`, summarize key sections, and confirm understanding before generating questions.
- When users mention derived values or skip logic, propose calculate rows and relevance/constraint logic with `add_calculations_and_conditions(...)` and show the updated survey columns.
- When users provide a reference XLSForm, compare it with `compare_forms(...)` to surface missing sheets/columns/choice lists and borrow patterns (e.g., age calculations, consent gating) before rewriting the draft.
- Use `merge_form_spec(...)` to combine drafts with user-provided specs while avoiding duplicate names and keeping column order sane.

## Form management
- Keep a working form spec (survey rows, choice lists, settings). After each batch of changes, show the current rows and ask the user to double-check logic, relevance, and constraints.
- Use `new_form_spec(...)` to scaffold settings, then gather questions, choice options, groups/repeats, and skip logic one section at a time.

## Saving / exporting
- When the user confirms or asks to save/export, call `save_xlsform_draft(...)` (or `write_xlsform(...)`) to emit an XLSX; default to a timestamped filename if none is provided and report the path back. Every generated file is automatically based on the official ODK XLSForm Template (https://github.com/getodk/xlsform-template) so it inherits proper structure and formatting.
- **Only include columns that have data.** Empty columns (especially `media::image`, `media::audio`, `media::video`, and unpopulated `label::<Language (code)>` columns) cause ODK validation errors and must not be written.

## Column conventions
- Follow XLSForm column conventions: type/name/label, language-specific labels (label::English (en) only when that language is in use), hints, required, relevant, constraint and constraint_message, calculation, appearance, choice_filter, repeat_count, default.
- Only include media columns (media::image, media::audio, media::video) when media files are actually referenced.
- For select_one/select_multiple questions, always ensure the matching choice list exists and is spelled correctly.
- Decide and support groups/repeats via begin_group/end_group and begin_repeat/end_repeat rows; keep grouping balanced.
- If something seems ambiguous or risky, ask clarifying questions instead of guessing.
""".strip()

root_agent = Agent(
    name="odk_xlsform_agent",
    model="gemini-2.5-flash",
    instruction=INSTRUCTION,
    tools=[
        load_description_document,
        design_survey_outline,
        merge_form_spec,
        add_calculations_and_conditions,
        compare_forms,
        new_form_spec,
        load_xlsform,
        write_xlsform,
        save_xlsform_draft,
    ],
)
