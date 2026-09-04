from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd


DEFAULT_ALIASES = {
    "section": ["section"],
    "subsection": ["sub-section", "sub section", "subsection"],
    "model_name": ["model name", "model_name"],
    "model_parameter": ["model parameter", "parameter", "model_parameter"],
    "portfolio": ["portfolio"],
    "status": ["status"],
    "materiality": ["materiality"],
    "estimate_direction": [
        "over/under estimate",
        "overunder estimate",
        "over under estimate",
        "estimate direction",
    ],
    "ecl_impact": ["ecl impact", "ecl_impact"],
    "comments": ["comments", "comment"],
}

DEFAULT_APPLICABLE_MODELS = "IFRS9 International Banking Non-FLI ECL Models"
DEFAULT_COMMENTS_OUTRO = (
    "Models not mentioned above are either unmonitored or have successfully "
    "completed the necessary tests. The table below provides specific "
    "monitoring information for each model."
)
BLANK_MARKERS = {"", "-", "--", "n/a", "na", "none", "null", "nan"}
PASS_STATUSES = {"pass", "passed", "ok", "acceptable"}
WARNING_STATUSES = {"warning", "warn", "yellow", "watch"}
FAIL_STATUSES = {"fail", "failed", "breach", "exception", "red"}
DEFAULT_FORWARD_FILL_COLUMNS = [
    "section",
    "subsection",
    "model_name",
    "model_parameter",
]


@dataclass
class MemoReport:
    title: str
    monitoring_summary: Dict[str, str]
    review_summary: Dict[str, str]
    requirements: List[str]
    intro_paragraphs: List[str]
    closing_paragraphs: List[str]
    bullets: List[str]
    normalized_rows: List[Dict[str, str]]


def load_config(path: Optional[Union[str, Path]]) -> Dict[str, Any]:
    if not path:
        return {}
    with Path(path).open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_dataframe(
    input_path: Union[str, Path], sheet_name: Optional[str] = None
) -> pd.DataFrame:
    path = Path(input_path)
    suffix = path.suffix.lower()
    if suffix in {".csv", ".tsv"}:
        sep = "\t" if suffix == ".tsv" else ","
        return pd.read_csv(path, sep=sep)
    if suffix in {".xlsx", ".xls"}:
        effective_sheet = 0 if sheet_name is None else sheet_name
        return pd.read_excel(path, sheet_name=effective_sheet)
    raise ValueError(f"Unsupported input format: {suffix}")


def build_report(dataframe: pd.DataFrame, config: Dict[str, Any]) -> MemoReport:
    input_config = config.get("input", {})
    report_config = config.get("report", {})

    normalized = normalize_dataframe(
        dataframe=dataframe,
        explicit_columns=input_config.get("columns", {}),
        alias_overrides=input_config.get("column_aliases", {}),
        forward_fill_columns=input_config.get("forward_fill_columns", DEFAULT_FORWARD_FILL_COLUMNS),
    )

    bullets = build_section_bullets(normalized, input_config.get("group_by", "section"))
    assessment = infer_assessment(normalized)
    persistent_breach = infer_persistent_breach(normalized)

    monitoring_summary = {
        "Applicable Model(s)": report_config.get("applicable_models", DEFAULT_APPLICABLE_MODELS),
        "Model Owner": report_config.get("model_owner", "TBD"),
        "Date Report Received": report_config.get("date_report_received", "TBD"),
        "Performance Assessment Date": report_config.get("performance_assessment_date", "TBD"),
    }

    review_summary = {
        "Monitoring Results Assessment": report_config.get(
            "monitoring_results_assessment", assessment
        ),
        "Assessment Comments": report_config.get("assessment_comments", "See comments below."),
        "Persistent Breaches Identified": report_config.get(
            "persistent_breaches_identified",
            "Yes - Model Exception Required" if persistent_breach else "No",
        ),
        "Trigger Earlier Review": report_config.get(
            "trigger_earlier_review",
            "Yes" if assessment == "Needs Attention" else "No",
        ),
        "MVA Sign-off": report_config.get("mva_signoff", "TBD"),
    }

    intro_paragraphs: List[str] = []
    if report_config.get("include_intro", True):
        intro_paragraphs = [
            report_config.get(
                "intro_template",
                "MVA has reviewed the Model Monitoring Report submitted by "
                "{model_owner} regarding {applicable_models}.",
            ).format(
                model_owner=monitoring_summary["Model Owner"],
                applicable_models=monitoring_summary["Applicable Model(s)"],
            ),
            report_config.get(
                "conclusion_template",
                "MVA has assessed the results from the model monitoring report and concludes the following:",
            ),
        ]

    closing_paragraphs: List[str] = []
    if report_config.get("include_outro", report_config.get("include_intro", True)):
        closing_paragraphs = [
            report_config.get("outro_template", DEFAULT_COMMENTS_OUTRO),
        ]

    requirements = report_config.get(
        "requirements",
        [
            "Monitoring Frequency meets minimum requirements",
            "Nature and type of monitoring is aligned with type of model and its output characteristics",
        ],
    )

    return MemoReport(
        title=report_config.get("title", "MEMO: Model Monitoring Report Review"),
        monitoring_summary=monitoring_summary,
        review_summary=review_summary,
        requirements=requirements,
        intro_paragraphs=intro_paragraphs,
        closing_paragraphs=closing_paragraphs,
        bullets=bullets,
        normalized_rows=normalized.fillna("").to_dict(orient="records"),
    )


def normalize_dataframe(
    dataframe: pd.DataFrame,
    explicit_columns: Dict[str, str],
    alias_overrides: Dict[str, List[str]],
    forward_fill_columns: List[str],
) -> pd.DataFrame:
    resolved_columns: Dict[str, str] = {}
    combined_aliases = {
        key: [*value, *alias_overrides.get(key, [])]
        for key, value in DEFAULT_ALIASES.items()
    }

    normalized_source_headers = {
        normalize_header(column): column for column in dataframe.columns.astype(str)
    }

    for canonical_name, source_name in explicit_columns.items():
        if source_name not in dataframe.columns:
            raise KeyError(f"Configured column '{source_name}' was not found in the input file.")
        resolved_columns[canonical_name] = source_name

    for canonical_name, aliases in combined_aliases.items():
        if canonical_name in resolved_columns:
            continue
        for alias in aliases:
            source_name = normalized_source_headers.get(normalize_header(alias))
            if source_name is not None:
                resolved_columns[canonical_name] = source_name
                break

    normalized = pd.DataFrame()
    for canonical_name in DEFAULT_ALIASES:
        source_name = resolved_columns.get(canonical_name)
        if source_name is None:
            normalized[canonical_name] = ""
        else:
            normalized[canonical_name] = dataframe[source_name]

    for column in normalized.columns:
        normalized[column] = normalized[column].map(clean_cell)

    apply_forward_fill(normalized, forward_fill_columns)

    normalized = normalized.loc[~normalized.apply(row_is_blank, axis=1)].copy()
    normalized["row_order"] = range(len(normalized))
    normalized["status_raw"] = normalized["status"]
    normalized["status"] = normalized["status"].map(normalize_status)
    normalized["group_key"] = normalized["section"].map(
        lambda value: first_non_blank(value, "Unlabeled Section")
    )
    normalized["severity_rank"] = normalized["status"].map(status_rank)
    return normalized


def build_section_bullets(normalized: pd.DataFrame, group_by: str) -> List[str]:
    key = group_by.lower().replace(" ", "_").replace("-", "_")
    if key != "section" or key not in normalized.columns:
        key = "group_key"

    ordered = normalized.sort_values(by=["row_order"], ascending=[True])
    bullets: List[str] = []
    for group_value, group_frame in ordered.groupby(key, sort=False, dropna=False):
        label = clean_cell(group_value) or "Ungrouped Items"
        if group_has_fail(group_frame):
            body = build_fail_group_body(group_frame)
        else:
            body = build_non_fail_group_body(group_frame)

        body = normalize_whitespace(body)
        if not body:
            body = "No narrative comments were available for this section."
        if not body.lower().startswith(label.lower()):
            body = f"{label}: {body}"
        bullets.append(ensure_terminal_punctuation(body))
    return bullets


def infer_assessment(normalized: pd.DataFrame) -> str:
    statuses = {value for value in normalized["status"].tolist() if value}
    if any(status in FAIL_STATUSES for status in statuses):
        return "Needs Attention"
    return "Acceptable"


def infer_persistent_breach(normalized: pd.DataFrame) -> bool:
    for _, row in normalized.iterrows():
        comment = row.get("comments", "").lower()
        status = row.get("status", "")
        materiality = row.get("materiality", "")
        ecl_impact = row.get("ecl_impact", "")
        if "persistent breach" in comment:
            return True
        if status in FAIL_STATUSES | WARNING_STATUSES and (
            not is_blank_like(materiality) or not is_blank_like(ecl_impact)
        ):
            return True
    return False
def group_has_fail(group_frame: pd.DataFrame) -> bool:
    return any(status in FAIL_STATUSES for status in group_frame["status"].tolist())


def build_fail_group_body(group_frame: pd.DataFrame) -> str:
    fail_rows = group_frame.loc[group_frame["status"].isin(FAIL_STATUSES)].sort_values(
        by=["row_order"]
    )
    sentences = dedupe_exact_preserve_order(
        structured_fail_sentence(row) for _, row in fail_rows.iterrows()
    )
    return " ".join(sentences)


def build_non_fail_group_body(group_frame: pd.DataFrame) -> str:
    comments = dedupe_exact_preserve_order(
        item for item in group_frame["comments"].tolist() if not is_blank_like(item)
    )
    if comments:
        return " ".join(ensure_terminal_punctuation(item) for item in comments)

    ordered_rows = group_frame.sort_values(by=["row_order"])
    sentences = dedupe_exact_preserve_order(
        structured_parameter_sentence(row) for _, row in ordered_rows.iterrows()
    )
    return " ".join(sentences)

def structured_sentence(row: pd.Series) -> str:
    parameter = first_non_blank(row.get("model_parameter", ""), "Parameter")
    portfolio = row.get("portfolio", "")
    status = row.get("status", "")
    estimate_direction = row.get("estimate_direction", "")
    ecl_impact = row.get("ecl_impact", "")
    materiality = row.get("materiality", "")

    subject = parameter
    if not is_blank_like(portfolio):
        subject = f"{parameter} for {portfolio}"

    parts = []
    if status in PASS_STATUSES and is_blank_like(estimate_direction) and is_blank_like(ecl_impact):
        parts.append(f"{subject} performed within expected thresholds")
    elif status in PASS_STATUSES:
        parts.append(f"{subject} passed review")
    elif status in WARNING_STATUSES:
        parts.append(f"{subject} is flagged as a warning")
    elif status in FAIL_STATUSES:
        parts.append(f"{subject} failed review")
    else:
        parts.append(f"{subject} was assessed")

    if not is_blank_like(estimate_direction) and not is_blank_like(ecl_impact):
        parts.append(f"it is {estimate_direction} with ECL impact of {ecl_impact}")
    elif not is_blank_like(estimate_direction):
        parts.append(f"it is {estimate_direction}")
    elif not is_blank_like(ecl_impact):
        parts.append(f"ECL impact is {ecl_impact}")

    if not is_blank_like(materiality):
        parts.append(f"materiality is {materiality}")

    sentence = ", ".join(parts)
    return ensure_terminal_punctuation(sentence[:1].upper() + sentence[1:])


def structured_fail_sentence(row: pd.Series) -> str:
    parameter = first_non_blank(row.get("model_parameter", ""), "Parameter")
    portfolio = row.get("portfolio", "")
    estimate_direction = normalize_estimate_direction(row.get("estimate_direction", ""))
    ecl_impact = row.get("ecl_impact", "")
    materiality = row.get("materiality", "")

    subject = f"The {parameter} parameter"
    if not is_blank_like(portfolio):
        subject = f"The {parameter} parameter for {portfolio}"

    parts = [subject]
    if not is_blank_like(estimate_direction):
        parts.append(f"is {estimate_direction}")
    else:
        parts.append("failed review")

    if not is_blank_like(materiality):
        parts.append(f"with materiality of {materiality}")

    if not is_blank_like(ecl_impact):
        parts.append(f"and an ECL impact of {ecl_impact}")

    sentence = " ".join(parts)
    sentence = sentence.replace("review and an ECL impact", "review with an ECL impact")
    return ensure_terminal_punctuation(sentence)


def structured_parameter_sentence(row: pd.Series) -> str:
    parameter = first_non_blank(row.get("model_parameter", ""), "Parameter")
    portfolio = row.get("portfolio", "")
    estimate_direction = normalize_estimate_direction(row.get("estimate_direction", ""))
    ecl_impact = row.get("ecl_impact", "")
    status = row.get("status", "")
    status_raw = normalize_whitespace(row.get("status_raw", "")).lower()

    subject = f"The {parameter} parameter"
    if not is_blank_like(portfolio):
        subject = f"The {parameter} parameter for {portfolio}"

    if not is_blank_like(estimate_direction) and not is_blank_like(ecl_impact):
        return ensure_terminal_punctuation(
            f"{subject} is {estimate_direction} with an ECL impact of {ecl_impact}"
        )

    if not is_blank_like(estimate_direction):
        return ensure_terminal_punctuation(f"{subject} is {estimate_direction}")

    if not is_blank_like(ecl_impact):
        return ensure_terminal_punctuation(f"{subject} has an ECL impact of {ecl_impact}")

    if status_raw == "acceptable":
        return ensure_terminal_punctuation(f"{subject} is acceptable")

    if status in PASS_STATUSES:
        return ensure_terminal_punctuation(f"{subject} passed review")

    if status in WARNING_STATUSES:
        return ensure_terminal_punctuation(f"{subject} is flagged as a warning")

    if status in FAIL_STATUSES:
        return ensure_terminal_punctuation(f"{subject} failed review")

    return ensure_terminal_punctuation(f"{subject} was assessed")


def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.strip().lower()).strip()


def clean_cell(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return normalize_whitespace(str(value))


def apply_forward_fill(dataframe: pd.DataFrame, columns: List[str]) -> None:
    for column in columns:
        if column not in dataframe.columns:
            continue
        values: List[str] = []
        last_seen = ""
        for raw_value in dataframe[column].tolist():
            current = clean_cell(raw_value)
            if is_blank_like(current):
                values.append(last_seen)
            else:
                last_seen = current
                values.append(current)
        dataframe[column] = values


def normalize_status(value: str) -> str:
    cleaned = value.lower().strip()
    if cleaned in PASS_STATUSES:
        return "pass"
    if cleaned in WARNING_STATUSES:
        return "warning"
    if cleaned in FAIL_STATUSES:
        return "fail"
    return cleaned


def normalize_estimate_direction(value: str) -> str:
    cleaned = normalize_whitespace(value).lower()
    if cleaned in {"under", "underestimate", "underestimated"}:
        return "underestimated"
    if cleaned in {"over", "overestimate", "overestimated"}:
        return "overestimated"
    return cleaned


def status_rank(value: str) -> int:
    if value in FAIL_STATUSES:
        return 3
    if value in WARNING_STATUSES:
        return 2
    if value in PASS_STATUSES:
        return 1
    return 0


def row_is_blank(row: pd.Series) -> bool:
    return all(is_blank_like(value) for value in row.tolist())


def is_blank_like(value: Any) -> bool:
    return normalize_whitespace(str(value)).lower() in BLANK_MARKERS


def dedupe_preserve_order(values: Any) -> List[str]:
    seen = set()
    deduped: List[str] = []
    for item in values:
        normalized = normalize_whitespace(item)
        key = normalized.lower()
        if not normalized or key in seen:
            continue
        deduped.append(normalized)
        seen.add(key)
    return deduped


def dedupe_exact_preserve_order(values: Any) -> List[str]:
    seen = set()
    deduped: List[str] = []
    for item in values:
        normalized = normalize_whitespace(item)
        key = normalized.lower()
        if not normalized or key in seen:
            continue
        deduped.append(normalized)
        seen.add(key)
    return deduped


def prune_nested_comments(values: List[str]) -> List[str]:
    pruned: List[str] = []
    for index, current in enumerate(values):
        current_lower = current.lower()
        contained_elsewhere = False
        for other_index, other in enumerate(values):
            if index == other_index:
                continue
            other_lower = other.lower()
            if current_lower == other_lower:
                continue
            if current_lower in other_lower:
                contained_elsewhere = True
                break
        if not contained_elsewhere:
            pruned.append(current)
    return pruned


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def ensure_terminal_punctuation(value: str) -> str:
    if not value:
        return value
    if value.endswith((".", "!", "?")):
        return value
    return f"{value}."


def first_non_blank(*values: str) -> str:
    for value in values:
        if not is_blank_like(value):
            return value
    return ""


def build_comments_text(report: MemoReport, bullet_prefix: str = "- ") -> str:
    sections: List[str] = []

    if report.intro_paragraphs:
        sections.append("\n\n".join(report.intro_paragraphs))

    if report.bullets:
        sections.append(
            "\n".join(
                "{prefix}{bullet}".format(prefix=bullet_prefix, bullet=bullet)
                for bullet in report.bullets
            )
        )

    if report.closing_paragraphs:
        sections.append("\n\n".join(report.closing_paragraphs))

    return "\n\n".join(section for section in sections if section)


def report_to_dict(report: MemoReport) -> Dict[str, Any]:
    return asdict(report)
