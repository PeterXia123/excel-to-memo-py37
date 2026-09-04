from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Union

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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build memo comments text directly from an Excel/CSV file."
    )
    parser.add_argument("input_file", help="Path to the source Excel/CSV file.")
    parser.add_argument(
        "--config",
        help="Optional JSON config file. If omitted and config.json exists, it is used.",
    )
    parser.add_argument(
        "--sheet",
        help="Optional Excel sheet name. If omitted, the first sheet is used.",
    )
    parser.add_argument(
        "--output-file",
        help="Optional text file path. If omitted, text is printed to stdout.",
    )
    return parser


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
    raise ValueError("Unsupported input format: %s" % suffix)


def build_comments_text_from_excel(
    input_path: Union[str, Path],
    config_path: Optional[Union[str, Path]] = None,
    sheet_name: Optional[str] = None,
) -> str:
    config = load_config(config_path)
    if sheet_name:
        config.setdefault("input", {})
        config["input"]["sheet_name"] = sheet_name

    dataframe = load_dataframe(
        input_path=input_path,
        sheet_name=config.get("input", {}).get("sheet_name"),
    )
    return build_comments_text_from_dataframe(dataframe, config)


def build_comments_text_from_dataframe(
    dataframe: pd.DataFrame,
    config: Optional[Dict[str, Any]] = None,
) -> str:
    resolved_config = config or {}
    input_config = resolved_config.get("input", {})
    report_config = resolved_config.get("report", {})

    normalized = normalize_dataframe(
        dataframe=dataframe,
        explicit_columns=input_config.get("columns", {}),
        alias_overrides=input_config.get("column_aliases", {}),
        forward_fill_columns=input_config.get(
            "forward_fill_columns", DEFAULT_FORWARD_FILL_COLUMNS
        ),
    )

    bullets = build_section_bullets(normalized, input_config.get("group_by", "section"))
    intro_paragraphs = build_intro_paragraphs(report_config)
    closing_paragraphs = build_closing_paragraphs(report_config)

    return join_comments_blocks(intro_paragraphs, bullets, closing_paragraphs)


def build_intro_paragraphs(report_config: Dict[str, Any]) -> List[str]:
    if not report_config.get("include_intro", False):
        return []

    applicable_models = report_config.get(
        "applicable_models", DEFAULT_APPLICABLE_MODELS
    )
    model_owner = report_config.get("model_owner", "TBD")

    return [
        report_config.get(
            "intro_template",
            "MVA has reviewed the Model Monitoring Report submitted by "
            "{model_owner} regarding {applicable_models}.",
        ).format(
            model_owner=model_owner,
            applicable_models=applicable_models,
        ),
        report_config.get(
            "conclusion_template",
            "MVA has assessed the results from the model monitoring report and concludes the following:",
        ),
    ]


def build_closing_paragraphs(report_config: Dict[str, Any]) -> List[str]:
    if not report_config.get("include_outro", report_config.get("include_intro", False)):
        return []
    return [report_config.get("outro_template", DEFAULT_COMMENTS_OUTRO)]


def join_comments_blocks(
    intro_paragraphs: List[str],
    bullets: List[str],
    closing_paragraphs: List[str],
    bullet_prefix: str = "- ",
) -> str:
    sections: List[str] = []

    if intro_paragraphs:
        sections.append("\n\n".join(intro_paragraphs))

    if bullets:
        sections.append(
            "\n".join(
                "{prefix}{bullet}".format(prefix=bullet_prefix, bullet=bullet)
                for bullet in bullets
            )
        )

    if closing_paragraphs:
        sections.append("\n\n".join(closing_paragraphs))

    return "\n\n".join(section for section in sections if section)


def normalize_dataframe(
    dataframe: pd.DataFrame,
    explicit_columns: Dict[str, str],
    alias_overrides: Dict[str, List[str]],
    forward_fill_columns: List[str],
) -> pd.DataFrame:
    resolved_columns = {}
    combined_aliases = {}
    for key, value in DEFAULT_ALIASES.items():
        combined_aliases[key] = list(value) + list(alias_overrides.get(key, []))

    normalized_source_headers = {}
    for column in dataframe.columns.astype(str):
        normalized_source_headers[normalize_header(column)] = column

    for canonical_name, source_name in explicit_columns.items():
        if source_name not in dataframe.columns:
            raise KeyError(
                "Configured column '%s' was not found in the input file." % source_name
            )
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
    return normalized


def build_section_bullets(normalized: pd.DataFrame, group_by: str) -> List[str]:
    key = group_by.lower().replace(" ", "_").replace("-", "_")
    if key != "section" or key not in normalized.columns:
        key = "group_key"

    ordered = normalized.sort_values(by=["row_order"], ascending=[True])
    bullets = []
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
            body = "%s: %s" % (label, body)
        bullets.append(ensure_terminal_punctuation(body))
    return bullets


def group_has_fail(group_frame: pd.DataFrame) -> bool:
    return any(status in FAIL_STATUSES for status in group_frame["status"].tolist())


def build_fail_group_body(group_frame: pd.DataFrame) -> str:
    fail_rows = group_frame.loc[group_frame["status"].isin(FAIL_STATUSES)].sort_values(
        by=["row_order"]
    )
    comments = dedupe_exact_preserve_order(
        item for item in fail_rows["comments"].tolist() if not is_blank_like(item)
    )
    if comments:
        return " ".join(ensure_terminal_punctuation(item) for item in comments)

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


def structured_fail_sentence(row: pd.Series) -> str:
    parameter = first_non_blank(row.get("model_parameter", ""), "Parameter")
    portfolio = row.get("portfolio", "")
    estimate_direction = normalize_estimate_direction(row.get("estimate_direction", ""))
    ecl_impact = row.get("ecl_impact", "")
    materiality = row.get("materiality", "")

    subject = "The %s parameter" % parameter
    if not is_blank_like(portfolio):
        subject = "The %s parameter for %s" % (parameter, portfolio)

    parts = [subject]
    if not is_blank_like(estimate_direction):
        parts.append("is %s" % estimate_direction)
    else:
        parts.append("failed review")

    if not is_blank_like(materiality):
        parts.append("with materiality of %s" % materiality)

    if not is_blank_like(ecl_impact):
        parts.append("and an ECL impact of %s" % ecl_impact)

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

    subject = "The %s parameter" % parameter
    if not is_blank_like(portfolio):
        subject = "The %s parameter for %s" % (parameter, portfolio)

    if not is_blank_like(estimate_direction) and not is_blank_like(ecl_impact):
        return ensure_terminal_punctuation(
            "%s is %s with an ECL impact of %s"
            % (subject, estimate_direction, ecl_impact)
        )

    if not is_blank_like(estimate_direction):
        return ensure_terminal_punctuation("%s is %s" % (subject, estimate_direction))

    if not is_blank_like(ecl_impact):
        return ensure_terminal_punctuation("%s has an ECL impact of %s" % (subject, ecl_impact))

    if status_raw == "acceptable":
        return ensure_terminal_punctuation("%s is acceptable" % subject)

    if status in PASS_STATUSES:
        return ensure_terminal_punctuation("%s passed review" % subject)

    if status in WARNING_STATUSES:
        return ensure_terminal_punctuation("%s is flagged as a warning" % subject)

    if status in FAIL_STATUSES:
        return ensure_terminal_punctuation("%s failed review" % subject)

    return ensure_terminal_punctuation("%s was assessed" % subject)


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
        values = []
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


def row_is_blank(row: pd.Series) -> bool:
    return all(is_blank_like(value) for value in row.tolist())


def is_blank_like(value: Any) -> bool:
    return normalize_whitespace(str(value)).lower() in BLANK_MARKERS


def dedupe_exact_preserve_order(values: Iterable[str]) -> List[str]:
    seen = set()
    deduped = []
    for item in values:
        normalized = normalize_whitespace(item)
        key = normalized.lower()
        if not normalized or key in seen:
            continue
        deduped.append(normalized)
        seen.add(key)
    return deduped


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def ensure_terminal_punctuation(value: str) -> str:
    if not value:
        return value
    if value.endswith((".", "!", "?")):
        return value
    return "%s." % value


def first_non_blank(*values: str) -> str:
    for value in values:
        if not is_blank_like(value):
            return value
    return ""


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input_file)
    if not input_path.exists():
        parser.error("Input file was not found: %s" % args.input_file)

    config_arg = args.config
    if config_arg:
        config_path = Path(config_arg)
        if not config_path.exists():
            parser.error("Config file was not found: %s" % config_arg)
    elif Path("config.json").exists():
        config_arg = "config.json"

    comments_text = build_comments_text_from_excel(
        input_path=input_path,
        config_path=config_arg,
        sheet_name=args.sheet,
    )

    if args.output_file:
        output_path = Path(args.output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(comments_text, encoding="utf-8")
        print("Wrote %s" % output_path)
        return

    print(comments_text)


if __name__ == "__main__":
    main()
