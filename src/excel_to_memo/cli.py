from __future__ import annotations

import argparse
import json
from pathlib import Path

from .renderers import export_normalized_rows, render_docx, render_markdown
from .transform import build_report, load_config, load_dataframe, report_to_dict


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a monitoring-detail Excel or CSV file into a memo-style report."
    )
    parser.add_argument("input_file", help="Path to the source Excel/CSV file.")
    parser.add_argument(
        "--config",
        help="Optional JSON config file with metadata and column mapping.",
    )
    parser.add_argument(
        "--sheet",
        help="Optional Excel sheet name. If omitted, the first sheet is used.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Directory where the memo files will be written.",
    )
    parser.add_argument(
        "--json-summary",
        action="store_true",
        help="Also export the resolved memo payload as JSON.",
    )
    return parser


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
            parser.error(
                "Config file was not found: %s. Try --config config.json or "
                "--config examples/sample_config.json" % config_arg
            )
    elif Path("config.json").exists():
        config_arg = "config.json"

    config = load_config(config_arg)
    if args.sheet:
        config.setdefault("input", {})
        config["input"]["sheet_name"] = args.sheet

    dataframe = load_dataframe(
        input_path=input_path,
        sheet_name=config.get("input", {}).get("sheet_name"),
    )
    report = build_report(dataframe, config)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    markdown_path = render_markdown(report, output_dir / "memo_review.md")
    docx_path = render_docx(report, output_dir / "memo_review.docx")
    csv_path = export_normalized_rows(report, output_dir / "normalized_rows.csv")

    print(f"Wrote {markdown_path}")
    print(f"Wrote {docx_path}")
    print(f"Wrote {csv_path}")

    if args.json_summary:
        json_path = output_dir / "memo_payload.json"
        json_path.write_text(
            json.dumps(report_to_dict(report), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"Wrote {json_path}")


if __name__ == "__main__":
    main()
