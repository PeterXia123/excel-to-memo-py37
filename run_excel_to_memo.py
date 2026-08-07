#!/usr/bin/env python3

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_DIR = REPO_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

try:
    from excel_to_memo.cli import main
except ModuleNotFoundError as error:
    missing = getattr(error, "name", "a required package")
    print(
        "Missing runtime dependency: %s\n"
        "Install the runtime dependencies first, then run this script again.\n\n"
        "Recommended command:\n"
        "  pip install -r requirements-py37.txt\n"
        % missing,
        file=sys.stderr,
    )
    raise SystemExit(1)


if __name__ == "__main__":
    main()
