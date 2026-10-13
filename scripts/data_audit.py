"""Run the source-file audit and write data_audit_report.txt."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_project import audit, load_data  # noqa: E402


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    (root / "data_audit_report.txt").write_text(audit(load_data()), encoding="utf-8")
    print(root / "data_audit_report.txt")
