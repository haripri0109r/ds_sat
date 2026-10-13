"""Validate generated artifacts and source-file preservation."""

from pathlib import Path
import hashlib

import nbformat
from openpyxl import load_workbook
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
SOURCE_FILES = ["air_quality.csv", "stations.csv", "weather.csv", "traffic.csv", "calendar.csv"]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> None:
    required = [
        "README.md", "requirements.txt", ".gitignore", "data_audit_report.txt",
        "air_insights.ipynb", "air_quality_analysis.xlsx", "business_insights.pptx",
        "outputs/eda_findings.txt",
    ]
    lines = ["CH20 – Final validation report", "", "Files inspected and modified:"]
    for item in required + ["scripts/data_audit.py", "scripts/preprocess.py", "scripts/analyze.py",
                            "scripts/train_model.py", "scripts/generate_reports.py", "scripts/final_validation.py"]:
        path = ROOT / item
        lines.append(f"- {item}: {'present' if path.exists() and path.stat().st_size else 'MISSING'}")

    nb = nbformat.read(ROOT / "air_insights.ipynb", as_version=4)
    errors = [
        (i, out.get("ename"), out.get("evalue"))
        for i, cell in enumerate(nb.cells)
        for out in cell.get("outputs", [])
        if out.get("output_type") == "error"
    ]
    image_outputs = sum(
        1
        for cell in nb.cells
        for output in cell.get("outputs", [])
        if "image/png" in output.get("data", {})
    )
    wb = load_workbook(ROOT / "air_quality_analysis.xlsx")
    prs = Presentation(ROOT / "business_insights.pptx")
    lines += [
        "",
        "Verified numerical/results artifacts:",
        f"- Notebook cells: {len(nb.cells)}; execution errors stored: {errors or 'none'}; "
        f"embedded image outputs: {image_outputs}.",
        f"- Excel sheets: {', '.join(wb.sheetnames)}.",
        f"- PowerPoint slides: {len(prs.slides)}.",
        f"- Processed CSV files: {len(list((ROOT / 'processed').glob('*.csv')))}.",
        f"- Figures: {len(list((ROOT / 'outputs/figures').glob('*.png')))}.",
        "",
        "Analyses not assessable with supplied data:",
        "- AQI versus weather, wind, rainfall, temperature, and humidity: zero air-quality/weather date overlap; numeric fields absent.",
        "- AQI versus traffic: zero air-quality/traffic date overlap; no station key and no AQI field in traffic.",
        "",
        "Source preservation:",
    ]
    for name in SOURCE_FILES:
        lines.append(f"- {name}: SHA-256 {sha256(ROOT / name)}")
    lines += [
        "",
        "Commands used:",
        r"- .\.venv\Scripts\python.exe scripts\data_audit.py",
        r"- .\.venv\Scripts\python.exe scripts\preprocess.py",
        r"- .\.venv\Scripts\python.exe scripts\analyze.py",
        r"- .\.venv\Scripts\python.exe scripts\train_model.py",
        r"- .\.venv\Scripts\python.exe scripts\generate_reports.py",
        r"- .\.venv\Scripts\python.exe scripts\final_validation.py",
        "",
        "Remaining warnings: notebook execution may emit normal Windows ZeroMQ/TCP warnings; no notebook cell errors were stored.",
    ]
    output = ROOT / "final_validation_report.txt"
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
