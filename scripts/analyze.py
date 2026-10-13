"""Generate EDA figures and verified findings from the cleaned air-quality data."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_project import clean_data, findings, load_data, make_figures, model_results  # noqa: E402


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    aq = clean_data(load_data())["air_quality"]
    make_figures(aq)
    model, _, note = model_results(aq)
    text = findings(aq, model) + "\n" + note + "\n"
    (root / "outputs" / "eda_findings.txt").write_text(text, encoding="utf-8")
    print("Figures and outputs/eda_findings.txt regenerated.")
