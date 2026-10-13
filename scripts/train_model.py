"""Train and report the justified same-record AQI estimation model."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_project import clean_data, load_data, model_results  # noqa: E402


if __name__ == "__main__":
    aq = clean_data(load_data())["air_quality"]
    results, _, note = model_results(aq)
    root = Path(__file__).resolve().parents[1]
    results.to_csv(root / "outputs" / "model_results.csv", index=False)
    print(results.to_string(index=False))
    print(note)
