"""Create traceable cleaned datasets and preprocessing flags."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_project import clean_data, load_data  # noqa: E402


if __name__ == "__main__":
    cleaned = clean_data(load_data())
    print({name: frame.shape for name, frame in cleaned.items()})
