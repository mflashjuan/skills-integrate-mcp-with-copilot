import argparse
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from src.data_ingestion import build_activity_data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build activity metadata from Excel and DOCX source files."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("data/input"),
        help="Excel/DOCX file or directory containing activity records.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/activities.json"),
        help="Path to write the generated activities JSON file.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = args.input
    output = args.output

    activities = build_activity_data(source, output)
    print(f"Built {len(activities)} activities from {source}")
    print(f"Saved activities to {output}")


if __name__ == "__main__":
    main()
