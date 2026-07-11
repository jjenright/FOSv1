"""FOS command-line entry point for the current development release."""

from pathlib import Path

from src.extract import LayoutDetector


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    detector = LayoutDetector(project_root / "config" / "layouts.yaml")
    print("FOS data engine foundation is installed.")
    print(f"Configured annual worksheets: {len(detector.configured_sheets())}")


if __name__ == "__main__":
    main()
