"""
run_delivery.py — SignalStack one-shot delivery workflow.
=========================================================
Runs the full operator flow in sequence:
  1) export_to_csv.py
  2) run_pipeline.py
  3) generate_report.py
  4) package_output.py

This is intended for non-developer operators using Module Gateway.
"""

import argparse
import datetime
import subprocess
import sys
from pathlib import Path

from config import SOURCE_REGISTRY

BASE_DIR = Path(__file__).resolve().parent
SOURCE_OPTIONS = [*SOURCE_REGISTRY.keys(), "all"]


def run_step(name: str, command: list[str]) -> None:
    print(f"\n[delivery] STEP: {name}")
    print(f"[delivery] CMD: {' '.join(command)}")
    subprocess.run(command, cwd=str(BASE_DIR), check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="SignalStack full delivery workflow runner")
    parser.add_argument("--source", default="all", choices=SOURCE_OPTIONS)
    parser.add_argument("--skip-search", action="store_true")
    parser.add_argument("--skip-root-fix", action="store_true")
    parser.add_argument("--no-code", action="store_true")
    parser.add_argument("--out-dir", default=None, help="Output directory for packaged ZIP/HTML")
    parser.add_argument("--report-out", default=None, help="Optional explicit output path for DOCX report")
    args = parser.parse_args()

    source = args.source
    today = datetime.date.today()
    iso_week = today.strftime("%G-W%V")

    report_out = args.report_out
    if not report_out and args.out_dir:
        report_out = str((Path(args.out_dir) / f"SignalStack_Report_{iso_week}.docx").resolve())

    export_cmd = [sys.executable, "export_to_csv.py", "--source", source]
    if args.skip_root_fix:
        export_cmd.append("--skip-root-fix")
    run_step("Export source workbook(s) to CSV", export_cmd)

    pipeline_cmd = [sys.executable, "run_pipeline.py", "--source", source]
    if args.skip_search:
        pipeline_cmd.append("--skip-search")
    run_step("Run forecasting pipeline", pipeline_cmd)

    report_cmd = [sys.executable, "generate_report.py"]
    if source != "all":
        report_cmd.extend(["--source", source])
    if report_out:
        report_cmd.extend(["--out", report_out])
    run_step("Generate report", report_cmd)

    package_cmd = [sys.executable, "package_output.py"]
    if source != "all":
        package_cmd.extend(["--source", source])
    if args.no_code:
        package_cmd.append("--no-code")
    if args.out_dir:
        package_cmd.extend(["--out-dir", args.out_dir])
    run_step("Package delivery outputs", package_cmd)

    print("\n[delivery] COMPLETE: export -> pipeline -> report -> package")


if __name__ == "__main__":
    main()
