"""Recalculate the workbook with LibreOffice and confirm it matches the Python model.

For each scenario, the script sets the scenario selector, has LibreOffice compute every
formula, then compares each LRP line with lrp.model.project(). Requires LibreOffice
(`soffice`) on the PATH.

    python scripts/check_excel.py [path/to/ServiceNow_LRP_Model.xlsx]
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from lrp.excel_export import LRP_LAYOUT  # noqa: E402
from lrp.model import HIST_YEARS, load_config, run_all  # noqa: E402

ERROR_STRINGS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def recalc(src: Path, scenario: str, workdir: Path) -> Path:
    wb = load_workbook(src)
    wb["Assumptions"]["C3"] = scenario.capitalize()
    staged = workdir / "in" / src.name
    staged.parent.mkdir(parents=True, exist_ok=True)
    wb.save(staged)
    outdir = workdir / "out"
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("LibreOffice (soffice) not found; open the file in Excel instead.")
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx",
                    "--outdir", str(outdir), str(staged)],
                   check=True, capture_output=True, timeout=180)
    return outdir / src.name


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "outputs" / "ServiceNow_LRP_Model.xlsx"
    cfg = load_config()
    results = run_all(cfg)
    years = HIST_YEARS + cfg.years
    keys = [item[0] for item in LRP_LAYOUT]
    rows, r = {}, 5
    for k in keys:
        if k not in ("section", "blank"):
            rows[k] = r
        r += 1

    failures = 0
    with tempfile.TemporaryDirectory() as tmp:
        for scen in cfg.scenarios:
            out = recalc(src, scen, Path(tmp) / scen)
            wb = load_workbook(out, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value in ERROR_STRINGS:
                            print(f"  formula error {ws.title}!{c.coordinate}: {c.value}")
                            failures += 1
            ws = wb["LRP"]
            df = results[scen]
            worst = 0.0
            for key, row_no in rows.items():
                for i, y in enumerate(years):
                    xl = ws.cell(row=row_no, column=3 + i).value
                    py = df.loc[key, y]
                    if xl is None or py != py:  # blank in Excel or NaN in Python
                        continue
                    diff = abs(float(xl) - float(py))
                    tol = 1e-6 if abs(py) < 5 else 1e-6 * abs(py) + 1e-6
                    worst = max(worst, diff)
                    if diff > tol:
                        failures += 1
                        print(f"  {scen}: LRP {key} FY{y} Excel={xl:.6f} Python={py:.6f}")
            print(f"{scen:>9}: LRP matches Python (largest absolute difference {worst:.2e})"
                  if failures == 0 else f"{scen:>9}: mismatches found")
    print("PASS" if failures == 0 else f"FAIL ({failures} issues)")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
