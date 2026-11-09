"""Write the LRP to an Excel workbook with live formulas.

The Assumptions tab holds every driver for all three scenarios plus a scenario selector;
the LRP tab is pure formulas off those drivers and the reported historicals. Change a
driver (blue cell) or the selected scenario (yellow cell) and the plan recalculates.
Scenario comparison and sensitivity tabs are snapshots written by run_model.py.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .model import HIST_YEARS, Config

FONT = "Arial"
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "666666"
NAVY = "1F3864"
F_INPUT = Font(name=FONT, size=10, color=BLUE)
F_CALC = Font(name=FONT, size=10, color=BLACK)
F_LINK = Font(name=FONT, size=10, color=GREEN)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color=GREY)
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_FCST = PatternFill("solid", fgColor="F2F2F2")
THIN_TOP = Border(top=Side(style="thin", color="000000"))

NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.00x'

SOURCE_Q3 = ("ServiceNow Q3 2026 earnings release, Form 8-K filed Oct 2026: "
             "https://www.sec.gov/Archives/edgar/data/0001373715/000137371526000092/erq3fy26.htm")


def _font(cell, font):
    cell.font = font


def _header_row(ws, row, values, start_col=2):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center" if i else "left")


def _set_widths(ws, widths: dict[str, float]):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# --------------------------------------------------------------------------- Assumptions
def _assumptions_sheet(wb: Workbook, cfg: Config) -> dict[str, int]:
    ws = wb.create_sheet("Assumptions")
    years = cfg.years
    drivers = list(cfg.raw["drivers"].keys())
    meta = cfg.raw["drivers"]
    ws["B1"] = "Assumptions & scenario drivers"
    ws["B1"].font = F_TITLE

    ws["B3"], ws["B4"] = "Selected scenario", "Scenario index (calc)"
    ws["B3"].font = ws["B4"].font = F_BOLD
    ws["C3"] = "Base"
    ws["C3"].font, ws["C3"].fill = F_INPUT, FILL_KEY
    ws["C3"].comment = Comment("Pick Base, Upside or Downside. The LRP tab recalculates.",
                               "Model")
    names = [s.capitalize() for s in cfg.scenarios]
    for i, n in enumerate(names):
        c = ws.cell(row=3, column=10 + i, value=n)
        c.font = F_NOTE
    ws["I3"] = "Scenario list:"
    ws["I3"].font = F_NOTE
    ws["C4"] = f"=MATCH(C3,$J$3:${get_column_letter(9 + len(names))}$3,0)"
    ws["C4"].font = F_CALC
    dv = DataValidation(type="list", formula1='"' + ",".join(names) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C3")

    block_len = len(drivers)
    live_start = 8
    starts = {"live": live_start}
    row = live_start + block_len + 3
    for s in cfg.scenarios:
        starts[s] = row + 1
        row += block_len + 4

    year_hdr = [f"FY{y}E" for y in years]

    # Live block (selected scenario)
    ws.cell(row=live_start - 1, column=2)
    _header_row(ws, live_start - 1, ["LIVE DRIVERS (selected scenario) - feeds LRP tab"]
                + year_hdr + ["Notes"])
    for j, d in enumerate(drivers):
        r = live_start + j
        ws.cell(row=r, column=2, value=meta[d]["label"]).font = F_CALC
        for k in range(len(years)):
            col = get_column_letter(3 + k)
            refs = ",".join(f"{col}{starts[s] + j}" for s in cfg.scenarios)
            c = ws.cell(row=r, column=3 + k, value=f"=CHOOSE($C$4,{refs})")
            c.font = F_CALC
            c.number_format = MULT if meta[d]["fmt"] == "x" else PCT
    note = ws.cell(row=live_start + block_len, column=2,
                   value="FY2026E is anchored to company guidance and identical across "
                         "scenarios (three quarters already reported). " + SOURCE_Q3)
    note.font = F_NOTE

    # Scenario blocks (inputs)
    for s in cfg.scenarios:
        st = starts[s]
        desc = " ".join(cfg.raw["scenarios"][s]["description"].split())
        _header_row(ws, st - 1, [f"{s.upper()} CASE - inputs"] + year_hdr + ["Notes"])
        for j, d in enumerate(drivers):
            r = st + j
            ws.cell(row=r, column=2, value=meta[d]["label"]).font = F_CALC
            for k, v in enumerate(cfg.drivers(s)[d]):
                c = ws.cell(row=r, column=3 + k, value=float(v))
                c.font = F_INPUT
                c.number_format = MULT if meta[d]["fmt"] == "x" else PCT
        c = ws.cell(row=st + block_len, column=2, value=f"Thesis: {desc}")
        c.font = F_NOTE

    notes = {
        "sub_growth": "FY26 = guidance midpoint $15,785M / FY25 $12,883M - 1. Q4'26 cRPO guide +18.5% sets the FY27 starting point.",
        "ps_growth": "9M FY26 professional services +12% y/y; a low-margin, shrinking share of revenue.",
        "sub_gm": "FY26 guide 75% GAAP (81% non-GAAP); gap = acquisition amortization (Armis, Veza), assumed to fade.",
        "ps_gm": "GAAP services margin was -5% in FY25; services run near breakeven by design.",
        "sm_pct": "FY25 33.0%, 9M FY26 32.5%. Main lever for margin expansion.",
        "rd_pct": "FY25 22.3%, 9M FY26 22.7% (includes acquired R&D teams).",
        "ga_pct": "FY25 8.5%, 9M FY26 7.9%.",
        "sbc_pct": "FY25 14.7%, 9M FY26 15.6% of revenue.",
        "other_adj_pct": "Calibrated so FY26 GAAP margin = 10% and non-GAAP = 31.5% (guidance).",
        "cash_tax_rate": "Assumption; rises as tax attributes are used up.",
        "wc_capture": "Annual upfront billing makes deferred revenue grow with revenue; calibrated to FY26 36% FCF margin guide.",
        "net_interest_pct": "Interest on cash & investments net of debt cost (Armis term loan). Assumption.",
        "capex_pct": "9M FY26 2.9% of revenue; assumed higher for AI data center capacity.",
        "rev_per_emp_growth": "FY25 revenue per employee $455K. AI-enabled productivity assumption.",
        "crpo_ratio": "FY25 cRPO $12.85B / subscription revenue $12.88B = 1.00x.",
    }
    ncol = 3 + len(years)
    for j, d in enumerate(drivers):
        c = ws.cell(row=live_start + j, column=ncol, value=notes.get(d, ""))
        c.font = F_NOTE

    _set_widths(ws, {"A": 2, "B": 50, **{get_column_letter(3 + k): 11 for k in range(len(years))},
                     get_column_letter(ncol): 95})
    ws.freeze_panes = "C8"
    return {d: live_start + j for j, d in enumerate(drivers)}


# --------------------------------------------------------------------------- Historicals
def _historicals_sheet(wb: Workbook, csv_path: Path) -> dict[str, int]:
    ws = wb.create_sheet("Historicals")
    df = pd.read_csv(csv_path)
    ws["B1"] = "Reported historicals (as filed)"
    ws["B1"].font = F_TITLE
    ws["B2"] = "$ in millions unless noted. Blue = hardcoded from filings; see source column."
    ws["B2"].font = F_NOTE
    _header_row(ws, 4, ["Metric", "Unit", "FY2023A", "FY2024A", "FY2025A", "Source"])
    rows = {}
    for i, rec in enumerate(df.itertuples(index=False)):
        r = 5 + i
        rows[rec.metric] = r
        ws.cell(row=r, column=2, value=rec.label).font = F_CALC
        ws.cell(row=r, column=3, value=rec.unit).font = F_CALC
        for k, y in enumerate(HIST_YEARS):
            v = getattr(rec, f"FY{y}")
            c = ws.cell(row=r, column=4 + k, value=None if pd.isna(v) else float(v))
            c.font = F_INPUT
            c.number_format = PCT if rec.unit == "%" else NUM
        ws.cell(row=r, column=7, value=rec.source).font = F_NOTE
    _set_widths(ws, {"A": 2, "B": 48, "C": 7, "D": 11, "E": 11, "F": 11, "G": 110})
    return rows


# --------------------------------------------------------------------------- LRP
LRP_LAYOUT = [
    ("section", "Income statement ($M)"),
    ("sub_rev", "Subscription revenue", NUM),
    ("ps_rev", "Professional services & other revenue", NUM),
    ("total_rev", "Total revenue", NUM),
    ("sub_growth", "  Subscription revenue growth", PCT),
    ("blank",),
    ("sub_cogs", "Cost of subscription revenue", NUM),
    ("ps_cogs", "Cost of professional services revenue", NUM),
    ("gross_profit", "Gross profit (GAAP)", NUM),
    ("gross_margin", "  Gross margin", PCT),
    ("blank",),
    ("sm", "Sales & marketing", NUM),
    ("rd", "Research & development", NUM),
    ("ga", "General & administrative", NUM),
    ("total_opex", "Total operating expenses", NUM),
    ("gaap_oi", "Operating income (GAAP)", NUM),
    ("gaap_om", "  GAAP operating margin", PCT),
    ("blank",),
    ("sbc", "Add: stock-based compensation", NUM),
    ("other_adj", "Add: amortization & acquisition-related costs", NUM),
    ("ng_oi", "Operating income (non-GAAP)", NUM),
    ("ng_om", "  Non-GAAP operating margin", PCT),
    ("section", "Cash flow ($M)"),
    ("cash_taxes", "Less: cash taxes", NUM),
    ("wc_inflow", "Add: working capital inflow (deferred revenue)", NUM),
    ("net_interest", "Add: net interest & other cash income", NUM),
    ("ocf", "Operating cash flow", NUM),
    ("capex", "Less: capital expenditures (shown positive)", NUM),
    ("fcf", "Free cash flow", NUM),
    ("fcf_margin", "  FCF margin", PCT),
    ("section", "SaaS & operating metrics"),
    ("headcount", "Headcount (year end)", NUM),
    ("rev_per_emp", "Revenue per employee ($K)", NUM),
    ("crpo", "cRPO, year end ($M)", NUM),
    ("rule_of_40", "Rule of 40 (subscription growth + FCF margin)", PCT),
]
BOLD_LINES = {"total_rev", "gross_profit", "gaap_oi", "ng_oi", "ocf", "fcf"}


def _lrp_sheet(wb: Workbook, cfg: Config, drv_rows: dict[str, int],
               hist_rows: dict[str, int]) -> dict[str, int]:
    ws = wb.create_sheet("LRP")
    years = cfg.years
    all_years = HIST_YEARS + years
    ws["B1"] = "ServiceNow integrated long-range plan"
    ws["B1"].font = F_TITLE
    ws["B2"] = '="Scenario shown: "&Assumptions!C3&"   |   $ in millions unless noted"'
    ws["B2"].font = F_BOLD
    _header_row(ws, 4, ["Line item"] + [f"FY{y}A" for y in HIST_YEARS]
                + [f"FY{y}E" for y in years])

    # Row map
    rows, r = {}, 5
    for item in LRP_LAYOUT:
        if item[0] not in ("section", "blank"):
            rows[item[0]] = r
        r += 1

    col = {y: get_column_letter(3 + i) for i, y in enumerate(all_years)}
    acol = {y: get_column_letter(3 + i) for i, y in enumerate(years)}  # Assumptions cols

    def A(driver, y):
        return f"Assumptions!{acol[y]}${drv_rows[driver]}"

    def H(metric, y):
        hc = get_column_letter(4 + HIST_YEARS.index(y))
        return f"Historicals!{hc}{hist_rows[metric]}"

    def R(line, y):
        return f"{col[y]}{rows[line]}"

    r = 5
    for item in LRP_LAYOUT:
        kind = item[0]
        if kind == "section":
            c = ws.cell(row=r, column=2, value=item[1])
            c.font, c.fill = F_BOLD, FILL_SECTION
            for k in range(len(all_years)):
                ws.cell(row=r, column=3 + k).fill = FILL_SECTION
            r += 1
            continue
        if kind == "blank":
            r += 1
            continue
        key, label, fmt = item
        ws.cell(row=r, column=2, value=label).font = F_BOLD if key in BOLD_LINES else F_CALC
        for y in all_years:
            f, font = _formula(key, y, R, A, H, all_years), F_CALC
            if f is None:
                continue
            if y in HIST_YEARS and f.startswith("=Historicals!"):
                font = F_LINK
            elif y not in HIST_YEARS and f.startswith("=Assumptions!"):
                font = F_LINK
            c = ws.cell(row=r, column=3 + all_years.index(y), value=f)
            c.font = Font(name=FONT, size=10, color=font.color,
                          bold=key in BOLD_LINES)
            c.number_format = fmt
            if y not in HIST_YEARS:
                c.fill = FILL_FCST
            if key in BOLD_LINES:
                c.border = THIN_TOP
        r += 1

    note_r = r + 1
    notes = [
        "Colour code: blue = hardcoded input, black = formula, green = link to another tab. "
        "Grey shading = forecast years.",
        "Historical non-GAAP operating income = reported (rounded) non-GAAP margin x revenue; "
        "the amortization/acquisition line is backed out from it, so it is approximate.",
        "Historical free cash flow is as reported by ServiceNow, which differs slightly from "
        "operating cash flow less capex (the company excludes some items).",
        "Historical cash taxes, working capital and interest are not broken out; the forecast "
        "builds operating cash flow from non-GAAP operating income.",
        "FY2024 cRPO left blank: not stated in the releases used.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=note_r + i, column=2, value=n).font = F_NOTE

    _set_widths(ws, {"A": 2, "B": 48, **{get_column_letter(3 + i): 11
                                          for i in range(len(all_years))}})
    ws.freeze_panes = "C5"
    return rows


def _formula(key, y, R, A, H, all_years):
    """Excel formula for a line in a year; mirrors lrp.model.project."""
    hist = y in HIST_YEARS
    first = y == all_years[0]
    p = all_years[all_years.index(y) - 1] if not first else None
    direct = {"sub_rev", "ps_rev", "sub_cogs", "ps_cogs", "sm", "rd", "ga", "sbc",
              "ocf", "capex", "fcf", "headcount", "crpo"}

    if hist:
        if key in direct:
            if key == "crpo" and y == 2024:
                return None
            return f"={H(key, y)}"
        return {
            "total_rev": f"={R('sub_rev', y)}+{R('ps_rev', y)}",
            "sub_growth": None if first else f"={R('sub_rev', y)}/{R('sub_rev', p)}-1",
            "gross_profit": f"={R('total_rev', y)}-{R('sub_cogs', y)}-{R('ps_cogs', y)}",
            "gross_margin": f"=IF({R('total_rev', y)}=0,0,{R('gross_profit', y)}/{R('total_rev', y)})",
            "total_opex": f"={R('sm', y)}+{R('rd', y)}+{R('ga', y)}",
            "gaap_oi": f"={R('gross_profit', y)}-{R('total_opex', y)}",
            "gaap_om": f"=IF({R('total_rev', y)}=0,0,{R('gaap_oi', y)}/{R('total_rev', y)})",
            "ng_oi": f"={H('ng_om_reported', y)}*{R('total_rev', y)}",
            "other_adj": f"={R('ng_oi', y)}-{R('gaap_oi', y)}-{R('sbc', y)}",
            "ng_om": f"=IF({R('total_rev', y)}=0,0,{R('ng_oi', y)}/{R('total_rev', y)})",
            "fcf_margin": f"=IF({R('total_rev', y)}=0,0,{R('fcf', y)}/{R('total_rev', y)})",
            "rev_per_emp": f"={R('total_rev', y)}*1000/{R('headcount', y)}",
            "rule_of_40": None if first else f"={R('sub_growth', y)}+{R('fcf_margin', y)}",
        }.get(key)

    tr = R("total_rev", y)
    return {
        "sub_rev": f"={R('sub_rev', p)}*(1+{A('sub_growth', y)})",
        "ps_rev": f"={R('ps_rev', p)}*(1+{A('ps_growth', y)})",
        "total_rev": f"={R('sub_rev', y)}+{R('ps_rev', y)}",
        "sub_growth": f"={R('sub_rev', y)}/{R('sub_rev', p)}-1",
        "sub_cogs": f"={R('sub_rev', y)}*(1-{A('sub_gm', y)})",
        "ps_cogs": f"={R('ps_rev', y)}*(1-{A('ps_gm', y)})",
        "gross_profit": f"={tr}-{R('sub_cogs', y)}-{R('ps_cogs', y)}",
        "gross_margin": f"=IF({tr}=0,0,{R('gross_profit', y)}/{tr})",
        "sm": f"={tr}*{A('sm_pct', y)}",
        "rd": f"={tr}*{A('rd_pct', y)}",
        "ga": f"={tr}*{A('ga_pct', y)}",
        "total_opex": f"={R('sm', y)}+{R('rd', y)}+{R('ga', y)}",
        "gaap_oi": f"={R('gross_profit', y)}-{R('total_opex', y)}",
        "gaap_om": f"=IF({tr}=0,0,{R('gaap_oi', y)}/{tr})",
        "sbc": f"={tr}*{A('sbc_pct', y)}",
        "other_adj": f"={tr}*{A('other_adj_pct', y)}",
        "ng_oi": f"={R('gaap_oi', y)}+{R('sbc', y)}+{R('other_adj', y)}",
        "ng_om": f"=IF({tr}=0,0,{R('ng_oi', y)}/{tr})",
        "cash_taxes": f"=-{R('ng_oi', y)}*{A('cash_tax_rate', y)}",
        "wc_inflow": f"=({tr}-{R('total_rev', p)})*{A('wc_capture', y)}",
        "net_interest": f"={tr}*{A('net_interest_pct', y)}",
        "ocf": f"={R('ng_oi', y)}+{R('cash_taxes', y)}+{R('wc_inflow', y)}+{R('net_interest', y)}",
        "capex": f"={tr}*{A('capex_pct', y)}",
        "fcf": f"={R('ocf', y)}-{R('capex', y)}",
        "fcf_margin": f"=IF({tr}=0,0,{R('fcf', y)}/{tr})",
        "rev_per_emp": f"={R('rev_per_emp', p)}*(1+{A('rev_per_emp_growth', y)})",
        "headcount": f"={tr}*1000/{R('rev_per_emp', y)}",
        "crpo": f"={R('sub_rev', y)}*{A('crpo_ratio', y)}",
        "rule_of_40": f"={R('sub_growth', y)}+{R('fcf_margin', y)}",
    }[key]


# --------------------------------------------------------------------------- Guidance
def _guidance_sheet(wb: Workbook, cfg: Config, lrp_rows: dict[str, int]):
    ws = wb.create_sheet("Guidance Check")
    gd = cfg.raw["guidance_2026"]
    ws["B1"] = "FY2026: model vs. company guidance"
    ws["B1"].font = F_TITLE
    ws["B2"] = "Ties the long-range plan to the current-year operating plan. " + SOURCE_Q3
    ws["B2"].font = F_NOTE
    _header_row(ws, 4, ["Metric", "Model FY2026E", "Guidance", "Difference"])
    c26 = "F"  # FY2026 column on LRP
    L = lambda k: f"LRP!{c26}{lrp_rows[k]}"  # noqa: E731
    items = [
        ("Subscription revenue ($M)", f"={L('sub_rev')}",
         (gd["sub_rev_low"] + gd["sub_rev_high"]) / 2, NUM),
        ("Subscription gross margin (GAAP)",
         f"=IF({L('sub_rev')}=0,0,1-{L('sub_cogs')}/{L('sub_rev')})", gd["sub_gm_gaap"], PCT),
        ("GAAP operating margin", f"={L('gaap_om')}", gd["gaap_om"], PCT),
        ("Non-GAAP operating margin", f"={L('ng_om')}", gd["ng_om"], PCT),
        ("Free cash flow margin", f"={L('fcf_margin')}", gd["fcf_margin"], PCT),
    ]
    for i, (label, f, g, fmt) in enumerate(items):
        r = 5 + i
        ws.cell(row=r, column=2, value=label).font = F_CALC
        c = ws.cell(row=r, column=3, value=f)
        c.font, c.number_format = F_LINK, fmt
        c = ws.cell(row=r, column=4, value=float(g))
        c.font, c.number_format = F_INPUT, fmt
        c = ws.cell(row=r, column=5, value=f"=C{r}-D{r}")
        c.font, c.number_format = F_CALC, fmt
    ws.cell(row=11, column=2, value="Subscription revenue guidance is the midpoint of "
            "$15,780-15,790M. Margins are company guidance as of the Q3 2026 release.").font = F_NOTE
    _set_widths(ws, {"A": 2, "B": 38, "C": 16, "D": 14, "E": 14})


# --------------------------------------------------------------------------- Snapshots
def _scenarios_sheet(wb: Workbook, results: dict[str, pd.DataFrame], cfg: Config,
                     charts: dict[str, Path]):
    ws = wb.create_sheet("Scenarios")
    ws["B1"] = "Scenario comparison (snapshot)"
    ws["B1"].font = F_TITLE
    ws["B2"] = ("Values written by run_model.py from config/assumptions.yaml. The LRP tab "
                "shows one scenario live; this tab shows all three side by side.")
    ws["B2"].font = F_NOTE
    years = [2025] + cfg.years
    lines = [("total_rev", "Total revenue ($M)", NUM), ("sub_growth", "Subscription growth", PCT),
             ("ng_om", "Non-GAAP operating margin", PCT), ("fcf", "Free cash flow ($M)", NUM),
             ("fcf_margin", "FCF margin", PCT), ("headcount", "Headcount", NUM),
             ("rule_of_40", "Rule of 40", PCT)]
    r = 4
    for s, df in results.items():
        _header_row(ws, r, [f"{s.capitalize()} case"] +
                    [f"FY{y}{'A' if y <= 2025 else 'E'}" for y in years])
        for key, label, fmt in lines:
            r += 1
            ws.cell(row=r, column=2, value=label).font = F_CALC
            for k, y in enumerate(years):
                c = ws.cell(row=r, column=3 + k, value=float(df.loc[key, y]))
                c.font, c.number_format = F_CALC, fmt
        r += 1
        desc = " ".join(cfg.raw["scenarios"][s]["description"].split())
        ws.cell(row=r, column=2, value=desc).font = F_NOTE
        r += 2
    _set_widths(ws, {"A": 2, "B": 30, **{get_column_letter(3 + k): 11 for k in range(len(years))}})
    anchor_row = 2
    for i, key in enumerate(["revenue", "fcf_margin", "headcount"]):
        if key in charts and charts[key].exists():
            img = XLImage(str(charts[key]))
            img.width, img.height = 600, 336
            ws.add_image(img, f"K{anchor_row + i * 18}")


def _sensitivity_sheet(wb: Workbook, tornado_df: pd.DataFrame, grid: pd.DataFrame,
                       cfg: Config, chart: Path | None):
    ws = wb.create_sheet("Sensitivity")
    year = cfg.raw["sensitivity"]["target_year"]
    ws["B1"] = f"Sensitivity of FY{year} free cash flow - base case (snapshot)"
    ws["B1"].font = F_TITLE
    ws["B2"] = ("Each driver shifted up and down in FY2027-30 (FY2026 stays on guidance), "
                "one at a time. Written by run_model.py.")
    ws["B2"].font = F_NOTE
    _header_row(ws, 4, ["Driver", "Shock (+/-)", "Low case ($M)", "High case ($M)",
                        "Base ($M)", "Swing ($M)"])
    for i, rec in enumerate(tornado_df.itertuples()):
        r = 5 + i
        ws.cell(row=r, column=2, value=rec.label).font = F_CALC
        vals = [(rec.shock, PCT), (rec.low_case, NUM), (rec.high_case, NUM),
                (rec.base, NUM), (rec.swing, NUM)]
        for k, (v, fmt) in enumerate(vals):
            c = ws.cell(row=r, column=3 + k, value=float(v))
            c.font, c.number_format = F_CALC, fmt

    r0 = 6 + len(tornado_df) + 2
    rd, cd = cfg.raw["sensitivity"]["two_way"]["rows"], cfg.raw["sensitivity"]["two_way"]["cols"]
    labels = {k: v["label"] for k, v in cfg.raw["drivers"].items()}
    ws.cell(row=r0 - 1, column=2,
            value=f"Two-way: FY{year} FCF ($M). Rows = shift in {labels[rd['driver']]}; "
                  f"columns = shift in {labels[cd['driver']]}.").font = F_BOLD
    _header_row(ws, r0, ["Growth shift \\ S&M shift"] +
                [f"{v * 100:+.0f} pts" for v in grid.columns])
    for i, idx in enumerate(grid.index):
        r = r0 + 1 + i
        ws.cell(row=r, column=2, value=f"{idx * 100:+.0f} pts").font = F_BOLD
        for k, col in enumerate(grid.columns):
            c = ws.cell(row=r, column=3 + k, value=float(grid.loc[idx, col]))
            c.font, c.number_format = F_CALC, NUM
            if idx == 0 and col == 0:
                c.fill = FILL_KEY
    _set_widths(ws, {"A": 2, "B": 50, "C": 13, "D": 14, "E": 14, "F": 13, "G": 13})
    if chart and chart.exists():
        img = XLImage(str(chart))
        img.width, img.height = 600, 336
        ws.add_image(img, "I4")


# --------------------------------------------------------------------------- Investment
def _investment_sheet(wb: Workbook, cfg: Config):
    ws = wb.create_sheet("Investment Case")
    ic = cfg.raw["investment_case"]
    years = ic["years"]
    n = len(years)
    ws["B1"] = "Investment proposal evaluation"
    ws["B1"].font = F_TITLE
    ws["B2"] = ic["name"]
    ws["B2"].font = F_BOLD
    ws["B3"] = " ".join(ic["description"].split())
    ws["B3"].font = F_NOTE

    ws["B5"] = "Inputs"
    ws["B5"].font = F_BOLD
    inputs = [("Discount rate (WACC)", ic["discount_rate"], PCT),
              ("Terminal growth", ic["terminal_growth"], PCT),
              ("Cash tax rate", ic["cash_tax_rate"], PCT),
              ("Contribution margin on new revenue", ic["contribution_margin"], PCT),
              ("Payback hurdle (years)", ic["hurdle_payback_years"], "0.0")]
    for i, (label, v, fmt) in enumerate(inputs):
        r = 6 + i
        ws.cell(row=r, column=2, value=label).font = F_CALC
        c = ws.cell(row=r, column=3, value=float(v))
        c.font, c.number_format, c.fill = F_INPUT, fmt, FILL_KEY
    WACC, TG, TAX, CM, HURDLE = "$C$6", "$C$7", "$C$8", "$C$9", "$C$10"

    hr = 12
    _header_row(ws, hr, ["Cash flows ($M)"] + [f"FY{y}" for y in years])
    cols = [get_column_letter(3 + k) for k in range(n)]
    labels = ["Year number", "Incremental subscription revenue", "Contribution (rev x margin)",
              "Run-rate cost of added team", "Upfront investment", "Pre-tax cash flow",
              "Cash taxes", "Free cash flow", "Cumulative free cash flow", "Discount factor",
              "Present value", "Payback helper"]
    R = {lab: hr + 1 + i for i, lab in enumerate(labels)}
    for lab, r in R.items():
        ws.cell(row=r, column=2, value=lab).font = F_NOTE if lab == "Payback helper" else (
            F_BOLD if lab == "Free cash flow" else F_CALC)
    for k, c in enumerate(cols):
        def put(lab, v, fmt=NUM, font=F_CALC):
            cell = ws[f"{c}{R[lab]}"]
            cell.value, cell.number_format, cell.font = v, fmt, font
        put("Year number", k + 1, "0", F_INPUT)
        put("Incremental subscription revenue", float(ic["incremental_revenue"][k]), NUM, F_INPUT)
        put("Contribution (rev x margin)", f"={c}{R['Incremental subscription revenue']}*{CM}")
        put("Run-rate cost of added team", float(ic["run_rate_cost"][k]), NUM, F_INPUT)
        put("Upfront investment", float(ic["upfront_investment"][k]), NUM, F_INPUT)
        put("Pre-tax cash flow", f"={c}{R['Contribution (rev x margin)']}-"
            f"{c}{R['Run-rate cost of added team']}-{c}{R['Upfront investment']}")
        put("Cash taxes", f"=-MAX({c}{R['Pre-tax cash flow']},0)*{TAX}")
        put("Free cash flow", f"={c}{R['Pre-tax cash flow']}+{c}{R['Cash taxes']}",
            NUM, F_BOLD)
        prev = f"+{cols[k - 1]}{R['Cumulative free cash flow']}" if k else ""
        put("Cumulative free cash flow", f"={c}{R['Free cash flow']}{prev}")
        put("Discount factor", f"=1/(1+{WACC})^{c}{R['Year number']}", "0.000")
        put("Present value", f"={c}{R['Free cash flow']}*{c}{R['Discount factor']}")
        cum, fcf = R["Cumulative free cash flow"], R["Free cash flow"]
        if k == 0:
            helper = f'=IF({c}{cum}>=0,0,"")'
        else:
            pc = cols[k - 1]
            helper = (f'=IF(AND({c}{cum}>=0,{pc}{cum}<0),'
                      f'{pc}{R["Year number"]}+(-{pc}{cum}/{c}{fcf}),"")')
        put("Payback helper", helper, "0.00", F_NOTE)
    ws.cell(row=R["Run-rate cost of added team"], column=3 + n,
            value="Inputs entered as positive costs").font = F_NOTE

    first, last = cols[0], cols[-1]
    fr = R["Free cash flow"]
    o = R["Payback helper"] + 2
    ws.cell(row=o, column=2, value="Results").font = F_BOLD
    out = [
        ("PV of explicit cash flows ($M)", f"=NPV({WACC},{first}{fr}:{last}{fr})", NUM),
        ("Terminal value at end of FY%d ($M)" % years[-1],
         f"={last}{fr}*(1+{TG})/({WACC}-{TG})", NUM),
        ("PV of terminal value ($M)", f"=C{o + 2}*{last}{R['Discount factor']}", NUM),
        ("NPV including terminal value ($M)", f"=C{o + 1}+C{o + 3}", NUM),
        ("Share of NPV from terminal value", f"=IF(C{o + 4}=0,0,C{o + 3}/C{o + 4})", PCT),
        ("IRR on explicit cash flows (no terminal)", f"=IRR({first}{fr}:{last}{fr})", PCT),
        ("Payback (years)", f'=IF(COUNT({first}{R["Payback helper"]}:{last}{R["Payback helper"]})=0,'
                            f'"Beyond FY{years[-1]}",MIN({first}{R["Payback helper"]}:{last}{R["Payback helper"]}))', "0.0"),
        ("Decision rule: NPV > 0 and payback within hurdle",
         f'=IF(AND(C{o + 4}>0,ISNUMBER(C{o + 7}),C{o + 7}<={HURDLE}),"Approve",'
         f'"Revisit: restructure or stage-gate")', None),
    ]
    for i, (label, f, fmt) in enumerate(out):
        r = o + 1 + i
        ws.cell(row=r, column=2, value=label).font = F_CALC
        c = ws.cell(row=r, column=3, value=f)
        c.font = F_BOLD if i in (3, 7) else F_CALC
        if fmt:
            c.number_format = fmt
    ws.cell(row=o + 10, column=2, value=(
        "Read-across: when most of the value sits in the terminal value and payback misses the "
        "hurdle, fund tranche 1 and release tranche 2 only if AI attach rates hit target.")).font = F_NOTE
    _set_widths(ws, {"A": 2, "B": 46, **{c: 12 for c in cols}, get_column_letter(3 + n): 30})


# --------------------------------------------------------------------------- Cover
def _cover_sheet(wb: Workbook, cfg: Config, lrp_rows: dict[str, int]):
    ws = wb.active
    ws.title = "Cover"
    ws["B2"] = "ServiceNow (NOW) - Long-Range Plan Model"
    ws["B2"].font = Font(name=FONT, size=18, bold=True, color=NAVY)
    ws["B3"] = ("Driver-based five-year plan (FY2026-FY2030) built from public filings: revenue, "
                "margins, opex, headcount, cash flow and SaaS metrics, with three scenarios, "
                "sensitivities and an investment-proposal evaluator.")
    ws["B3"].font = F_NOTE
    ws["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B3:H4")

    ws["B6"] = "How to use"
    ws["B6"].font = F_BOLD
    steps = [
        "1. Assumptions tab: choose Base / Upside / Downside in the yellow cell (C3).",
        "2. Edit any blue driver; the LRP, Guidance Check and this cover recalculate.",
        "3. Scenarios and Sensitivity tabs are snapshots; re-run `python run_model.py` after editing the YAML.",
        "4. Investment Case tab: change the yellow inputs or blue cash flows to test a proposal.",
    ]
    for i, s in enumerate(steps):
        ws.cell(row=7 + i, column=2, value=s).font = F_CALC

    ws["B12"] = "Legend"
    ws["B12"].font = F_BOLD
    legend = [("Blue text", "Hardcoded input or reported figure", F_INPUT, None),
              ("Black text", "Formula", F_CALC, None),
              ("Green text", "Link to another tab", F_LINK, None),
              ("Yellow fill", "Key assumption / selector", F_CALC, FILL_KEY)]
    for i, (a, b, f, fill) in enumerate(legend):
        c = ws.cell(row=13 + i, column=2, value=a)
        c.font = f
        if fill:
            c.fill = fill
        ws.cell(row=13 + i, column=3, value=b).font = F_CALC

    ws["B18"] = '="Key outputs - "&Assumptions!C3&" case ($M)"'
    ws["B18"].font = F_BOLD
    _header_row(ws, 19, ["Metric", "FY2025A", "FY2026E", "FY2030E", "CAGR / change"])
    L = lambda k, c: f"LRP!{c}{lrp_rows[k]}"  # noqa: E731
    kpis = [("Total revenue", "total_rev", NUM, True),
            ("Non-GAAP operating margin", "ng_om", PCT, False),
            ("Free cash flow", "fcf", NUM, True),
            ("FCF margin", "fcf_margin", PCT, False),
            ("Headcount", "headcount", NUM, True),
            ("Rule of 40", "rule_of_40", PCT, False)]
    for i, (label, k, fmt, cagr) in enumerate(kpis):
        r = 20 + i
        ws.cell(row=r, column=2, value=label).font = F_CALC
        for j, c in enumerate(["E", "F", "J"]):
            cell = ws.cell(row=r, column=3 + j, value=f"={L(k, c)}")
            cell.font, cell.number_format = F_LINK, fmt
        if cagr:
            f = f"=IF(C{r}<=0,0,(E{r}/C{r})^(1/5)-1)"
        else:
            f = f"=E{r}-C{r}"
        cell = ws.cell(row=r, column=6, value=f)
        cell.font, cell.number_format = F_CALC, PCT
    ws.cell(row=26, column=2, value="CAGR for $ and headcount lines (FY25-30); "
            "percentage-point change for ratios.").font = F_NOTE
    ws.cell(row=28, column=2, value="Independent analysis from public information; not "
            "affiliated with or endorsed by ServiceNow. Forecasts are illustrative.").font = F_NOTE
    _set_widths(ws, {"A": 2, "B": 34, "C": 14, "D": 14, "E": 14, "F": 16, "G": 12, "H": 12})


def _sources_sheet(wb: Workbook):
    ws = wb.create_sheet("Sources")
    ws["B1"] = "Sources"
    ws["B1"].font = F_TITLE
    src = [
        ("FY2024-FY2025 financials, FY2026 initial guidance",
         "ServiceNow Q4 & FY2025 earnings release (Jan 28, 2026)",
         "https://s205.q4cdn.com/916135447/files/doc_financials/2025/q4/ServiceNow-Reports-Fourth-Quarter-and-Full-Year-2025-Financial-Results.pdf"),
        ("FY2023 financials", "ServiceNow Q4 & FY2023 earnings release (Jan 24, 2024)",
         "https://www.businesswire.com/news/home/20240124662293/en/ServiceNow-Reports-Fourth-Quarter-and-Full-Year-2023-Financial-Results"),
        ("Q3 2026 results, 9M FY2026 run-rates, updated FY2026 guidance",
         "ServiceNow Q3 2026 earnings release, Form 8-K",
         "https://www.sec.gov/Archives/edgar/data/0001373715/000137371526000092/erq3fy26.htm"),
        ("Headcount (year end)", "Form 10-K employee counts, compiled by stockanalysis.com",
         "https://stockanalysis.com/stocks/now/employees/"),
        ("Renewal rate (98%), customer counts", "ServiceNow Q4 2025 fact sheet",
         "https://s205.q4cdn.com/916135447/files/doc_downloads/fact-sheet/q4-fact-sheet.pdf"),
    ]
    _header_row(ws, 3, ["Used for", "Document", "Link"])
    for i, (a, b, c) in enumerate(src):
        ws.cell(row=4 + i, column=2, value=a).font = F_CALC
        ws.cell(row=4 + i, column=3, value=b).font = F_CALC
        cell = ws.cell(row=4 + i, column=4, value=c)
        cell.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        cell.hyperlink = c
    _set_widths(ws, {"A": 2, "B": 52, "C": 55, "D": 120})


# --------------------------------------------------------------------------- Entry point
def write_workbook(path: Path, cfg: Config, results: dict[str, pd.DataFrame],
                   tornado_df: pd.DataFrame, grid: pd.DataFrame,
                   charts: dict[str, Path], historicals_csv: Path) -> Path:
    wb = Workbook()
    # Create tabs in reading order; Cover is filled last because it needs LRP row numbers.
    wb.active.title = "Cover"
    drv_rows = _assumptions_sheet(wb, cfg)
    hist_rows = _historicals_sheet(wb, historicals_csv)
    lrp_rows = _lrp_sheet(wb, cfg, drv_rows, hist_rows)
    _guidance_sheet(wb, cfg, lrp_rows)
    _scenarios_sheet(wb, results, cfg, charts)
    _sensitivity_sheet(wb, tornado_df, grid, cfg, charts.get("tornado"))
    _investment_sheet(wb, cfg)
    _sources_sheet(wb)
    _cover_sheet(wb, cfg, lrp_rows)
    # Put LRP right after Cover, Assumptions next.
    order = ["Cover", "LRP", "Assumptions", "Guidance Check", "Scenarios", "Sensitivity",
             "Investment Case", "Historicals", "Sources"]
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
