# ServiceNow Long-Range Plan Model

A driver-based, five-year financial plan for ServiceNow (NYSE: NOW), built from the company's public filings. It forecasts revenue, margins, operating expenses, headcount, cash flow and SaaS metrics for FY2026–FY2030 under three scenarios, measures which assumptions matter most, and includes a tool for evaluating investment proposals against the plan.

The model exists twice, and the two versions agree to the cent:

- **Python** (`lrp/`), so assumptions live in one readable file and every run is reproducible.
- **Excel** (`outputs/ServiceNow_LRP_Model.xlsx`), with live formulas. Change a driver or switch scenarios and the whole plan recalculates.

> Independent analysis from public information. Not affiliated with or endorsed by ServiceNow. Forecasts are illustrative.

---

## The answer first

| FY2030 | Base | Upside | Downside |
|---|---:|---:|---:|
| Total revenue ($M) | 30,675 | 32,821 | 25,989 |
| Subscription growth | 16.0% | 18.0% | 10.0% |
| Non-GAAP operating margin | 34.2% | 36.9% | 31.2% |
| Free cash flow ($M) | 10,614 | 12,513 | 7,575 |
| FCF margin | 34.6% | 38.1% | 29.1% |
| Headcount | 48,533 | 48,205 | 44,359 |
| Rule of 40 | 50.6% | 56.1% | 39.1% |

**What the plan says:**

1. **Growth is the decision that matters.** Two points of subscription growth a year, up or down, moves FY2030 free cash flow by roughly ±$950M. At the shocks tested, that is more than twice the effect of any cost lever. The central planning question is whether AI products (Now Assist, AI Agents) grow faster than AI shrinks seat-based licences.
2. **Margin expansion comes from operating leverage, led by sales & marketing.** In the base case, non-GAAP operating margin rises ~270 bp to 34.2% by FY2030. Opex falls 4.6 points as a share of revenue, 2.9 of them in sales & marketing (32.4% to 29.5%), partly offset by stock comp falling as a share of revenue (more of the cost base is paid in cash, which non-GAAP margins count).
3. **FCF margin stays flat even as operating margin rises.** Working capital inflows from upfront billing grow more slowly than revenue as growth decelerates, and cash taxes rise. A plan that promises both decelerating growth and rising FCF margin needs a reason, and this model makes that tension visible.
4. **AI productivity changes the headcount plan more than the revenue plan.** The upside case reaches ~$2.1B more revenue than the base with *fewer* employees, because revenue per employee compounds faster.

![Revenue by scenario](outputs/charts/revenue_by_scenario.png)
![FCF margin by scenario](outputs/charts/fcf_margin_by_scenario.png)
![FCF sensitivity](outputs/charts/fcf_tornado.png)

---

## How the plan ties to the current year

A long-range plan is only credible if its first year matches the operating plan. FY2026 is calibrated to ServiceNow's guidance from the Q3 2026 earnings release and is the same in every scenario, because three quarters are already reported.

| FY2026 | Model | Company guidance |
|---|---:|---:|
| Subscription revenue ($M) | 15,786 | 15,780–15,790 |
| Subscription gross margin (GAAP) | 75.0% | 75% |
| GAAP operating margin | 10.0% | 10% |
| Non-GAAP operating margin | 31.5% | 31.5% |
| Free cash flow margin | 36.0% | 36% |

---

## How the model works

Every forecast line is either a reported figure or one driver applied to another line. There are no hidden plugs.

```
Subscription revenue  = prior year × (1 + growth)
Total revenue         = subscription + professional services
Gross profit (GAAP)   = revenue − cost of revenue   (cost = revenue × (1 − gross margin))
Opex (S&M, R&D, G&A)  = total revenue × % of revenue
GAAP operating income = gross profit − opex
Non-GAAP op. income   = GAAP + stock-based comp + amortization & acquisition costs
Operating cash flow   = non-GAAP op. income − cash taxes + working capital inflow + net interest
Free cash flow        = operating cash flow − capex
Headcount             = revenue ÷ revenue per employee (which grows with AI productivity)
cRPO                  = subscription revenue × cRPO ratio
```

**Scenarios** (all drivers in [`config/assumptions.yaml`](config/assumptions.yaml)):

- **Base:** growth glides from ~22% to 16% as the installed base matures; AI products offset seat-pricing pressure; acquisition amortization (Armis, Veza) fades.
- **Upside:** agentic AI monetizes faster than seats decline; growth holds near 18–20%; faster productivity gains.
- **Downside:** AI agents reduce seats and pricing faster than new products replace them; growth slows to 10%; management protects investment, so margins stall.

**Sensitivity analysis** shifts each driver up and down in FY2027–30, one at a time, and measures FY2030 free cash flow. A two-way table crosses subscription growth with sales & marketing spend.

**Investment case** evaluates an illustrative proposal: funding an AI go-to-market team. It reports NPV, IRR, payback and a decision rule. The example is deliberately borderline. NPV is positive, but 97% of it comes from the terminal value and payback misses a four-year hurdle by a month. The recommendation is to stage-gate it: fund the first tranche, release the second only if AI attach rates hit target.

---

## Workbook tour

| Tab | What it does |
|---|---|
| **Cover** | How to use, colour legend, key outputs for the selected scenario |
| **LRP** | The integrated plan: income statement, cash flow, SaaS metrics. All formulas |
| **Assumptions** | Scenario selector (yellow cell) plus every driver for every scenario, with notes on where each comes from |
| **Guidance Check** | FY2026 model vs. company guidance |
| **Scenarios** | All three scenarios side by side, with charts |
| **Sensitivity** | Tornado table and two-way grid on FY2030 free cash flow |
| **Investment Case** | NPV / IRR / payback evaluator with editable inputs |
| **Historicals** | Reported FY2023–FY2025 figures with sources |
| **Sources** | Links to every filing used |

Standard banking colour code: blue = input, black = formula, green = link to another tab.

---

## Run it

```bash
pip install -r requirements.txt
python run_model.py              # prints the summary, writes charts + workbook to outputs/
pytest                           # 13 tests: historicals tie, identities hold, FY26 ties to guidance
python scripts/check_excel.py    # optional: recalculates the workbook in LibreOffice and
                                 # confirms Excel matches Python in every scenario
```

To test your own view, edit `config/assumptions.yaml` and re-run, or change the blue cells in the workbook.

```
├── config/assumptions.yaml   # every driver, scenario, sensitivity and investment input
├── data/historicals.csv      # reported FY2023–25 figures with sources
├── lrp/
│   ├── model.py              # forecast engine and guidance check
│   ├── analysis.py           # tornado, two-way table, NPV/IRR/payback
│   ├── excel_export.py       # writes the formula-driven workbook
│   └── charts.py             # PNG charts
├── scripts/check_excel.py    # Excel vs. Python reconciliation
├── tests/test_model.py
└── outputs/                  # workbook, charts, summary.md
```

---

## How AI was used

I built this with Claude as a coding and drafting partner. To keep the numbers trustworthy:

- Every historical figure was checked against ServiceNow's earnings releases and SEC filings (see `data/historicals.csv`).
- FY2026 was calibrated to the company's guidance rather than to a model's guess.
- The Excel formulas are checked against the independent Python calculation in all three scenarios (`scripts/check_excel.py`).
- The scenario theses and the investment recommendation are my judgment.

---

## Limitations

- Annual model. A planning team would build quarterly, by product line and geography, with bookings (NNACV) driving revenue.
- Non-GAAP operating income for FY2023–25 is reconstructed from the reported rounded margin, so the historical amortization line is approximate.
- Cash taxes, working capital capture and net interest are calibrated assumptions, not reported line items.
- The 2026 acquisitions are captured only through margin drivers, not modelled as separate businesses.
- The investment case is illustrative and uses no ServiceNow data.

## Sources

- [ServiceNow Q3 2026 earnings release (Form 8-K)](https://www.sec.gov/Archives/edgar/data/0001373715/000137371526000092/erq3fy26.htm)
- [ServiceNow Q4 & FY2025 earnings release](https://s205.q4cdn.com/916135447/files/doc_financials/2025/q4/ServiceNow-Reports-Fourth-Quarter-and-Full-Year-2025-Financial-Results.pdf)
- [ServiceNow Q4 & FY2023 earnings release](https://www.businesswire.com/news/home/20240124662293/en/ServiceNow-Reports-Fourth-Quarter-and-Full-Year-2023-Financial-Results)
- [ServiceNow Q4 2025 fact sheet](https://s205.q4cdn.com/916135447/files/doc_downloads/fact-sheet/q4-fact-sheet.pdf)
- [Year-end employee counts (from Form 10-K)](https://stockanalysis.com/stocks/now/employees/)
