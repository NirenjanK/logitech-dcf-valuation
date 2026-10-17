# Logitech Financial Modelling and Valuation (DCF)

A five-year discounted cash flow valuation of Logitech International S.A. (NASDAQ: LOGI), built in Excel with live formulas.

**Status:** DCF model complete. Peer comparables and the written valuation note are next.

## What the model does
- **Historicals (FY2022–FY2026):** income statement, balance sheet and cash flow, entered from Logitech's SEC filings and tied out with 16 integrity checks (all pass).
- **Revenue build:** eight product categories (Gaming, Keyboards & Combos, Pointing Devices, Video Collaboration, Webcams, Tablet Accessories, Headsets, Other) forecast for FY2027–FY2031 with growth tapering towards 3%.
- **Cash flow forecast:** gross margin, operating costs, tax, depreciation, capex and working capital drivers produce unlevered free cash flow.
- **Cost of capital:** CAPM (risk-free rate, Blume-adjusted beta, equity risk premium). Logitech carries no financial debt, so WACC equals the cost of equity.
- **Terminal value:** perpetuity growth, cross-checked with an exit EV/EBITDA multiple.
- **Sensitivity analysis:** WACC × terminal growth and WACC × exit multiple grids, written as live formulas rather than Excel data tables.
- **Scenarios:** Bear / Base / Bull selector on the `Inputs` sheet.

## Base-case result (valuation date 17 Oct 2026)

| | Value |
|---|---|
| WACC | 8.4% |
| Terminal growth | 2.5% |
| Enterprise value | USD 13.2 bn |
| Implied value per share (perpetuity growth) | USD 104.26 |
| Implied value per share (exit multiple, 14x) | USD 106.00 |
| Share price, 16 Oct 2026 | USD 104.41 |
| Terminal value share of EV | 78% |

| Scenario | Value per share (perpetuity) | vs. price |
|---|---|---|
| Bear (growth −2pp, gross margin −1pp) | USD 91.95 | −11.9% |
| Base | USD 104.26 | −0.1% |
| Bull (growth +2pp, gross margin +1pp) | USD 118.04 | +13.1% |

The sensitivity grid spans roughly USD 83–147 per share, so the base case should be read as "the market price is consistent with these assumptions", not as a precise target.

## Key assumptions and judgement calls
- Share-based compensation is treated as a real cost: it stays inside operating costs and is **not** added back.
- Terminal value is discounted from the end of FY2031 (conservative versus a mid-year terminal convention).
- Cash is the 31 March 2026 balance; there is no debt. Operating leases are excluded.
- The first forecast year includes only the part of FY2027 remaining after the valuation date (stub period).
- Risk-free rate 5.10%, raw beta 0.59 and equity risk premium 4.5% are inputs on the `Inputs` sheet and can be changed.

## Limitations
- Historical figures come from Logitech's full-year results releases (Form 8-K Ex. 99.1), which are preliminary; audited 10-K numbers may differ marginally.
- Category history covers FY2023–FY2026 only.
- No peer-comparables analysis yet; the football field uses the 52-week range and the two DCF methods.
- Forecast drivers are my own analyst assumptions, not company guidance. Logitech guided only Q1 FY2027 (sales +4–6% in USD).

## Files
- `Logitech_DCF_Model.xlsx` – the model (sheets: Summary, Historicals, Inputs, DCF, Sensitivity, Checks)
- `scripts/build_model.py` – Python (openpyxl) script that generates the workbook

## Sources
- Logitech Form 8-K Ex. 99.1, Q4 and full-year results: [FY2026](https://www.sec.gov/Archives/edgar/data/0001032975/000103297526000014/exhibit991q4fy26.htm), [FY2024](https://www.sec.gov/Archives/edgar/data/1032975/000103297524000015/exhibit991q4fy24.htm), [FY2022](https://www.sec.gov/Archives/edgar/data/1032975/000103297522000008/exhibit991q4fy22ng.htm)
- Logitech FY2026 [Form 10-K](https://www.sec.gov/Archives/edgar/data/0001032975/000103297526000021/logi-20260331.htm)
- Market data: stockanalysis.com (price, beta, 52-week range), FRED series DGS10 (10-year Treasury)

*Educational project. Not investment advice.*
