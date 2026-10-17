# Logitech Financial Modelling and Valuation (DCF)

A discounted cash flow (DCF) valuation of Logitech International S.A. (NASDAQ: LOGI) in Excel, using five years of public financial statements (FY2022–FY2026).

## Purpose
A DCF estimates what a company is worth today from the cash it is expected to generate in future. I built this to learn the full valuation process on a company whose products I already knew, and to compare the result with the market price. It uses public filings and market data only.

## How the model works
1. **Historicals** – income statement, balance sheet, cash flow and sales by product category from Logitech's SEC filings, with two tie-out checks (balance sheet balances; categories add up to total sales).
2. **Assumptions** – share price, cost of capital inputs, terminal value inputs, and forecast drivers for FY2027–FY2031.
3. **DCF** – forecasts sales by product category, then operating profit and free cash flow; discounts the cash flows at the WACC; adds a terminal value; adds cash; and divides by shares outstanding.
4. **Sensitivity** – how value per share changes with WACC and terminal growth, and with WACC and the exit multiple.

## Result (base case)

| | |
|---|---|
| WACC (required return) | 8.4% = 5.10% risk-free + 0.73 beta × 4.5% equity risk premium |
| Terminal growth rate | 2.5% |
| Enterprise value | USD 12.9 bn |
| Value per share – perpetuity growth method | USD 102.04 |
| Value per share – exit multiple method (14x EBITDA) | USD 103.70 |
| Share price, 16 Oct 2026 | USD 104.41 |
| Terminal value as % of enterprise value | 76% |

The sensitivity tables span roughly USD 81–144 per share, so the result says the market price is consistent with these assumptions rather than giving a precise target.

## Conclusion
- On my base-case assumptions Logitech is worth about **USD 102 per share** (perpetuity growth method) and **USD 104** (exit multiple method), against a share price of USD 104.41. That is within about 2% of the market, so Logitech looks **roughly fairly valued**.
- Put differently, the market price is consistent with sales growth tapering from about 5% to 3% a year and an operating margin of about 17%. At an 8.4% WACC, the current price needs roughly 2.7% growth in perpetuity, or a WACC of about 8.2% at 2.5% growth.
- The answer depends mainly on the discount rate and long-run growth. Terminal value is 76% of enterprise value, and moving WACC by one percentage point changes value by roughly USD 13 to USD 19 per share (USD 89 at 9.4%, USD 121 at 7.4%).
- A gap of a few per cent is within the uncertainty of the assumptions, so the result supports "about fairly valued", not a buy or sell call.
- Next steps: add peer comparables and a bear/bull case on sales growth and margins.

## Key assumptions
- Sales grow by product category, tapering from about 5% to about 3% a year.
- Gross margin stays near FY2026's 43.2%; operating costs stay near FY2026 ratios; tax rate 15% on operating profit.
- Share-based compensation is treated as a real cost (not added back).
- Logitech has no financial debt, so WACC equals the cost of equity.
- Simplification: valuation is as at the 31 March 2026 balance sheet, with cash flows discounted at the end of each year.

## Limitations
- Historical figures come from Logitech's full-year results releases (preliminary); audited 10-K figures may differ slightly.
- Product-category history covers FY2023–FY2026 only.
- Forecast drivers are my own assumptions, not company guidance.
- No peer comparables yet.

## Sources
- Logitech Form 8-K Ex. 99.1 results: [FY2026](https://www.sec.gov/Archives/edgar/data/0001032975/000103297526000014/exhibit991q4fy26.htm), [FY2024](https://www.sec.gov/Archives/edgar/data/1032975/000103297524000015/exhibit991q4fy24.htm), [FY2022](https://www.sec.gov/Archives/edgar/data/1032975/000103297522000008/exhibit991q4fy22ng.htm)
- Market data: stockanalysis.com (price, beta), FRED series DGS10 (10-year Treasury yield)

*Educational project. Not investment advice.*
