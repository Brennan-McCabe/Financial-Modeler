# Financial Modeler

An end-to-end Python engine that bridges the gap between programmatic data science and traditional institutional banking. 

This framework dynamically pulls historical equity data, runs predictive financial engines, and programmatically writes fully linked, CFI-standard 3-Statement and DCF models directly into Excel. It is built to automate the most tedious parts of quantitative equity research while preserving the strict formatting and auditability required on the desk.

## 🚀 Key Features

*   **Automated Data Ingestion:** Leverages the Yahoo Finance API to pull historical Income Statements, Balance Sheets, and Cash Flow Statements, alongside live market data and Beta.
*   **Dynamic Scaling Engine:** Automatically scans historical revenue to scale the entire model output to Billions, Millions, or Thousands, ensuring the output is perfectly readable whether you are screening a mega-cap tech giant or a micro-cap pharmaceutical stock.
*   **Institutional 3-Statement Model:** Projects 5 years of financials based on historical trailing average margins and base growth assumptions.
*   **Dynamic Debt & Cash Waterfall:** Actively calculates cash available for debt service (CADS), automatically triggering Revolver draws for shortfalls and debt sweeps for excess cash, with dynamic Interest Expense/Income linked to beginning balances.
*   **Discounted Cash Flow (DCF):** Calculates Unlevered Free Cash Flow (UFCF) and bridges Enterprise Value to Equity Value using both Perpetual Growth and EV/EBITDA exit multiples.
*   **Stochastic Valuation:** Runs a 1,000-iteration Monte Carlo simulation randomizing WACC, Terminal Growth, and Revenue Growth to generate 10th, 50th, and 90th percentile target price scenarios.

## 🧠 Under the Hood: The Code

Building complex financial models via code introduces unique challenges, specifically regarding Excel's strict typing and circular reference errors. This project solves those through strict, object-oriented architecture:

*   **Type-Safe Writing Wrapper:** Passing `NaN` or `INF` floats to `xlsxwriter` causes catastrophic compilation crashes. The `WorkbookBuilder` class utilizes a custom `safe_write()` wrapper that intercepts corrupted data, defaults it safely to `0.0`, and strictly bifurcates string values, numeric floats, and programmatic formulas.
*   **Explicit Coordinate Mapping:** Dynamic string concatenation for Excel cell references (e.g., `chr(ord('A')+c)`) is brittle and prone to off-by-one errors. This framework abandons string math in favor of a strictly defined integer grid and `xl_rowcol_to_cell`, ensuring bulletproof mathematical linkages.
*   **Pre-Compilation Auditing:** Before the `.xlsx` file is saved, an internal engine scans the registry of written formulas using regex boundaries to detect recursive mapping. If a cell references itself, the compilation logs a terminal warning, ensuring the final output is 100% free of circular reference loops.

## 💻 Tech Stack
*   **Language:** Python
*   **Data Manipulation:** Pandas, NumPy
*   **Data Extraction:** `yfinance`
*   **Compilation:** `xlsxwriter`
