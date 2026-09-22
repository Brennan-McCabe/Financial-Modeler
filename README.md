Financial Modeler

A Python tool that pulls historical financial statements, computes forward projections, and generates fully linked, audit-ready Excel models (.xlsx) for 3-statement analysis and DCF valuation.

The goal is to automate the extraction and formula-building workflow in Excel while preserving standard financial formatting, dynamic cell references, and balance sheet integrity.
Key Features

    Automated Financial Ingestion: Fetches multi-year balance sheets, income statements, and cash flow statements via yfinance, along with market cap, live pricing, and historical beta.

    Dynamic Magnitude Scaling: Automatically scales line items to Billions, Millions, or Thousands based on historical revenue baselines so numbers remain legible across different market caps.

    Linked 3-Statement Model: Builds 5-year forecasts using trailing average margin and growth assumptions, maintaining dynamic links across all three statements.

    Debt & Cash Waterfall: Implements cash available for debt service (CADS) logic, modeling automatic revolver draws during deficits, cash sweeps for debt paydown during surpluses, and interest expense/income based on beginning period balances.

    DCF Valuation: Projects Unlevered Free Cash Flow (UFCF) and calculates Enterprise Value and Equity Value using both Perpetual Growth and EV/EBITDA multiple methods.

    Monte Carlo Simulation: Runs 1,000 randomized iterations over WACC, terminal growth rates, and top-line growth to output 10th, 50th, and 90th percentile valuation distributions.

Implementation Details

Programmatically generating linked financial models in Excel creates practical edge cases around formula syntax, grid offsets, and type handling. This engine addresses them through several explicit design choices:

    Formula & Type Sanitization: Passing unhandled NaN or inf values to xlsxwriter breaks workbook creation. The WorkbookBuilder class uses a safe_write() helper that coerces missing numerical data to 0.0 while maintaining strict separation between numeric floats, labels, and raw Excel formula strings.

    Coordinate-Based Formula Generation: Replaces fragile string concatenation (e.g., manual column character math) with coordinate-based cell addressing via xl_rowcol_to_cell, ensuring offsets and relative references stay consistent across statements.

    Circular Reference Checks: Pre-scans generated formulas prior to workbook write using regex pattern matching on target cell coordinates to flag unintended self-referencing loops before compilation.

Tech Stack

    Python

    Data Processing: Pandas, NumPy

    Market Data: yfinance

    Excel Engine: xlsxwriter

Getting Started
Prerequisites
Bash

pip install pandas numpy yfinance xlsxwriter

Usage
Python

from modeler import FinancialModeler

# Initialize and generate workbook
model = FinancialModeler(ticker="AAPL")
model.build(output_path="AAPL_valuation_model.xlsx")
