"""
Financial Modeler
Architecture: Object-Oriented, Type-Safe, CFI-Standard Formatting
Features: Dynamic Scaling (B/M/K), Full CFI DCF, Debt Waterfall, Monte Carlo
Dependencies: yfinance, pandas, numpy, xlsxwriter
"""

import logging
import datetime
import re
import numpy as np
import pandas as pd
import yfinance as yf
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# Configure logging for standard output
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(message)s', datefmt='%H:%M:%S')

# ==============================================================================
# 1. CONSTANT COORDINATE MAPPING (STRICT PREVENTION OF CIRCULAR REFERENCES)
# ==============================================================================

# --- 3-Statement Model Coordinates ---
R_3SM_REV     = 5
R_3SM_COGS    = 6
R_3SM_GP      = 7
R_3SM_OPEX    = 8
R_3SM_EBIT    = 9
R_3SM_INT_EXP = 10
R_3SM_INT_INC = 11
R_3SM_EBT     = 12
R_3SM_TAX     = 13
R_3SM_NI      = 14

R_3SM_CASH    = 17
R_3SM_ASSET   = 18
R_3SM_DEBT    = 19
R_3SM_LIAB    = 20
R_3SM_EQ      = 21

R_3SM_CF_NI   = 24
R_3SM_DNA     = 25
R_3SM_CAPEX   = 26
R_3SM_NWC     = 27
R_3SM_CFOF    = 28 
R_3SM_DEBT_I  = 29 
R_3SM_NET_CF  = 30

R_SCH_BEG_CASH = 33
R_SCH_CF_PRE   = 34
R_SCH_MIN_CASH = 35
R_SCH_CADS     = 36
R_SCH_BEG_DEBT = 37
R_SCH_DRAW     = 38
R_SCH_SWEEP    = 39
R_SCH_END_DEBT = 40
R_SCH_END_CASH = 41

# --- Assumptions Sheet Coordinates ---
R_ASM_TRANS_DT = 6
R_ASM_FYE      = 7

R_ASM_PRICE    = 10
R_ASM_SHRS     = 11
R_ASM_DEBT     = 12
R_ASM_CASH     = 13

R_ASM_REV_G    = 16
R_ASM_TAX      = 17
R_ASM_CAPEX    = 18
R_ASM_MIN_C    = 19

R_ASM_WACC     = 22
R_ASM_TGR      = 23
R_ASM_MULT     = 24
R_ASM_INT_D    = 25
R_ASM_INT_C    = 26

# --- DCF Model Coordinates ---
# DCF Schedule (Top Left)
R_DCF_HDR      = 3
R_DCF_DATE     = 4
R_DCF_PER      = 5
R_DCF_FRAC     = 6
R_DCF_EBIT     = 7
R_DCF_TAX      = 8
R_DCF_NOPAT    = 9
R_DCF_DNA      = 10
R_DCF_CAPEX    = 11
R_DCF_NWC      = 12
R_DCF_UFCF     = 13
R_DCF_ENT_EXT  = 14
R_DCF_TX_CF    = 15
R_DCF_DISC     = 16
R_DCF_PV       = 17

# Intrinsic Value Block (Bottom Left)
R_INT_HDR      = 20
R_INT_EV       = 21
R_INT_CASH     = 22
R_INT_DEBT     = 23
R_INT_EQ       = 24
R_INT_EQPS     = 26

# Market Value Block (Bottom Middle)
R_MKT_HDR      = 20
R_MKT_CAP      = 21
R_MKT_DEBT     = 22
R_MKT_CASH     = 23
R_MKT_EV       = 24
R_MKT_EQPS     = 26

# Terminal Value Block (Middle Right)
R_TV_HDR       = 19
R_TV_PG        = 20
R_TV_MULT      = 21
R_TV_AVG       = 22

# Rate of Return Block (Middle Right)
R_ROR_HDR      = 24
R_ROR_UPSIDE   = 25
R_ROR_IRR      = 26

# Market Value vs Intrinsic Value Block (Far Right)
R_MVI_HDR      = 28
R_MVI_MKT      = 29
R_MVI_UPS      = 30
R_MVI_INT      = 31

# Stochastic Scenarios Block (Far Right)
R_MC_HDR       = 33
R_MC_P10       = 34
R_MC_P50       = 35
R_MC_P90       = 36

# Right Block Columns
C_R_LBL        = 9  # Column J
C_R_VAL        = 11 # Column L

# ==============================================================================
# 2. DATA ACQUISITION & SANITIZATION LAYER
# ==============================================================================
class DataFetcher:
    def __init__(self, ticker_symbol: str):
        self.ticker_symbol = ticker_symbol.upper()
        self.ticker = yf.Ticker(self.ticker_symbol)
        self.scale = 1.0 
        
    def _extract_row(self, df: pd.DataFrame, standard_names: list, custom_scale: float = None) -> np.ndarray:
        s = custom_scale if custom_scale else self.scale
        if df is None or df.empty: return np.zeros(4)
        for name in standard_names:
            try:
                if name in df.index:
                    row_data = df.loc[name].iloc[::-1].fillna(0).values
                    return np.array([float(x) / s if x is not None else 0.0 for x in row_data])
            except Exception: continue
        return np.zeros(4)

    def get_financials(self) -> dict:
        logging.info(f"Extracting financial suite for {self.ticker_symbol}...")
        try:
            info = self.ticker.info
            inc, bs, cf = self.ticker.financials, self.ticker.balance_sheet, self.ticker.cashflow
        except Exception as e:
            raise RuntimeError(f"API connectivity failure for {self.ticker_symbol}: {e}")

        if inc is None or inc.empty: raise ValueError(f"No financials available for {self.ticker_symbol}.")

        # Determine Dynamic Scale (Billions, Millions, or Thousands)
        raw_rev = self._extract_row(inc, ['Total Revenue', 'Revenue'], custom_scale=1.0)
        max_rev = np.max(np.abs(raw_rev)) if len(raw_rev) > 0 else 0
        
        if max_rev >= 1e10:
            self.scale, self.scale_name = 1e9, "Billions"
        elif max_rev >= 1e7:
            self.scale, self.scale_name = 1e6, "Millions"
        else:
            self.scale, self.scale_name = 1e3, "Thousands"

        # Multipliers to bridge absolute per-share values with scaled aggregates
        scale_mult_mcap = 1e6 / self.scale  # Price * Shares(M) * Multiplier = Scaled MktCap
        scale_mult_price = self.scale / 1e6 # Scaled EqVal * Multiplier / Shares(M) = Absolute Price

        try: tnx = yf.Ticker("^TNX"); rfr = float(tnx.history(period="1d")['Close'].iloc[-1]) / 100
        except Exception: rfr = 0.045 

        price = float(info.get('currentPrice', info.get('regularMarketPrice', 0.0)))
        shares = float(info.get('sharesOutstanding', 0.0))
        beta = float(info.get('beta', 1.10))
        
        rev = self._extract_row(inc, ['Total Revenue', 'Revenue'])
        cogs = self._extract_row(inc, ['Cost Of Revenue', 'Cost of Goods Sold'])
        opex = self._extract_row(inc, ['Operating Expense', 'Total Operating Expenses'])
        ebit = self._extract_row(inc, ['Operating Income', 'EBIT'])
        pretax = self._extract_row(inc, ['Pretax Income'])
        tax = self._extract_row(inc, ['Tax Provision', 'Income Tax Expense'])
        ni = self._extract_row(inc, ['Net Income'])
        
        int_exp = ebit - pretax
        int_inc = np.zeros_like(ebit) 
        
        cash = self._extract_row(bs, ['Cash And Cash Equivalents', 'Cash'])
        assets = self._extract_row(bs, ['Total Assets'])
        debt = self._extract_row(bs, ['Total Debt', 'Long Term Debt'])
        liab = self._extract_row(bs, ['Total Liabilities Net Minority Interest', 'Total Liabilities'])
        equity = self._extract_row(bs, ['Stockholders Equity', 'Total Equity'])
        
        dna = self._extract_row(cf, ['Depreciation And Amortization', 'Depreciation'])
        capex = self._extract_row(cf, ['Capital Expenditure', 'CapEx'])
        nwc = self._extract_row(cf, ['Change In Working Capital'])

        self.years_hist = min(len(rev), len(assets), len(cf.columns) if cf is not None else 4)
        if self.years_hist == 0: raise ValueError("Insufficient data to establish baseline.")

        current_pretax = pretax[-1] if len(pretax) > 0 else 0.0
        current_tax = tax[-1] if len(tax) > 0 else 0.0
        effective_tax = float(current_tax / current_pretax) if current_pretax > 0 else 0.21

        return {
            'ticker': self.ticker_symbol,
            'years_hist': self.years_hist,
            'scale_name': self.scale_name,
            'scale_mult_mcap': scale_mult_mcap,
            'scale_mult_price': scale_mult_price,
            'price': price,
            'shares_out': shares / 1e6, # Shares always extracted in Millions
            'beta': beta,
            'risk_free_rate': rfr,
            'tax_rate': effective_tax,
            'hist_rev': rev[-self.years_hist:],
            'hist_cogs': cogs[-self.years_hist:],
            'hist_opex': opex[-self.years_hist:],
            'hist_ebit': ebit[-self.years_hist:],
            'hist_int_exp': int_exp[-self.years_hist:],
            'hist_int_inc': int_inc[-self.years_hist:],
            'hist_pretax': pretax[-self.years_hist:],
            'hist_tax': tax[-self.years_hist:],
            'hist_ni': ni[-self.years_hist:],
            'hist_cash': cash[-self.years_hist:],
            'hist_assets': assets[-self.years_hist:],
            'hist_debt': debt[-self.years_hist:],
            'hist_liab': liab[-self.years_hist:],
            'hist_equity': equity[-self.years_hist:],
            'hist_dna': dna[-self.years_hist:],
            'hist_capex': capex[-self.years_hist:],
            'hist_nwc': nwc[-self.years_hist:],
            'current_cash': float(cash[-1]) if len(cash) > 0 else 0.0,
            'current_debt': float(debt[-1]) if len(debt) > 0 else 0.0,
            'current_capex': float(abs(capex[-1])) if len(capex) > 0 else 0.0
        }

# ==============================================================================
# 3. FINANCIAL ENGINE LAYER
# ==============================================================================
class FinancialEngine:
    def __init__(self, data: dict):
        self.data = data
        self.margins = self._calculate_margins()
        self.wacc = self._calculate_wacc()
        
    def _safe_mean_ratio(self, num_arr: np.ndarray, den_arr: np.ndarray) -> float:
        valid_indices = den_arr != 0
        if not np.any(valid_indices): return 0.0
        return float(np.mean(num_arr[valid_indices] / den_arr[valid_indices]))

    def _calculate_margins(self) -> dict:
        rev = self.data['hist_rev']
        margins = {}
        for key in ['cogs', 'opex', 'dna', 'capex', 'nwc']:
            hist_array = self.data[f'hist_{key}']
            if key == 'capex': margins[key] = self._safe_mean_ratio(np.abs(hist_array), rev)
            else: margins[key] = self._safe_mean_ratio(hist_array, rev)
        return margins

    def _calculate_wacc(self) -> float:
        mkt_cap_scaled = self.data['price'] * self.data['shares_out'] * self.data['scale_mult_mcap']
        debt_scaled = self.data['current_debt']
        total_capital = mkt_cap_scaled + debt_scaled
        
        weight_eq = mkt_cap_scaled / total_capital if total_capital > 0 else 1.0
        weight_d = debt_scaled / total_capital if total_capital > 0 else 0.0
        
        cost_of_equity = self.data['risk_free_rate'] + (self.data['beta'] * 0.055)
        cost_of_debt = 0.05 
        
        wacc = (weight_eq * cost_of_equity) + (weight_d * cost_of_debt * (1 - self.data['tax_rate']))
        if np.isnan(wacc) or wacc <= 0: return 0.08
        return float(wacc)

    def run_monte_carlo(self, iterations: int = 1000) -> dict:
        logging.info(f"Executing Monte Carlo simulation ({iterations} paths)...")
        np.random.seed(42) 
        
        base_rev_growth, base_tgr = 0.10, 0.025
        sim_wacc = np.random.normal(loc=self.wacc, scale=0.005, size=iterations)
        sim_tgr = np.random.normal(loc=base_tgr, scale=0.002, size=iterations)
        sim_growth = np.random.normal(loc=base_rev_growth, scale=0.015, size=iterations)
        
        implied_prices = []
        latest_rev = self.data['hist_rev'][-1] if len(self.data['hist_rev']) > 0 else 0.0
        
        for i in range(iterations):
            if sim_wacc[i] <= sim_tgr[i]: continue 
            
            c_rev = latest_rev
            fcfs = []
            for _ in range(5):
                c_rev *= (1 + sim_growth[i])
                c_ebit = c_rev - (c_rev * self.margins['cogs']) - (c_rev * self.margins['opex'])
                c_tax = c_ebit * self.data['tax_rate'] 
                ufcf = c_ebit - c_tax + (c_rev * self.margins['dna']) - (c_rev * self.margins['capex']) + (c_rev * self.margins['nwc'])
                fcfs.append(ufcf)
                
            fcfs = np.array(fcfs)
            dfs = np.array([(1 + sim_wacc[i])**t for t in range(1, 6)])
            pv_fcf = np.sum(fcfs / dfs)
            
            tv = (fcfs[-1] * (1 + sim_tgr[i])) / (sim_wacc[i] - sim_tgr[i])
            eq_val_scaled = pv_fcf + (tv / ((1 + sim_wacc[i])**5)) - self.data['current_debt'] + self.data['current_cash']
            
            if self.data['shares_out'] > 0: 
                implied_prices.append((eq_val_scaled * self.data['scale_mult_price']) / self.data['shares_out'])
                
        prices = np.array(implied_prices)
        if len(prices) == 0: prices = np.array([self.data['price']])
        
        return {
            'p10': float(np.percentile(prices, 10)),
            'p50': float(np.median(prices)),
            'p90': float(np.percentile(prices, 90))
        }

# ==============================================================================
# 4. EXCEL ARCHITECTURE & EXPORT LAYER
# ==============================================================================
class WorkbookBuilder:
    def __init__(self, data: dict, engine: FinancialEngine, filename: str):
        self.data, self.engine, self.filename = data, engine, filename
        self.writer = pd.ExcelWriter(self.filename, engine='xlsxwriter')
        self.wb = self.writer.book
        self.formula_log = []
        self._init_styles()
        
    def _init_styles(self):
        self.s = {
            'title': self.wb.add_format({'bold': True, 'bg_color': '#b4c6e7', 'font_color': '#203764', 'font_size': 14, 'valign': 'vcenter'}),
            'header': self.wb.add_format({'bold': True, 'bottom': 1, 'align': 'left'}),
            'header_dark': self.wb.add_format({'bold': True, 'bottom': 2, 'bg_color': '#203764', 'font_color': '#FFFFFF', 'align': 'left'}),
            'year_head': self.wb.add_format({'bold': True, 'bottom': 1, 'align': 'center'}),
            'year_head_r': self.wb.add_format({'bold': True, 'bottom': 1, 'align': 'right'}),
            
            'data': self.wb.add_format({'num_format': '#,##0.0', 'align': 'right'}),
            'data_date': self.wb.add_format({'num_format': 'm/d/yy', 'align': 'right'}),
            'calc': self.wb.add_format({'font_color': '#0070C0', 'num_format': '#,##0.0', 'align': 'right'}),
            'calc_dec': self.wb.add_format({'font_color': '#0070C0', 'num_format': '0.00', 'align': 'right'}),
            
            'input': self.wb.add_format({'font_color': '#0070C0', 'num_format': '#,##0.0', 'bg_color': '#FFFFCC', 'border': 1}),
            'input_pct': self.wb.add_format({'font_color': '#0070C0', 'num_format': '0.0%', 'bg_color': '#FFFFCC', 'border': 1}),
            'input_dec': self.wb.add_format({'font_color': '#0070C0', 'num_format': '#,##0.00', 'bg_color': '#FFFFCC', 'border': 1}),
            'input_date': self.wb.add_format({'font_color': '#0070C0', 'num_format': 'm/d/yy', 'bg_color': '#FFFFCC', 'border': 1}),
            'input_curr': self.wb.add_format({'font_color': '#0070C0', 'num_format': '$#,##0.00', 'bg_color': '#FFFFCC', 'border': 1}),
            
            'legend': self.wb.add_format({'font_color': '#0070C0', 'bg_color': '#FFFFCC', 'border': 1, 'align': 'center'}),
            
            'total': self.wb.add_format({'num_format': '#,##0.0', 'top': 1, 'align': 'right'}),
            'total_btm': self.wb.add_format({'num_format': '#,##0.0', 'top': 1, 'bottom': 1, 'align': 'right'}),
            'pct': self.wb.add_format({'num_format': '0.0%', 'align': 'right'}),
            
            'curr': self.wb.add_format({'num_format': '$#,##0.00', 'align': 'right'}),
            'curr_calc': self.wb.add_format({'font_color': '#0070C0', 'num_format': '$#,##0.00', 'align': 'right'}),
            'curr_bold': self.wb.add_format({'bold': True, 'num_format': '$#,##0.00', 'align': 'right', 'top': 1}),
            'unit': self.wb.add_format({'font_color': '#7F7F7F', 'align': 'left', 'italic': True})
        }

    def safe_write(self, ws, r: int, c: int, val, fmt=None):
        if isinstance(val, (float, int, np.float64, np.int64)):
            if np.isnan(val) or np.isinf(val): ws.write(r, c, 0.0, fmt)
            else: ws.write(r, c, float(val), fmt)
        elif isinstance(val, str) and val.startswith('='):
            ws.write_formula(r, c, val, fmt)
            self.formula_log.append((ws.name, r, c, val))
        else: ws.write(r, c, str(val), fmt)

    def audit_circular_references(self):
        circ_found = False
        for sheet, r, c, formula in self.formula_log:
            pattern = rf"\b{xl_rowcol_to_cell(r, c)}\b"
            if re.search(pattern, formula):
                logging.error(f"[CIRCULAR REFERENCE] Sheet: '{sheet}', Cell {xl_rowcol_to_cell(r, c)} references itself: {formula}")
                circ_found = True
        if not circ_found: logging.info("Model topology passed (0 circular references).")

    def _setup_sheet(self, name: str, title: str, layout='standard') -> object:
        ws = self.wb.add_worksheet(name)
        ws.hide_gridlines(2)
        if layout == 'dcf':
            ws.set_column('A:A', 28); ws.set_column('B:B', 22); ws.set_column('C:I', 15)
            ws.set_column('J:J', 32); ws.set_column('K:K', 5); ws.set_column('L:L', 18)
            ws.merge_range('A1:L2', f" {title}", self.s['title'])
        elif layout == 'asm':
            ws.set_column('A:A', 35); ws.set_column('B:B', 28); ws.set_column('C:C', 15)
            ws.merge_range('A1:C2', f" {title}", self.s['title'])
        else:
            ws.set_column('A:A', 30); ws.set_column('B:H', 15); ws.set_column('I:K', 15)
            ws.merge_range('A1:K2', f" {title}", self.s['title'])
        return ws

    def _write_timeline_headers(self, ws, start_row, start_col):
        yh = self.data['years_hist']
        for i in range(yh): self.safe_write(ws, start_row, start_col + i, f"Hist Y{i+1}", self.s['year_head'])
        for p in range(5): self.safe_write(ws, start_row, start_col + yh + p, f"Proj Y{p+1}", self.s['year_head'])

    def build_assumptions(self):
        ws = self._setup_sheet('Assumptions', f"{self.data['ticker']} - Drivers & Assumptions", layout='asm')
        
        self.safe_write(ws, 3, 0, 'Legend:')
        self.safe_write(ws, 3, 1, 'Yellow Fill / Blue Text', self.s['legend'])
        self.safe_write(ws, 3, 2, '= Adjustable Input', self.s['unit'])
        
        t_date = datetime.date.today()
        fye_date = t_date.replace(month=12, day=31)
        scale = self.data['scale_name']
        
        self.safe_write(ws, R_ASM_TRANS_DT - 1, 0, '1. Timing Parameters', self.s['header_dark'])
        self.safe_write(ws, R_ASM_TRANS_DT, 0, 'Transaction Date'); self.safe_write(ws, R_ASM_TRANS_DT, 1, t_date, self.s['input_date'])
        self.safe_write(ws, R_ASM_FYE, 0, 'Fiscal Year End'); self.safe_write(ws, R_ASM_FYE, 1, fye_date, self.s['input_date'])
        
        self.safe_write(ws, R_ASM_PRICE - 1, 0, '2. Market Data & Capital Structure', self.s['header_dark'])
        self.safe_write(ws, R_ASM_PRICE, 0, 'Current Share Price'); self.safe_write(ws, R_ASM_PRICE, 1, self.data['price'], self.s['input_curr']); self.safe_write(ws, R_ASM_PRICE, 2, '$', self.s['unit'])
        self.safe_write(ws, R_ASM_SHRS, 0, 'Shares Outstanding (M)'); self.safe_write(ws, R_ASM_SHRS, 1, self.data['shares_out'], self.s['input_dec']); self.safe_write(ws, R_ASM_SHRS, 2, 'Millions', self.s['unit'])
        self.safe_write(ws, R_ASM_DEBT, 0, 'Total Debt'); self.safe_write(ws, R_ASM_DEBT, 1, self.data['current_debt'], self.s['input']); self.safe_write(ws, R_ASM_DEBT, 2, scale, self.s['unit'])
        self.safe_write(ws, R_ASM_CASH, 0, 'Total Cash'); self.safe_write(ws, R_ASM_CASH, 1, self.data['current_cash'], self.s['input']); self.safe_write(ws, R_ASM_CASH, 2, scale, self.s['unit'])
        
        self.safe_write(ws, R_ASM_REV_G - 1, 0, '3. Operating Assumptions', self.s['header_dark'])
        self.safe_write(ws, R_ASM_REV_G, 0, 'Base Revenue Growth'); self.safe_write(ws, R_ASM_REV_G, 1, 0.12, self.s['input_pct']); self.safe_write(ws, R_ASM_REV_G, 2, '%', self.s['unit'])
        self.safe_write(ws, R_ASM_TAX, 0, 'Effective Tax Rate'); self.safe_write(ws, R_ASM_TAX, 1, self.data['tax_rate'], self.s['input_pct']); self.safe_write(ws, R_ASM_TAX, 2, '%', self.s['unit'])
        self.safe_write(ws, R_ASM_CAPEX, 0, 'CapEx Baseline'); self.safe_write(ws, R_ASM_CAPEX, 1, self.data['current_capex'], self.s['input']); self.safe_write(ws, R_ASM_CAPEX, 2, scale, self.s['unit'])
        self.safe_write(ws, R_ASM_MIN_C, 0, 'Minimum Cash Target'); self.safe_write(ws, R_ASM_MIN_C, 1, 0.05, self.s['input_pct']); self.safe_write(ws, R_ASM_MIN_C, 2, '% of Rev', self.s['unit'])
        
        self.safe_write(ws, R_ASM_WACC - 1, 0, '4. Valuation & Return Assumptions', self.s['header_dark'])
        self.safe_write(ws, R_ASM_WACC, 0, 'WACC (Discount Rate)'); self.safe_write(ws, R_ASM_WACC, 1, self.engine.wacc, self.s['input_pct']); self.safe_write(ws, R_ASM_WACC, 2, '%', self.s['unit'])
        self.safe_write(ws, R_ASM_TGR, 0, 'Terminal Growth Rate'); self.safe_write(ws, R_ASM_TGR, 1, 0.025, self.s['input_pct']); self.safe_write(ws, R_ASM_TGR, 2, '%', self.s['unit'])
        self.safe_write(ws, R_ASM_MULT, 0, 'EV/EBITDA Multiple'); self.safe_write(ws, R_ASM_MULT, 1, 15.0, self.s['input_dec']); self.safe_write(ws, R_ASM_MULT, 2, 'x', self.s['unit'])
        self.safe_write(ws, R_ASM_INT_D, 0, 'Interest Rate on Debt'); self.safe_write(ws, R_ASM_INT_D, 1, 0.05, self.s['input_pct']); self.safe_write(ws, R_ASM_INT_D, 2, '%', self.s['unit'])
        self.safe_write(ws, R_ASM_INT_C, 0, 'Interest Rate on Cash'); self.safe_write(ws, R_ASM_INT_C, 1, 0.02, self.s['input_pct']); self.safe_write(ws, R_ASM_INT_C, 2, '%', self.s['unit'])

    def build_3_statement(self):
        ws = self._setup_sheet('3-Statement Model', f"{self.data['ticker']} - Institutional 3-Statement Model ($ in {self.data['scale_name']})")
        self._write_timeline_headers(ws, 3, 1)
        yh, m = self.data['years_hist'], self.engine.margins
        
        self.safe_write(ws, R_3SM_REV - 1, 0, 'Income Statement', self.s['header_dark'])
        labels_is = [
            (R_3SM_REV, 'Total Revenue', 'hist_rev'), (R_3SM_COGS, 'Cost of Goods Sold', 'hist_cogs'),
            (R_3SM_GP, 'Gross Profit', None), (R_3SM_OPEX, 'Operating Expenses', 'hist_opex'),
            (R_3SM_EBIT, 'EBIT', 'hist_ebit'), (R_3SM_INT_EXP, 'Interest Expense', 'hist_int_exp'),
            (R_3SM_INT_INC, 'Interest Income', 'hist_int_inc'), (R_3SM_EBT, 'Earnings Before Tax (EBT)', 'hist_pretax'),
            (R_3SM_TAX, 'Taxes', 'hist_tax'), (R_3SM_NI, 'Net Income', 'hist_ni')
        ]
        for r, label, hist_key in labels_is:
            self.safe_write(ws, r, 0, label)
            for i in range(yh):
                if hist_key: self.safe_write(ws, r, i+1, self.data[hist_key][i], self.s['data'])
                elif label == 'Gross Profit': self.safe_write(ws, r, i+1, f"={xl_rowcol_to_cell(R_3SM_REV, i+1)}-{xl_rowcol_to_cell(R_3SM_COGS, i+1)}", self.s['total'])
            
            for p in range(5):
                c = yh + p + 1
                if label == 'Total Revenue': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(r, c-1)}*(1+'Assumptions'!{xl_rowcol_to_cell(R_ASM_REV_G, 1)})", self.s['calc'])
                elif label == 'Cost of Goods Sold': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*{m['cogs']:.4f}", self.s['calc'])
                elif label == 'Gross Profit': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}-{xl_rowcol_to_cell(R_3SM_COGS, c)}", self.s['total'])
                elif label == 'Operating Expenses': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*{m['opex']:.4f}", self.s['calc'])
                elif label == 'EBIT': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_GP, c)}-{xl_rowcol_to_cell(R_3SM_OPEX, c)}", self.s['total'])
                elif label == 'Interest Expense': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_DEBT, c-1)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_INT_D, 1)}", self.s['calc'])
                elif label == 'Interest Income': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_CASH, c-1)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_INT_C, 1)}", self.s['calc'])
                elif label == 'Earnings Before Tax (EBT)': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_EBIT, c)}-{xl_rowcol_to_cell(R_3SM_INT_EXP, c)}+{xl_rowcol_to_cell(R_3SM_INT_INC, c)}", self.s['total'])
                elif label == 'Taxes': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_EBT, c)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_TAX, 1)}", self.s['calc'])
                elif label == 'Net Income': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_EBT, c)}-{xl_rowcol_to_cell(R_3SM_TAX, c)}", self.s['total'])

        self.safe_write(ws, R_3SM_CASH - 1, 0, 'Balance Sheet', self.s['header_dark'])
        labels_bs = [
            (R_3SM_CASH, 'Cash & Equivalents', 'hist_cash'), (R_3SM_ASSET, 'Total Assets', 'hist_assets'),
            (R_3SM_DEBT, 'Total Debt', 'hist_debt'), (R_3SM_LIAB, 'Total Liabilities', 'hist_liab'),
            (R_3SM_EQ, 'Stockholders Equity', 'hist_equity')
        ]
        for r, label, hist_key in labels_bs:
            self.safe_write(ws, r, 0, label)
            for i in range(yh): self.safe_write(ws, r, i+1, self.data[hist_key][i], self.s['data'])
            
            for p in range(5):
                c = yh + p + 1
                if label == 'Cash & Equivalents': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_SCH_END_CASH, c)}", self.s['calc']) 
                elif label == 'Total Assets': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(r, c-1)}*1.05", self.s['calc'])
                elif label == 'Total Debt': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_SCH_END_DEBT, c)}", self.s['calc']) 
                elif label == 'Total Liabilities': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(r, c-1)}*1.03", self.s['calc'])
                elif label == 'Stockholders Equity': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_ASSET, c)}-{xl_rowcol_to_cell(R_3SM_LIAB, c)}", self.s['total'])

        self.safe_write(ws, R_3SM_CF_NI - 1, 0, 'Cash Flow Statement', self.s['header_dark'])
        labels_cf = [
            (R_3SM_CF_NI, 'Net Income', 'hist_ni'), (R_3SM_DNA, 'Plus: D&A', 'hist_dna'),
            (R_3SM_CAPEX, 'Less: CapEx', 'hist_capex'), (R_3SM_NWC, 'Changes in NWC', 'hist_nwc'),
            (R_3SM_CFOF, 'Cash Flow (Pre-Financing)', None), (R_3SM_DEBT_I, 'Debt Issued / (Repaid)', None),
            (R_3SM_NET_CF, 'Net Change in Cash', None)
        ]
        for r, label, hist_key in labels_cf:
            self.safe_write(ws, r, 0, label)
            for i in range(yh):
                if hist_key == 'hist_capex': self.safe_write(ws, r, i+1, abs(self.data[hist_key][i]), self.s['data'])
                elif hist_key: self.safe_write(ws, r, i+1, self.data[hist_key][i], self.s['data'])
                elif label == 'Cash Flow (Pre-Financing)': self.safe_write(ws, r, i+1, f"={xl_rowcol_to_cell(R_3SM_CF_NI, i+1)}+{xl_rowcol_to_cell(R_3SM_DNA, i+1)}-{xl_rowcol_to_cell(R_3SM_CAPEX, i+1)}+{xl_rowcol_to_cell(R_3SM_NWC, i+1)}", self.s['total'])
            
            for p in range(5):
                c = yh + p + 1
                if label == 'Net Income': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_NI, c)}", self.s['calc'])
                elif label == 'Plus: D&A': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*{m['dna']:.4f}", self.s['calc'])
                elif label == 'Less: CapEx': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*{m['capex']:.4f}", self.s['calc'])
                elif label == 'Changes in NWC': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*{m['nwc']:.4f}", self.s['calc'])
                elif label == 'Cash Flow (Pre-Financing)': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_CF_NI, c)}+{xl_rowcol_to_cell(R_3SM_DNA, c)}-{xl_rowcol_to_cell(R_3SM_CAPEX, c)}+{xl_rowcol_to_cell(R_3SM_NWC, c)}", self.s['total'])
                elif label == 'Debt Issued / (Repaid)': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_SCH_DRAW, c)}-{xl_rowcol_to_cell(R_SCH_SWEEP, c)}", self.s['calc'])
                elif label == 'Net Change in Cash': self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_3SM_CFOF, c)}+{xl_rowcol_to_cell(R_3SM_DEBT_I, c)}", self.s['total'])

        self.safe_write(ws, R_SCH_BEG_CASH - 1, 0, 'Debt & Cash Waterfall Schedule', self.s['header_dark'])
        labels_sch = [
            (R_SCH_BEG_CASH, 'Beginning Cash'), (R_SCH_CF_PRE, 'CF Pre-Financing'),
            (R_SCH_MIN_CASH, 'Minimum Cash Target'), (R_SCH_CADS, 'Cash Avail. for Debt Service (CADS)'),
            (R_SCH_BEG_DEBT, 'Beginning Debt'), (R_SCH_DRAW, 'Revolver Draw'),
            (R_SCH_SWEEP, 'Debt Sweep (Repayment)'), (R_SCH_END_DEBT, 'Ending Debt'),
            (R_SCH_END_CASH, 'Ending Cash')
        ]
        for r, label in labels_sch: self.safe_write(ws, r, 0, label)
        for p in range(5):
            c = yh + p + 1
            if p == 0: self.safe_write(ws, R_SCH_BEG_CASH, c, f"={xl_rowcol_to_cell(R_3SM_CASH, c-1)}", self.s['calc'])
            else: self.safe_write(ws, R_SCH_BEG_CASH, c, f"={xl_rowcol_to_cell(R_SCH_END_CASH, c-1)}", self.s['calc'])
            
            self.safe_write(ws, R_SCH_CF_PRE, c, f"={xl_rowcol_to_cell(R_3SM_CFOF, c)}", self.s['calc'])
            self.safe_write(ws, R_SCH_MIN_CASH, c, f"={xl_rowcol_to_cell(R_3SM_REV, c)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_MIN_C, 1)}", self.s['calc'])
            self.safe_write(ws, R_SCH_CADS, c, f"={xl_rowcol_to_cell(R_SCH_BEG_CASH, c)}+{xl_rowcol_to_cell(R_SCH_CF_PRE, c)}-{xl_rowcol_to_cell(R_SCH_MIN_CASH, c)}", self.s['total'])
            
            if p == 0: self.safe_write(ws, R_SCH_BEG_DEBT, c, f"={xl_rowcol_to_cell(R_3SM_DEBT, c-1)}", self.s['calc'])
            else: self.safe_write(ws, R_SCH_BEG_DEBT, c, f"={xl_rowcol_to_cell(R_SCH_END_DEBT, c-1)}", self.s['calc'])
            
            self.safe_write(ws, R_SCH_DRAW, c, f"=IF({xl_rowcol_to_cell(R_SCH_CADS, c)}<0, ABS({xl_rowcol_to_cell(R_SCH_CADS, c)}), 0)", self.s['calc'])
            self.safe_write(ws, R_SCH_SWEEP, c, f"=IF({xl_rowcol_to_cell(R_SCH_CADS, c)}>0, MIN({xl_rowcol_to_cell(R_SCH_CADS, c)}, {xl_rowcol_to_cell(R_SCH_BEG_DEBT, c)}), 0)", self.s['calc'])
            self.safe_write(ws, R_SCH_END_DEBT, c, f"={xl_rowcol_to_cell(R_SCH_BEG_DEBT, c)}+{xl_rowcol_to_cell(R_SCH_DRAW, c)}-{xl_rowcol_to_cell(R_SCH_SWEEP, c)}", self.s['total'])
            self.safe_write(ws, R_SCH_END_CASH, c, f"={xl_rowcol_to_cell(R_SCH_BEG_CASH, c)}+{xl_rowcol_to_cell(R_SCH_CF_PRE, c)}+{xl_rowcol_to_cell(R_SCH_DRAW, c)}-{xl_rowcol_to_cell(R_SCH_SWEEP, c)}", self.s['total_btm'])

    def build_dcf_and_valuation(self):
        ws = self._setup_sheet('DCF & Valuation', f"{self.data['ticker']} - Discounted Cash Flow Analysis ($ in {self.data['scale_name']})", layout='dcf')
        yh = self.data['years_hist']
        mult_mcap = self.data['scale_mult_mcap']
        mult_price = self.data['scale_mult_price']
        
        # ----------------------------------------------------------------------
        # 1. DCF SCHEDULE (Top Left)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_DCF_HDR - 1, 0, 'Discounted Cash Flow', self.s['header'])
        self.safe_write(ws, R_DCF_HDR, 2, 'Entry', self.s['year_head_r'])
        for p in range(5): self.safe_write(ws, R_DCF_HDR, 3 + p, f"Year {p+1}", self.s['year_head_r'])
        self.safe_write(ws, R_DCF_HDR, 8, 'Exit', self.s['year_head_r'])

        dcf_labels = [
            (R_DCF_DATE, 'Date'), (R_DCF_PER, 'Time Periods'), (R_DCF_FRAC, 'Year Fraction'),
            (R_DCF_EBIT, 'EBIT'), (R_DCF_TAX, 'Less: Cash Taxes'), (R_DCF_NOPAT, 'NOPAT'),
            (R_DCF_DNA, 'Plus: D&A'), (R_DCF_CAPEX, 'Less: CapEx'), (R_DCF_NWC, 'Less: Changes in NWC'),
            (R_DCF_UFCF, 'Unlevered FCF'), (R_DCF_ENT_EXT, '(Entry)/Exit'),
            (R_DCF_TX_CF, 'Transaction CF'), (R_DCF_DISC, 'Discount Factor'), (R_DCF_PV, 'PV of FCF')
        ]
        
        for r, label in dcf_labels:
            self.safe_write(ws, r, 0, label)
            
            # Entry Col
            if r == R_DCF_DATE: self.safe_write(ws, r, 2, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_TRANS_DT, 1)}", self.s['data_date'])
            elif r == R_DCF_PER: self.safe_write(ws, r, 2, 0, self.s['data'])
            elif r == R_DCF_ENT_EXT: self.safe_write(ws, r, 2, f"=-{xl_rowcol_to_cell(R_MKT_EV, 4)}", self.s['calc']) 
            elif r == R_DCF_TX_CF: self.safe_write(ws, r, 2, f"={xl_rowcol_to_cell(R_DCF_ENT_EXT, 2)}", self.s['total'])
            
            # Projections
            for p in range(5):
                c = p + 3
                sm_col = xl_rowcol_to_cell(0, yh + p + 1)[0]
                
                if r == R_DCF_DATE: self.safe_write(ws, r, c, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_FYE, 1)}+({p}*365)", self.s['data_date'])
                elif r == R_DCF_PER: self.safe_write(ws, r, c, p+1, self.s['data'])
                elif r == R_DCF_FRAC: self.safe_write(ws, r, c, 1.00, self.s['calc_dec'])
                elif r == R_DCF_EBIT: self.safe_write(ws, r, c, f"='3-Statement Model'!{sm_col}{R_3SM_EBIT+1}", self.s['calc'])
                elif r == R_DCF_TAX: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_EBIT, c)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_TAX, 1)}", self.s['calc'])
                elif r == R_DCF_NOPAT: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_EBIT, c)}-{xl_rowcol_to_cell(R_DCF_TAX, c)}", self.s['total'])
                elif r == R_DCF_DNA: self.safe_write(ws, r, c, f"='3-Statement Model'!{sm_col}{R_3SM_DNA+1}", self.s['calc'])
                elif r == R_DCF_CAPEX: self.safe_write(ws, r, c, f"='3-Statement Model'!{sm_col}{R_3SM_CAPEX+1}", self.s['calc'])
                elif r == R_DCF_NWC: self.safe_write(ws, r, c, f"='3-Statement Model'!{sm_col}{R_3SM_NWC+1}", self.s['calc'])
                elif r == R_DCF_UFCF: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_NOPAT, c)}+{xl_rowcol_to_cell(R_DCF_DNA, c)}-{xl_rowcol_to_cell(R_DCF_CAPEX, c)}+{xl_rowcol_to_cell(R_DCF_NWC, c)}", self.s['total'])
                elif r == R_DCF_TX_CF:
                    if p < 4: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_UFCF, c)}", self.s['total'])
                    else: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_UFCF, c)}+{xl_rowcol_to_cell(R_TV_AVG, C_R_VAL)}", self.s['total'])
                elif r == R_DCF_DISC: self.safe_write(ws, r, c, f"=1/((1+'Assumptions'!{xl_rowcol_to_cell(R_ASM_WACC, 1)})^{c-2})", self.s['calc_dec'])
                elif r == R_DCF_PV: self.safe_write(ws, r, c, f"={xl_rowcol_to_cell(R_DCF_UFCF, c)}*{xl_rowcol_to_cell(R_DCF_DISC, c)}", self.s['total'])

            # Exit Col
            if r == R_DCF_DATE: self.safe_write(ws, r, 8, f"={xl_rowcol_to_cell(R_DCF_DATE, 7)}", self.s['data_date'])
            elif r == R_DCF_ENT_EXT: self.safe_write(ws, r, 8, f"={xl_rowcol_to_cell(R_TV_AVG, C_R_VAL)}", self.s['calc']) 
            elif r == R_DCF_TX_CF: self.safe_write(ws, r, 8, f"={xl_rowcol_to_cell(R_DCF_ENT_EXT, 8)}", self.s['total'])

        # Transaction CF array for IRR
        self.safe_write(ws, R_DCF_TX_CF+1, 0, 'Transaction CF')
        for c in [2, 3, 4, 5, 6]: self.safe_write(ws, R_DCF_TX_CF+1, c, f"={xl_rowcol_to_cell(R_DCF_TX_CF, c)}", self.s['data'])
        self.safe_write(ws, R_DCF_TX_CF+1, 7, f"={xl_rowcol_to_cell(R_DCF_TX_CF, 7)}+{xl_rowcol_to_cell(R_DCF_TX_CF, 8)}", self.s['data'])

        # ----------------------------------------------------------------------
        # 2. INTRINSIC VALUE (Bottom Left)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_INT_HDR - 1, 0, 'Intrinsic Value', self.s['header'])
        
        rng_pv = f"{xl_rowcol_to_cell(R_DCF_PV, 3)}:{xl_rowcol_to_cell(R_DCF_PV, 7)}"
        self.safe_write(ws, R_INT_EV, 0, 'Enterprise Value')
        self.safe_write(ws, R_INT_EV, 1, f"=SUM({rng_pv})+({xl_rowcol_to_cell(R_TV_AVG, C_R_VAL)}*{xl_rowcol_to_cell(R_DCF_DISC, 7)})", self.s['calc'])
        self.safe_write(ws, R_INT_CASH, 0, 'Plus: Cash')
        self.safe_write(ws, R_INT_CASH, 1, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_CASH, 1)}", self.s['calc'])
        self.safe_write(ws, R_INT_DEBT, 0, 'Less: Debt')
        self.safe_write(ws, R_INT_DEBT, 1, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_DEBT, 1)}", self.s['calc'])
        self.safe_write(ws, R_INT_EQ, 0, 'Equity Value', self.s['total'])
        self.safe_write(ws, R_INT_EQ, 1, f"={xl_rowcol_to_cell(R_INT_EV, 1)}+{xl_rowcol_to_cell(R_INT_CASH, 1)}-{xl_rowcol_to_cell(R_INT_DEBT, 1)}", self.s['total'])
        self.safe_write(ws, R_INT_EQPS, 0, 'Equity Value/Share')
        self.safe_write(ws, R_INT_EQPS, 1, f"=({xl_rowcol_to_cell(R_INT_EQ, 1)}*{mult_price})/'Assumptions'!{xl_rowcol_to_cell(R_ASM_SHRS, 1)}", self.s['curr_calc'])

        # ----------------------------------------------------------------------
        # 3. MARKET VALUE (Bottom Middle)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_MKT_HDR - 1, 3, 'Market Value', self.s['header'])
        
        self.safe_write(ws, R_MKT_CAP, 3, 'Market Cap')
        self.safe_write(ws, R_MKT_CAP, 4, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_PRICE, 1)}*'Assumptions'!{xl_rowcol_to_cell(R_ASM_SHRS, 1)}*{mult_mcap}", self.s['calc'])
        self.safe_write(ws, R_MKT_DEBT, 3, 'Plus: Debt')
        self.safe_write(ws, R_MKT_DEBT, 4, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_DEBT, 1)}", self.s['calc'])
        self.safe_write(ws, R_MKT_CASH, 3, 'Less: Cash')
        self.safe_write(ws, R_MKT_CASH, 4, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_CASH, 1)}", self.s['calc'])
        self.safe_write(ws, R_MKT_EV, 3, 'Enterprise Value', self.s['total'])
        self.safe_write(ws, R_MKT_EV, 4, f"={xl_rowcol_to_cell(R_MKT_CAP, 4)}+{xl_rowcol_to_cell(R_MKT_DEBT, 4)}-{xl_rowcol_to_cell(R_MKT_CASH, 4)}", self.s['total'])
        self.safe_write(ws, R_MKT_EQPS, 3, 'Equity Value/Share')
        self.safe_write(ws, R_MKT_EQPS, 4, f"='Assumptions'!{xl_rowcol_to_cell(R_ASM_PRICE, 1)}", self.s['curr_calc'])

        # ----------------------------------------------------------------------
        # 4. TERMINAL VALUE (Right)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_TV_HDR - 1, C_R_LBL, 'Terminal Value', self.s['header'])
        self.safe_write(ws, R_TV_PG, C_R_LBL, 'Perpetual Growth')
        self.safe_write(ws, R_TV_PG, C_R_VAL, f"=({xl_rowcol_to_cell(R_DCF_UFCF, 7)}*(1+'Assumptions'!{xl_rowcol_to_cell(R_ASM_TGR, 1)}))/('Assumptions'!{xl_rowcol_to_cell(R_ASM_WACC, 1)}-'Assumptions'!{xl_rowcol_to_cell(R_ASM_TGR, 1)})", self.s['calc'])
        self.safe_write(ws, R_TV_MULT, C_R_LBL, 'EV/EBITDA')
        self.safe_write(ws, R_TV_MULT, C_R_VAL, f"=({xl_rowcol_to_cell(R_DCF_EBIT, 7)}+{xl_rowcol_to_cell(R_DCF_DNA, 7)})*'Assumptions'!{xl_rowcol_to_cell(R_ASM_MULT, 1)}", self.s['calc'])
        self.safe_write(ws, R_TV_AVG, C_R_LBL, 'Average', self.s['total'])
        self.safe_write(ws, R_TV_AVG, C_R_VAL, f"=AVERAGE({xl_rowcol_to_cell(R_TV_PG, C_R_VAL)}:{xl_rowcol_to_cell(R_TV_MULT, C_R_VAL)})", self.s['total'])

        # ----------------------------------------------------------------------
        # 5. RATE OF RETURN (Right)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_ROR_HDR - 1, C_R_LBL, 'Rate of Return', self.s['header'])
        self.safe_write(ws, R_ROR_UPSIDE, C_R_LBL, 'Target Price Upside')
        self.safe_write(ws, R_ROR_UPSIDE, C_R_VAL, f"=({xl_rowcol_to_cell(R_INT_EQPS, 1)}/{xl_rowcol_to_cell(R_MKT_EQPS, 4)})-1", self.s['pct'])
        
        self.safe_write(ws, R_ROR_IRR, C_R_LBL, 'Internal Rate of Return (IRR)')
        self.safe_write(ws, R_ROR_IRR, C_R_VAL, f"=XIRR({xl_rowcol_to_cell(R_DCF_TX_CF, 2)}:{xl_rowcol_to_cell(R_DCF_TX_CF, 7)}, {xl_rowcol_to_cell(R_DCF_DATE, 2)}:{xl_rowcol_to_cell(R_DCF_DATE, 7)})", self.s['pct'])

        # ----------------------------------------------------------------------
        # 6. MARKET VS INTRINSIC (Right)
        # ----------------------------------------------------------------------
        self.safe_write(ws, R_MVI_HDR - 1, C_R_LBL, 'Market Value vs Intrinsic', self.s['header'])
        self.safe_write(ws, R_MVI_MKT, C_R_LBL, 'Market Value')
        self.safe_write(ws, R_MVI_MKT, C_R_VAL, f"={xl_rowcol_to_cell(R_MKT_EQPS, 4)}", self.s['curr_calc'])
        self.safe_write(ws, R_MVI_UPS, C_R_LBL, 'Upside')
        self.safe_write(ws, R_MVI_UPS, C_R_VAL, f"={xl_rowcol_to_cell(R_INT_EQPS, 1)}-{xl_rowcol_to_cell(R_MKT_EQPS, 4)}", self.s['curr_calc'])
        self.safe_write(ws, R_MVI_INT, C_R_LBL, 'Intrinsic Value', self.s['total'])
        self.safe_write(ws, R_MVI_INT, C_R_VAL, f"={xl_rowcol_to_cell(R_INT_EQPS, 1)}", self.s['curr_bold'])

        # ----------------------------------------------------------------------
        # 7. STOCHASTIC SCENARIOS (Right)
        # ----------------------------------------------------------------------
        mc_results = self.engine.run_monte_carlo(iterations=1000)
        self.safe_write(ws, R_MC_HDR - 1, C_R_LBL, 'Stochastic Scenarios', self.s['header'])
        self.safe_write(ws, R_MC_P10, C_R_LBL, '10th Percentile (Bear)')
        self.safe_write(ws, R_MC_P10, C_R_VAL, mc_results['p10'], self.s['curr_calc'])
        self.safe_write(ws, R_MC_P50, C_R_LBL, '50th Percentile (Base)')
        self.safe_write(ws, R_MC_P50, C_R_VAL, mc_results['p50'], self.s['curr_calc'])
        self.safe_write(ws, R_MC_P90, C_R_LBL, '90th Percentile (Bull)')
        self.safe_write(ws, R_MC_P90, C_R_VAL, mc_results['p90'], self.s['curr_calc'])

        # ----------------------------------------------------------------------
        # 8. CHARTS (Anchored Safely Below Data at Row 30)
        # ----------------------------------------------------------------------
        chart_cf = self.wb.add_chart({'type': 'column'})
        chart_cf.add_series({
            'categories': ['DCF & Valuation', R_DCF_HDR, 3, R_DCF_HDR, 7],
            'values':     ['DCF & Valuation', R_DCF_UFCF, 3, R_DCF_UFCF, 7],
            'fill':       {'color': '#D9E1F2'},
            'name':       'Unlevered Free Cash Flow'
        })
        chart_cf.set_title({'name': 'Projected Cash Flow Profile'})
        chart_cf.set_y_axis({'num_format': '$#,##0', 'major_gridlines': {'visible': True, 'line': {'color': '#F2F2F2'}}})
        chart_cf.set_legend({'none': True})
        chart_cf.set_size({'width': 380, 'height': 240})
        ws.insert_chart('A30', chart_cf) 

        chart_val = self.wb.add_chart({'type': 'column'})
        chart_val.add_series({
            'categories': ['DCF & Valuation', R_MVI_MKT, C_R_LBL, R_MVI_INT, C_R_LBL],
            'values':     ['DCF & Valuation', R_MVI_MKT, C_R_VAL, R_MVI_INT, C_R_VAL],
            'fill':       {'color': '#D9E1F2'},
            'data_labels': {'value': True, 'num_format': '$#,##0.00'} # Explicit formatting prevents label overflow
        })
        chart_val.set_title({'name': 'Market vs Intrinsic Value'})
        chart_val.set_y_axis({'num_format': '$#,##0', 'major_gridlines': {'visible': True, 'line': {'color': '#F2F2F2'}}})
        chart_val.set_legend({'none': True})
        chart_val.set_size({'width': 380, 'height': 240})
        ws.insert_chart('E30', chart_val) 

    def generate(self):
        self.build_assumptions()
        self.build_3_statement()
        self.build_dcf_and_valuation()
        self.audit_circular_references()
        self.writer.close()
        logging.info(f"Model exported successfully: {self.filename}")

# ==============================================================================
# 5. CLI EXECUTION ENTRY POINT
# ==============================================================================
def main():
    while True:
        try:
            ticker = input("\nEnter Target Ticker (or 'exit'): ").strip().upper()
            if ticker == 'EXIT': break
            if not ticker: continue
            
            fetcher = DataFetcher(ticker)
            data_bundle = fetcher.get_financials()
            engine = FinancialEngine(data_bundle)
            
            filename = f"{ticker} Model {datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.xlsx"
            builder = WorkbookBuilder(data_bundle, engine, filename)
            builder.generate()
            
        except Exception as e:
            logging.error(f"Execution Exception: {e}")

if __name__ == "__main__":
    main()
