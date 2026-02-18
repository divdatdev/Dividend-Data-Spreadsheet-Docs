# Dividend Data — Google Sheets Add-in

<p align="center">
  <img src="https://raw.githubusercontent.com/divdatdev/Dividend-Data-Spreadsheet-Docs/main/logo/LogoNew.png" alt="Dividend Data Logo" width="120"/>
</p>

<p align="center">
  <strong>Access 30+ years of financial data directly in Google Sheets.</strong>
</p>

<p align="center">
  <a href="https://www.dividenddata.com">Website</a> •
  <a href="https://help.dividenddata.com">Help Center</a> •
  <a href="https://www.youtube.com/@DividendData">YouTube</a>
</p>

---

## Quick Reference

> **Jump to:** [Installation](#how-to-install--update) · [Quick Action Sidebar](#quick-action-sidebar) · [Function Builder](#function-builder) · [Function Reference](#function-reference)

| Function | Description |
|----------|-------------|
| [`DIVIDENDDATA_DIVIDENDS`](#1-dividenddata_dividends) | Dividend data (payouts, yields, history, growth) |
| [`DIVIDENDDATA_DIVIDENDS_BATCH`](#2-dividenddata_dividends_batch) | Batch dividend data for multiple stocks |
| [`DIVIDENDDATA_STATEMENT`](#3-dividenddata_statement) | Financial statements (income, balance, cash flow) |
| [`DIVIDENDDATA_METRICS`](#4-dividenddata_metrics) | Individual financial metrics (revenue, EPS, etc.) |
| [`DIVIDENDDATA_RATIOS`](#5-dividenddata_ratios) | Financial ratios (P/E, current ratio, ROE, etc.) |
| [`DIVIDENDDATA_GROWTH`](#6-dividenddata_growth) | Growth metrics (revenue growth, EPS growth, etc.) |
| [`DIVIDENDDATA_QUOTE`](#7-dividenddata_quote) | Stock quotes (price, volume, history) |
| [`DIVIDENDDATA_QUOTE_BATCH`](#8-dividenddata_quote_batch) | Batch quotes for multiple stocks |
| [`DIVIDENDDATA_PROFILE`](#9-dividenddata_profile) | Company profiles (market cap, sector, CEO, etc.) |
| [`DIVIDENDDATA_FUND`](#10-dividenddata_fund) | ETF/Fund data (holdings, expense ratio) |
| [`DIVIDENDDATA_SEGMENTS`](#11-dividenddata_segments) | Revenue segments (product, geographic) |
| [`DIVIDENDDATA_KPIS`](#12-dividenddata_kpis) | Key performance indicators |
| [`DIVIDENDDATA_COMMODITIES`](#13-dividenddata_commodities) | Commodities data (oil, gold, etc.) |
| [`DIVIDENDDATA_CRYPTO`](#14-dividenddata_crypto) | Cryptocurrency data |
| [`DIVIDENDDATA_PRICE_TARGET`](#15-dividenddata_price_target) | Analyst price targets |
| [`DIVIDENDDATA_ESTIMATES`](#16-dividenddata_estimates) | Analyst estimates (EPS, revenue forecasts) |

---

## Google Sheets vs Excel

Function names are consistent across both platforms. The only difference is the separator: Google Sheets uses `_` (underscore) and Excel uses `.` (dot). This is a platform requirement — not a Dividend Data choice.

**The rule: same function name, swap the separator.**

| Function | Google Sheets | Excel |
|----------|--------------|-------|
| Dividends | `=DIVIDENDDATA_DIVIDENDS(...)` | `=DIVIDENDDATA.DIVIDENDS(...)` |
| Dividends Batch | `=DIVIDENDDATA_DIVIDENDS_BATCH(...)` | `=DIVIDENDDATA.DIVIDENDS_BATCH(...)` |
| Statement | `=DIVIDENDDATA_STATEMENT(...)` | `=DIVIDENDDATA.STATEMENT(...)` |
| Metrics | `=DIVIDENDDATA_METRICS(...)` | `=DIVIDENDDATA.METRICS(...)` |
| Ratios | `=DIVIDENDDATA_RATIOS(...)` | `=DIVIDENDDATA.RATIOS(...)` |
| Growth | `=DIVIDENDDATA_GROWTH(...)` | `=DIVIDENDDATA.GROWTH(...)` |
| Quote | `=DIVIDENDDATA_QUOTE(...)` | `=DIVIDENDDATA.QUOTE(...)` |
| Quote Batch | `=DIVIDENDDATA_QUOTE_BATCH(...)` | `=DIVIDENDDATA.QUOTE_BATCH(...)` |
| Profile | `=DIVIDENDDATA_PROFILE(...)` | `=DIVIDENDDATA.PROFILE(...)` |
| Fund | `=DIVIDENDDATA_FUND(...)` | `=DIVIDENDDATA.FUND(...)` |
| Estimates | `=DIVIDENDDATA_ESTIMATES(...)` | `=DIVIDENDDATA.ESTIMATES(...)` |
| Segments | `=DIVIDENDDATA_SEGMENTS(...)` | `=DIVIDENDDATA.SEGMENTS(...)` |
| KPIs | `=DIVIDENDDATA_KPIS(...)` | `=DIVIDENDDATA.KPIS(...)` |
| Commodities | `=DIVIDENDDATA_COMMODITIES(...)` | `=DIVIDENDDATA.COMMODITIES(...)` |
| Crypto | `=DIVIDENDDATA_CRYPTO(...)` | `=DIVIDENDDATA.CRYPTO(...)` |
| Price Target | `=DIVIDENDDATA_PRICE_TARGET(...)` | `=DIVIDENDDATA.PRICE_TARGET(...)` |

Parameters, metrics, and outputs are identical across both platforms.

> **📌 Note for existing Google Sheets users:** If you're using `=DIVIDENDDATA("AAPL", ...)` or `=DIVIDENDDATA_BATCH(...)`, those still work and will continue to work. The new names `DIVIDENDDATA_DIVIDENDS` and `DIVIDENDDATA_DIVIDENDS_BATCH` are the recommended standard for consistency with the Excel version.

# How to Install & Update
**Early Access: Google Sheets**

Use the following link: https://workspace.google.com/marketplace/app/dividend_data/427484236221

### ⚠️ CRITICAL: Authorization Required
**You must check all permission boxes when installing.**

The add-on requires specific permissions to run calculations and display data in your spreadsheet. If you do not accept the authorization prompts, the add-on **will not work**.

**Why?** Google requires your explicit permission for us to fetch financial data from our servers and display it inside your sheet. If you leave boxes unchecked, you are blocking the add-on from accessing the spreadsheet to write the data you requested.

---

### Step-by-Step Installation
1.  **Download:** Click the link above to visit the Google Workspace Marketplace and click **Install**.
2.  **Authorize:** When prompted, select your Google Account. **You must verify/check ALL boxes** to allow the add-on to function.
3.  **Launch:** Open a Google Sheet. Go to **Extensions > Dividend Data > Open Dividend Data Sidebar**.
    * *Note: If you do not see the menu, reload your browser tab.*

![Alt text for the image](https://github.com/divdatdev/Dividend-Data-Spreadsheet-Docs/blob/6faa86ed333f045c9a8bf7892a1e4c8bfd491185/Sheets%20Extensions%20Image.png)

### Troubleshooting: "Permission Denied" or Not Working?
If you are getting errors immediately after installing, you likely missed an authorization checkbox.

**How to Fix or Install the Newest Version:**
1.  Go to **Extensions > Add-ons > Manage Add-ons**.
2.  Find **Dividend Data** and select **Uninstall**.
3.  Refresh your page.
4.  **Reinstall** the add-on.
5.  **IMPORTANT:** When the "Sign in with Google" popup appears, make sure to **Check/Allow ALL permissions**.

---

_Also available for Microsoft Excel! See the [Excel Add-in Documentation](https://github.com/divdatdev/DividendData-Excel-Docs)._

---

## Quick Action Sidebar

The Quick Action Sidebar is the fastest way to get financial data into your spreadsheet. Instead of writing formulas, you search for a stock, pick a data category, and click **Load Data** — a full, formatted table is inserted starting at your currently selected cell.

### How to Open the Sidebar

1. Go to the **Extensions** menu at the top of Google Sheets.
2. Hover over **Dividend Data**.
3. Click **Open Dividend Data Sidebar**.
4. The sidebar panel opens on the right side of your spreadsheet.

> **Tip:** If you don't see "Dividend Data" under Extensions, the add-on may not be installed or your page needs a refresh. Try reloading the browser tab.

### Signing In

When the sidebar opens, you need to sign in with your Dividend Data account.

1. Click the **Sign In** button in the action bar at the top of the sidebar.
2. Follow the authentication prompt.
3. Once signed in, you're ready to pull data.

If you don't have an account yet, visit [DividendData.com](https://www.dividenddata.com/) to create a free account.

### Step 1 — Find a Stock or Fund

At the top of the sidebar, you'll see a search bar labeled **"Search by ticker or company name"**.

- Type a **ticker symbol** (e.g., `MSFT`) or **company name** (e.g., `Microsoft`) to search across 80,000+ supported tickers.
- Select a result from the dropdown. Popular tickers like NVDA, MSFT, AAPL, GOOG, AMZN, and META appear by default when you focus the search bar.
- Once selected, the sidebar shows **example formulas** you can click to copy directly to your clipboard, then paste into any cell. This is useful when you want a quick single data point rather than a full table.

### Step 2 — Load Data into Your Sheet

Below the search area, you'll see the **Quick Action** buttons arranged in a grid:

| Button | What It Loads |
|--------|--------------|
| **Dividends** | Dividend summary table — yields, payouts, growth rates, ex-dividend dates, payment dates |
| **Statements** | Full financial statements — Income Statement, Balance Sheet, or Cash Flow |
| **Metrics** | Key financial metrics — Revenue, EPS, Free Cash Flow, EBITDA, and more |
| **Ratios** | Financial ratios — P/E, ROE, Debt/Equity, Current Ratio, Payout Ratio, etc. |
| **KPIs** | Company-specific key performance indicators (e.g., Azure revenue for MSFT) |
| **Price Targets** | Analyst consensus price targets — high, low, median, number of analysts |

**To load data:**

1. **Search and select a stock** in Step 1.
2. **Click a category button** (e.g., Statements).
3. If you selected **Statements**, choose the statement type: **Income**, **Balance**, or **Cash Flow**.
4. Choose the **Period**: **Annual** or **Quarterly**.
5. **Click your target cell** in the spreadsheet — this is where the top-left corner of the table will go.
6. Click the green **Load Data** button.
7. The full table populates starting at your selected cell, complete with headers.

> **No formulas are inserted.** Load Data pastes static values directly into cells. This is great for quick analysis, one-time snapshots, or when you don't need the data to auto-refresh.

### Additional Sidebar Features

- **Refresh** — Recalculates all Dividend Data formulas in the current sheet. Useful after making changes or if data appears stale.
- **Sign In** — Authenticate your Dividend Data account.
- **Docs** — Quick link to this documentation on GitHub.
- **Dark / Light Mode** — Toggle the sidebar theme with the sun/moon switch in the top-right corner.
- **How to Use** — Expandable tips section at the bottom of the sidebar with a quick-start guide.

### When to Use the Sidebar vs. Formulas

| Use the Sidebar when... | Use formulas when... |
|--------------------------|---------------------|
| You want a full table dumped into your sheet quickly | You want data that auto-refreshes when you reopen the file |
| You're exploring what data is available for a stock | You're building a model or tracker that pulls live data |
| You need a one-time snapshot for a report | You want to reference a single data point in a calculation |
| You're new and want to see what the add-on offers | You know the exact metric you need |


---

## Function Builder

The Function Builder is a visual formula editor built into the add-on. It walks you through every data category, every available metric, and every option — then generates the exact formula for you. No memorization, no guessing syntax. Just configure and copy.

### How to Open the Function Builder

1. Go to the **Extensions** menu at the top of Google Sheets.
2. Hover over **Dividend Data**.
3. Click **Open Formula Builder**.
4. The Function Builder panel opens on the right side of your spreadsheet, replacing the main sidebar.

> To return to the main sidebar, close the Function Builder and reopen via **Extensions → Dividend Data → Open Dividend Data Sidebar**.

### How It Works

The Function Builder is organized into four steps. As you make selections, the formula preview at the bottom updates in real-time.

### Step 1 — Choose Data Category

Select the type of financial data you want from the dropdown menu. There are **13 data categories** available:

| Category | Function | What You Get |
|----------|----------|-------------|
| Dividends | `DIVIDENDDATA_DIVIDENDS` | Payouts, yields, history, growth rates, payout ratios |
| Financials | `DIVIDENDDATA_STATEMENT` | Income Statement, Balance Sheet, Cash Flow |
| Metrics | `DIVIDENDDATA_METRICS` | 100+ individual metrics (Revenue, EPS, FCF, etc.) |
| Ratios | `DIVIDENDDATA_RATIOS` | 50+ financial ratios (P/E, ROE, D/E, etc.) |
| Growth | `DIVIDENDDATA_GROWTH` | Growth rates (revenue growth, EPS growth, etc.) |
| Quotes | `DIVIDENDDATA_QUOTE` | Live price, volume, moving averages, price history |
| Profile | `DIVIDENDDATA_PROFILE` | Company info — sector, industry, CEO, description |
| Funds | `DIVIDENDDATA_FUND` | ETF/fund holdings, expense ratio, sector allocation |
| Estimates | `DIVIDENDDATA_ESTIMATES` | Analyst EPS and revenue forecasts |
| Segments | `DIVIDENDDATA_SEGMENTS` | Product and geographic revenue breakdowns |
| KPIs | `DIVIDENDDATA_KPIS` | Company-specific operational metrics |
| Crypto | `DIVIDENDDATA_CRYPTO` | Cryptocurrency prices and history |
| Commodities | `DIVIDENDDATA_COMMODITIES` | Commodity prices and history (oil, gold, etc.) |
| Price Targets | `DIVIDENDDATA_PRICE_TARGET` | Analyst consensus, high, low, median targets |

### Step 2 — Select Asset

Search for any stock, ETF, or fund by **ticker symbol** or **company name**. The search covers 80,000+ supported tickers.

**Batch Mode:** Check the **Multiple tickers (batch)** checkbox to switch to batch mode. Enter comma-separated tickers (e.g., `AAPL,MSFT,KO`) to pull data for multiple stocks in one formula. Batch mode is available for Dividends and Quotes categories.

### Step 3 — Configure Formula

Once you've chosen a category and asset, the configuration options appear:

- **Metric** — A dropdown showing every available metric for your selected category. For example, choosing "Dividends" reveals options like `summary`, `history`, `fwd_yield`, `5y_cagr`, `payout_ratio`, and many more.
- **Period** — `Annual`, `Quarterly`, or `TTM` (where applicable).
- **Year** — Filter to a specific year (optional).
- **Date Range** — For price history, set a `From` and `To` date.
- **Show Headers** — Check this box to include column/row headers in the output. Recommended for tables.

Every time you change an option, the formula preview below updates instantly.

### Step 4 — Copy Formula

At the bottom of the Function Builder, you'll see the **live formula preview** — a code box that shows the exact Google Sheets formula based on your current selections.

**Example preview:**
```
=DIVIDENDDATA_DIVIDENDS("AAPL", "summary", TRUE)
```

**To use the formula:**
1. Click the formula preview box or the **Copy Formula** button to copy it to your clipboard.
2. Click any cell in your spreadsheet.
3. Paste (`Ctrl+V` or `Cmd+V`).
4. Press **Enter** — the data populates.

> The formula preview turns green briefly to confirm it was copied.

### Function Builder Tips

- **Explore before you build.** Change the category dropdown to browse what's available — each category reveals different metrics and options.
- **Use headers for tables.** When pulling summaries, history, or statements, check **Show Headers** so you know what each column represents.
- **Batch for watchlists.** Enable batch mode and enter your watchlist tickers to build a comparison table with a single formula.
- **Learn the syntax.** The Function Builder is a great way to learn formula syntax. Once you're comfortable, you can type formulas directly without opening the builder.
- **Crypto & Commodities.** For these categories, the asset dropdown changes to show available symbols (e.g., `BTCUSD` for Bitcoin, `GCUSD` for Gold). You can also use the `list` metric to see all available symbols.


---

## Function Reference

_Below, you'll find detailed documentation for each function:_

## 1) DIVIDENDDATA_DIVIDENDS
### Description
Retrieves various dividend-related data for a single stock symbol. This function is useful for analyzing a company's dividend payments, yields, growth trends, and sustainability through payout ratios. It can return single values for quick metrics or tables for historical data.

> _**Backwards compatible:** `=DIVIDENDDATA(...)` still works as an alias for this function._

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The dividend metric to retrieve (default: `"fwd_payout"`). 

_Available metrics:_ 
- `"fwd_payout"` (forward annual payout)
- `"ttm_payout"` (TTM payout)
- `"fwd_yield"` (forward dividend yield)
- `"ttm_yield"` (TTM dividend yield)
- `"next_dividend"` (Next declared payout amount)
- `"next_dividend_full"` (Table: Amount, Decl Date, Ex Date, Pay Date, Frequency)
- `"frequency"` (payment frequency)
- `"ex_div_date"` (most recent ex-dividend date)
- `"payment_date"` (most recent payment date)
- `"declaration_date"` (most recent declaration date)
- `"history"` (historical table)
- `"summary"` (summary table)
- `"growth"` (growth rates table)
- `"1y_cagr"` (1-year CAGR)
- `"3y_cagr"` (3-year CAGR)
- `"5y_cagr"` (5-year CAGR)
- `"10y_cagr"` (10-year CAGR)
- `"payout_ratio"` (EPS payout ratio)
- `"fcf_payout_ratio"` (FCF payout ratio).

**showHeaders**: Boolean to include headers in history or growth tables (default: false).

### Examples:

**Get Microsoft's forward dividend yield:**
```
=DIVIDENDDATA_DIVIDENDS("MSFT", "fwd_yield")
```
→ Returns: `0.008` (0.8%)

**Get Microsoft's most recent ex-dividend date:**
```
=DIVIDENDDATA_DIVIDENDS("MSFT", "ex_div_date")
```
→ Returns: `2025-11-19`

**Get Microsoft's full dividend history with headers:**
```
=DIVIDENDDATA_DIVIDENDS("MSFT", "history", TRUE)
```
→ Returns table:

| Declaration Date  | Record Date | Payment Date  | Adjusted Dividend | Dividend  | Growth Rate | Frequency |
| ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- |
| 9/14/2025 | 11/19/2025  | 12/10/2025  | $0.91  | $0.91  | 0.096 (9.6%)  | Quarterly |
| 6/9/2025 | 8/20/2025 | 9/10/2025  | $0.83  | $0.83 | 0  | Quarterly |
| 3/10/2025  | 5/14/2025  | 6/11/2025  | $0.83  | $0.83  | 0  | Quarterly |
| 12/2/2024 | 2/19/2025  | 3/12/2025  | $0.83  | $0.83  | 0 | Quarterly  |

_(The data above is formatted. In reality, it will return the raw numbers. You can choose how to format within the returned cells.)_

**Get Procter & Gamble's dividend summary:**
```
=DIVIDENDDATA_DIVIDENDS("PG", "summary")
```

| Annual Dividend (FWD)	| Dividend Yield (FWD)	| Annual Dividend (TTM)	| Dividend Yield (TTM)	| Payment Frequency	| Next Dividend | Declaration Date | Ex-Div Date | Payment Date | 1 Year CAGR |	3 Year CAGR |	5 Year CAGR |	10 Year CAGR |	Payout Ratio (Net Income) | Payout Ratio (Free Cash Flow) |
| ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- | ------------- |
| $3.64 |	0.70% |	$3.32 |	0.64% |	Quarterly |	0.91 |	2025-12-01 |	2026-02-18 |	2026-03-11 |	9.64% |	10.1% |	10.1% |	9.7% |	23.64% |	33.62% |

_(The data above is formatted. In reality, it will return the raw numbers. You can choose how to format within the returned cells.)_


## 2) DIVIDENDDATA_DIVIDENDS_BATCH
### Description
Retrieves batch dividend data for multiple stock symbols. This function is efficient for fetching dividend information across several tickers at once, such as latest payouts or historical tables. It supports both latest values and full history.

> _**Backwards compatible:** `=DIVIDENDDATA_BATCH(...)` still works as an alias for this function._

### Parameters

**symbols**: Comma-separated stock ticker symbols (e.g., `"MSFT,KMB,O"`). Required.

**metric**: Select a metric (default: `"fwd_payout"`). Or select multiple with comma-separated metrics _(Aside from `"all"` and `"history"`)_. 

- `"all"` Get all the latest dividend metrics for the provided tickers.
- `"history"` Get the dividend history of all provided tickers.

_Available metrics:_ 
- `"fwd_payout"` (forward annual payout)
- `"adjdividend"` (most recent declared adjusted dividend)
- `"dividend"` (most recent declared dividend)
- `"record_date"` (or `"recorddate"`) (most recent record date)
- `"declaration_date"` (or `"declarationdate"`) (most recent declaration date)
- `"payment_date"` (or `"paymentdate"`) (most recent payment date)
- `"yield"` (dividend yield)
- `"frequency"` (payment frequency)

**showHeaders**: Boolean to include header row and symbol column (defaults to true for `"all"` or `"history"`).

### Examples

**Get forward payouts and yields for a dividend portfolio:**
```
=DIVIDENDDATA_DIVIDENDS_BATCH("MSFT,AAPL", "fwd_payout,yield", TRUE)
```
→ Returns table with symbols, forward payouts, and yields.

**Get combined dividend history for multiple stocks:**
```
=DIVIDENDDATA_DIVIDENDS_BATCH("MSFT,KMB", "history")
```
→ Returns flat historical dividend table for all symbols.

## 3) DIVIDENDDATA_STATEMENT
### Description
Retrieves full financial statements for a stock. This function is useful for in-depth financial analysis, providing complete income statements, balance sheets, or cash flow statements over time. It supports filtering by period and year.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The statement type: `"income"`, `"balance"`, or `"cash_flow"`. Required.

**showHeaders**: Boolean to include header row (default: `false`).

**period**: Period filter: `"FY"`, `"Q1"`, `"Q2"`, `"Q3"`, `"Q4"`, `"annual"`, `"quarter"`, `"ttm" (default: `''`).

**year**: Specific year to filter (e.g., `2025`, default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_STATEMENT("MSFT", "income", TRUE, "annual", "2024")`

**Output**: Income statement table for 2024 with headers.

**Input**: `=DIVIDENDDATA_STATEMENT("AAPL", "cash_flow", FALSE, "ttm")`

**Output**: TTM cash flow statement without headers.

## 4) DIVIDENDDATA_METRICS
### Description
Retrieves a specific metric from financial statements. This function allows targeted extraction of key financial figures like revenue or free cash flow, either as the latest value or historical table.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The specific financial metric (e.g., `"revenue"`, `"freeCashFlow"`). 

_All available metrics_: 
* `"Revenue"` (Revenue)
* `"CostOfRevenue"` (Cost of Revenue)
* `"GrossProfit"` (Gross Profit)
* `"ResearchAndDevelopmentExpenses"` (Research and Development Expenses)
* `"GeneralAndAdministrativeExpenses"` (General and Administrative Expenses)
* `"SellingAndMarketingExpenses"` (Selling and Marketing Expenses)
* `"SellingGeneralAndAdministrativeExpenses"` (Selling General and Administrative Expenses)
* `"OtherExpenses"` (Other Expenses)
* `"OperatingExpenses"` (Operating Expenses)
* `"CostAndExpenses"` (Cost and Expenses)
* `"NetInterestIncome"` (Net Interest Income)
* `"InterestIncome"` (Interest Income)
* `"InterestExpense"` (Interest Expense)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"Ebitda"` (Ebitda)
* `"Ebit"` (Ebit)
* `"NonOperatingIncomeExcludingInterest"` (Non Operating Income Excluding Interest)
* `"OperatingIncome"` (Operating Income)
* `"TotalOtherIncomeExpensesNet"` (Total Other Income Expenses Net)
* `"IncomeBeforeTax"` (Income Before Tax)
* `"IncomeTaxExpense"` (Income Tax Expense)
* `"NetIncomeFromContinuingOperations"` (Net Income From Continuing Operations)
* `"NetIncomeFromDiscontinuedOperations"` (Net Income From Discontinued Operations)
* `"OtherAdjustmentsToNetIncome"` (Other Adjustments To Net Income)
* `"NetIncome"` (Net Income)
* `"NetIncomeDeductions"` (Net Income Deductions)
* `"BottomLineNetIncome"` (Bottom Line Net Income)
* `"Eps"` (Eps)
* `"EpsDiluted"` (Eps Diluted)
* `"WeightedAverageShsOut"` (Weighted Average Shs Out)
* `"WeightedAverageShsOutDil"` (Weighted Average Shs Out Dil)
* `"CashAndCashEquivalents"` (Cash and Cash Equivalents)
* `"ShortTermInvestments"` (Short Term Investments)
* `"CashAndShortTermInvestments"` (Cash and Short Term Investments)
* `"NetReceivables"` (Net Receivables)
* `"AccountsReceivables"` (Accounts Receivables)
* `"OtherReceivables"` (Other Receivables)
* `"Inventory"` (Inventory)
* `"Prepaids"` (Prepaids)
* `"OtherCurrentAssets"` (Other Current Assets)
* `"TotalCurrentAssets"` (Total Current Assets)
* `"PropertyPlantEquipmentNet"` (Property Plant Equipment Net)
* `"Goodwill"` (Goodwill)
* `"IntangibleAssets"` (Intangible Assets)
* `"GoodwillAndIntangibleAssets"` (Goodwill and Intangible Assets)
* `"LongTermInvestments"` (Long Term Investments)
* `"TaxAssets"` (Tax Assets)
* `"OtherNonCurrentAssets"` (Other Non Current Assets)
* `"TotalNonCurrentAssets"` (Total Non Current Assets)
* `"OtherAssets"` (Other Assets)
* `"TotalAssets"` (Total Assets)
* `"TotalPayables"` (Total Payables)
* `"AccountPayables"` (Account Payables)
* `"OtherPayables"` (Other Payables)
* `"AccruedExpenses"` (Accrued Expenses)
* `"ShortTermDebt"` (Short Term Debt)
* `"CapitalLeaseObligationsCurrent"` (Capital Lease Obligations Current)
* `"TaxPayables"` (Tax Payables)
* `"DeferredRevenue"` (Deferred Revenue)
* `"OtherCurrentLiabilities"` (Other Current Liabilities)
* `"TotalCurrentLiabilities"` (Total Current Liabilities)
* `"LongTermDebt"` (Long Term Debt)
* `"CapitalLeaseObligationsNonCurrent"` (Capital Lease Obligations Non Current)
* `"DeferredRevenueNonCurrent"` (Deferred Revenue Non Current)
* `"DeferredTaxLiabilitiesNonCurrent"` (Deferred Tax Liabilities Non Current)
* `"OtherNonCurrentLiabilities"` (Other Non Current Liabilities)
* `"TotalNonCurrentLiabilities"` (Total Non Current Liabilities)
* `"OtherLiabilities"` (Other Liabilities)
* `"CapitalLeaseObligations"` (Capital Lease Obligations)
* `"TotalLiabilities"` (Total Liabilities)
* `"TreasuryStock"` (Treasury Stock)
* `"PreferredStock"` (Preferred Stock)
* `"CommonStock"` (Common Stock)
* `"RetainedEarnings"` (Retained Earnings)
* `"AdditionalPaidInCapital"` (Additional Paid In Capital)
* `"AccumulatedOtherComprehensiveIncomeLoss"` (Accumulated Other Comprehensive Income Loss)
* `"OtherTotalStockholdersEquity"` (Other Total Stockholders Equity)
* `"TotalStockholdersEquity"` (Total Stockholders Equity)
* `"TotalEquity"` (Total Equity)
* `"MinorityInterest"` (Minority Interest)
* `"TotalLiabilitiesAndTotalEquity"` (Total Liabilities and Total Equity)
* `"TotalInvestments"` (Total Investments)
* `"TotalDebt"` (Total Debt)
* `"NetDebt"` (Net Debt)
* `"NetIncome"` (Net Income)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"DeferredIncomeTax"` (Deferred Income Tax)
* `"StockBasedCompensation"` (Stock Based Compensation)
* `"ChangeInWorkingCapital"` (Change In Working Capital)
* `"AccountsReceivables"` (Accounts Receivables)
* `"Inventory"` (Inventory)
* `"AccountsPayables"` (Accounts Payables)
* `"OtherWorkingCapital"` (Other Working Capital)
* `"OtherNonCashItems"` (Other Non Cash Items)
* `"NetCashProvidedByOperatingActivities"` (Net Cash Provided By Operating Activities)
* `"InvestmentsInPropertyPlantAndEquipment"` (Investments In Property Plant and Equipment)
* `"AcquisitionsNet"` (Acquisitions Net)
* `"PurchasesOfInvestments"` (Purchases of Investments)
* `"SalesMaturitiesOfInvestments"` (Sales Maturities of Investments)
* `"OtherInvestingActivities"` (Other Investing Activities)
* `"NetCashProvidedByInvestingActivities"` (Net Cash Provided By Investing Activities)
* `"NetDebtIssuance"` (Net Debt Issuance)
* `"LongTermNetDebtIssuance"` (Long Term Net Debt Issuance)
* `"ShortTermNetDebtIssuance"` (Short Term Net Debt Issuance)
* `"NetStockIssuance"` (Net Stock Issuance)
* `"NetCommonStockIssuance"` (Net Common Stock Issuance)
* `"CommonStockIssuance"` (Common Stock Issuance)
* `"CommonStockRepurchased"` (Common Stock Repurchased)
* `"NetPreferredStockIssuance"` (Net Preferred Stock Issuance)
* `"NetDividendsPaid"` (Net Dividends Paid)
* `"CommonDividendsPaid"` (Common Dividends Paid)
* `"PreferredDividendsPaid"` (Preferred Dividends Paid)
* `"OtherFinancingActivities"` (Other Financing Activities)
* `"NetCashProvidedByFinancingActivities"` (Net Cash Provided By Financing Activities)
* `"EffectOfForexChangesOnCash"` (Effect of Forex Changes On Cash)
* `"NetChangeInCash"` (Net Change In Cash)
* `"CashAtEndOfPeriod"` (Cash At End of Period)
* `"CashAtBeginningOfPeriod"` (Cash At Beginning of Period)
* `"OperatingCashFlow"` (Operating Cash Flow)
* `"CapitalExpenditure"` (Capital Expenditure)
* `"FreeCashFlow"` (Free Cash Flow)
* `"IncomeTaxesPaid"` (Income Taxes Paid)
* `"InterestPaid"` (Interest Paid)

**showHeaders**: If `true`, returns historical table; else, latest value (default: `false`).
**period**: Period: `"annual"`, `"quarter"`, `"ttm"` (default: `''`).
**year**: Specific year filter (default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_METRICS("MSFT", "revenue")`

**Output**: Latest revenue value for Microsoft stock. Annual by default.

**Input**: `=DIVIDENDDATA_METRICS("AAPL", "freeCashFlow", TRUE, "quarter", "2024")`

**Output**: Quarterly free cash flow history table for 2024 of Apple stock.

## 5) DIVIDENDDATA_RATIOS
### Description
Retrieves a specific financial ratio or key metric for a stock. Useful for valuation, liquidity, and efficiency analysis, either latest value or historical.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The ratio or key metric. (e.g., `"currentRatio"`, `"peRatio"`).

_All available metrics:_
* `"RevenuePerShare"` (Revenue Per Share)
* `"NetIncomePerShare"` (Net Income Per Share)
* `"OperatingCashFlowPerShare"` (Operating Cash Flow Per Share)
* `"FreeCashFlowPerShare"` (Free Cash Flow Per Share)
* `"CashPerShare"` (Cash Per Share)
* `"BookValuePerShare"` (Book Value Per Share)
* `"TangibleBookValuePerShare"` (Tangible Book Value Per Share)
* `"ShareholdersEquityPerShare"` (Shareholders Equity Per Share)
* `"InterestDebtPerShare"` (Interest Debt Per Share)
* `"MarketCap"` (Market Cap)
* `"EnterpriseValue"` (Enterprise Value)
* `"PeRatio"` (Pe Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"Pocfratio"` (Price to Operating Cash Flow Ratio)
* `"PfcfRatio"` (Price to Free Cash Flow Ratio)
* `"PbRatio"` (Price Book Value Ratio)
* `"PtbRatio"` (Price to Book Ratio)
* `"EvToSales"` (Enterprise Value to Sales)
* `"EnterpriseValueOverEBITDA"` (Enterprise Value Over EBITDA)
* `"EvToOperatingCashFlow"` (Enterprise Value To Operating Cash Flow)
* `"EvToFreeCashFlow"` (Enterprise Value To Free Cash Flow)
* `"EarningsYield"` (Earnings Yield)
* `"FreeCashFlowYield"` (Free Cash Flow Yield)
* `"DebtToEquity"` (Debt To Equity)
* `"DebtToAssets"` (Debt To Assets)
* `"NetDebtToEBITDA"` (Net Debt To EBITDA)
* `"CurrentRatio"` (Current Ratio)
* `"InterestCoverage"` (Interest Coverage)
* `"IncomeQuality"` (Income Quality)
* `"DividendYield"` (Dividend Yield)
* `"PayoutRatio"` (Payout Ratio)
* `"SalesGeneralAndAdministrativeToRevenue"` (Sales General And Administrative To Revenue)
* `"ResearchAndDevelopmentToRevenue"` (Research And Development To Revenue)
* `"IntangiblesToTotalAssets"` (Intangibles To Total Assets)
* `"CapexToOperatingCashFlow"` (Capex To Operating Cash Flow)
* `"CapexToRevenue"` (Capex To Revenue)
* `"CapexToDepreciation"` (Capex To Depreciation)
* `"StockBasedCompensationToRevenue"` (Stock Based Compensation To Revenue)
* `"GrahamNumber"` (Graham Number)
* `"Roic"` (Return on Investing Capital)
* `"ReturnOnTangibleAssets"` (Return On Tangible Assets)
* `"GrahamNetNet"` (Graham Net Net)
* `"WorkingCapital"` (Working Capital)
* `"TangibleAssetValue"` (Tangible Asset Value)
* `"NetCurrentAssetValue"` (Net Current Asset Value)
* `"InvestedCapital"` (Invested Capital)
* `"AverageReceivables"` (Average Receivables)
* `"AveragePayables"` (Average Payables)
* `"AverageInventory"` (Average Inventory)
* `"DaysSalesOutstanding"` (Days Sales Outstanding)
* `"DaysPayablesOutstanding"` (Days Payables Outstanding)
* `"DaysOfInventoryOnHand"` (Days Of Inventory On Hand)
* `"ReceivablesTurnover"` (Receivables Turnover)
* `"PayablesTurnover"` (Payables Turnover)
* `"InventoryTurnover"` (Inventory Turnover)
* `"Roe"` (Return on Equity)
* `"CapexPerShare"` (Capex Per Share)
* `"CurrentRatio"` (Current Ratio)
* `"QuickRatio"` (Quick Ratio)
* `"CashRatio"` (Cash Ratio)
* `"DaysOfSalesOutstanding"` (Days Of Sales Outstanding)
* `"DaysOfInventoryOutstanding"` (Days Of Inventory Outstanding)
* `"OperatingCycle"` (Operating Cycle)
* `"DaysOfPayablesOutstanding"` (Days Of Payables Outstanding)
* `"CashConversionCycle"` (Cash Conversion Cycle)
* `"GrossProfitMargin"` (Gross Profit Margin)
* `"OperatingProfitMargin"` (Operating Profit Margin)
* `"PretaxProfitMargin"` (Pretax Profit Margin)
* `"NetProfitMargin"` (Net Profit Margin)
* `"EffectiveTaxRate"` (Effective Tax Rate)
* `"ReturnOnAssets"` (Return On Assets)
* `"ReturnOnEquity"` (Return On Equity)
* `"ReturnOnCapitalEmployed"` (Return On Capital Employed)
* `"NetIncomePerEBT"` (Net Income Per EBT)
* `"EbtPerEbit"` (Ebt Per Ebit)
* `"EbitPerRevenue"` (Ebit Per Revenue)
* `"DebtRatio"` (Debt Ratio)
* `"DebtEquityRatio"` (Debt Equity Ratio)
* `"LongTermDebtToCapitalization"` (Long Term Debt To Capitalization)
* `"TotalDebtToCapitalization"` (Total Debt To Capitalization)
* `"InterestCoverage"` (Interest Coverage)
* `"CashFlowCoverageRatios"` (Cash Flow Coverage Ratios)
* `"ShortTermCoverageRatios"` (Short Term Coverage Ratios)
* `"CapitalExpenditureCoverageRatio"` (Capital Expenditure Coverage Ratio)
* `"DividendPaidAndCapexCoverageRatio"` (Dividend Paid And Capex Coverage Ratio)
* `"DividendPayoutRatio"` (Dividend Payout Ratio)
* `"PriceBookValueRatio"` (Price Book Value Ratio)
* `"PriceToBookRatio"` (Price To Book Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"PriceEarningsRatio"` (Price Earnings Ratio)
* `"PriceToFreeCashFlowsRatio"` (Price To Free Cash Flows Ratio)
* `"PriceToOperatingCashFlowsRatio"` (Price To Operating Cash Flows Ratio)
* `"PriceCashFlowRatio"` (Price Cash Flow Ratio)
* `"PriceEarningsToGrowthRatio"` (Price Earnings To Growth Ratio)
* `"PriceSalesRatio"` (Price Sales Ratio)
* `"DividendYield"` (Dividend Yield)
* `"EnterpriseValueMultiple"` (Enterprise Value Multiple)
* `"PriceFairValue"` (Price Fair Value)

**showHeaders**: If `true`, returns historical table (default: `false`).

**period**: Period: `"annual"`, `"quarter"`, `"ttm"` (default: `''`).

**year**: Specific year filter like `2023` (default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_RATIOS("MSFT", "peRatio")`

**Output**: Latest P/E ratio for Microsoft.

**Input**: `=DIVIDENDDATA_RATIOS("AAPL", "currentRatio", TRUE, "annual")`

**Output**: Annual current ratio history table for Apple.

## 6) DIVIDENDDATA_GROWTH
### Description
Retrieves a specific growth metric for financial figures. Useful for trend analysis, like revenue or EPS growth rates over time.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The growth metric (e.g., `"revenueGrowth"`, `"epsGrowth"`). 

_All available metrics:_
* `"RevenueGrowth"` (Revenue Growth)
* `"GrossProfitGrowth"` (Gross Profit Growth)
* `"Ebitgrowth"` (Ebitgrowth)
* `"OperatingIncomeGrowth"` (Operating Income Growth)
* `"NetIncomeGrowth"` (Net Income Growth)
* `"Epsgrowth"` (GAAP EPS growth)
* `"EpsdilutedGrowth"` (EPS diluted Growth)
* `"WeightedAverageSharesGrowth"` (Weighted Average Shares Growth)
* `"WeightedAverageSharesDilutedGrowth"` (Weighted Average Shares Diluted Growth)
* `"DividendsPerShareGrowth"` (Dividends Per Share Growth)
* `"OperatingCashFlowGrowth"` (Operating Cash Flow Growth)
* `"ReceivablesGrowth"` (Receivables Growth)
* `"InventoryGrowth"` (Inventory Growth)
* `"AssetGrowth"` (Asset Growth)
* `"BookValuePerShareGrowth"` (Book Value Per Share Growth)
* `"DebtGrowth"` (Debt Growth)
* `"RdexpenseGrowth"` (Rdexpense Growth)
* `"SgaexpensesGrowth"` (Sgaexpenses Growth)
* `"FreeCashFlowGrowth"` (Free Cash Flow Growth)
* `"TenYRevenueGrowthPerShare"` (Ten Y Revenue Growth Per Share)
* `"FiveYRevenueGrowthPerShare"` (Five Y Revenue Growth Per Share)
* `"ThreeYRevenueGrowthPerShare"` (Three Y Revenue Growth Per Share)
* `"TenYOperatingCFGrowthPerShare"` (Ten Y Operating Cash Flow Growth Per Share)
* `"FiveYOperatingCFGrowthPerShare"` (Five Y Operating Cash Flow Growth Per Share)
* `"ThreeYOperatingCFGrowthPerShare"` (Three Y Operating Cash Flow Growth Per Share)
* `"TenYNetIncomeGrowthPerShare"` (Ten Y Net Income Growth Per Share)
* `"FiveYNetIncomeGrowthPerShare"` (Five Y Net Income Growth Per Share)
* `"ThreeYNetIncomeGrowthPerShare"` (Three Y Net Income Growth Per Share)
* `"TenYShareholdersEquityGrowthPerShare"` (Ten Y Shareholders Equity Growth Per Share)
* `"FiveYShareholdersEquityGrowthPerShare"` (Five Y Shareholders Equity Growth Per Share)
* `"ThreeYShareholdersEquityGrowthPerShare"` (Three Y Shareholders Equity Growth Per Share)
* `"TenYDividendPerShareGrowthPerShare"` (Ten Y Dividend Per Share Growth Per Share)
* `"FiveYDividendPerShareGrowthPerShare"` (Five Y Dividend Per Share Growth Per Share)
* `"ThreeYDividendPerShareGrowthPerShare"` (Three Y Dividend Per Share Growth Per Share)
* `"EbitdaGrowth"` (Ebitda Growth)
* `"GrowthCapitalExpenditure"` (Growth Capital Expenditure)
* `"TenYBottomLineNetIncomeGrowthPerShare"` (Ten Y Bottom Line Net Income Growth Per Share)
* `"FiveYBottomLineNetIncomeGrowthPerShare"` (Five Y Bottom Line Net Income Growth Per Share)
* `"ThreeYBottomLineNetIncomeGrowthPerShare"` (Three Y Bottom Line Net Income Growth Per Share)

**showHeaders**: If `true`, returns historical table (default: `false`).

**period**: Period: `"annual"`, `"quarter"`, `"q1"`, `"q2"`, `"q3"`, `"q4"`, `"fy"` (default: `''`).

**year**: Specific year filter like `2023` (default: `''`).

### Examples

**Input**: `=DIVIDENDDATA_GROWTH("MSFT", "revenueGrowth")`

**Output**: Latest revenue growth rate.

**Input**: `=DIVIDENDDATA_GROWTH("AAPL", "epsGrowth", TRUE, "quarter")`

**Output**: Quarterly EPS growth history table.

## 7) DIVIDENDDATA_QUOTE
### Description
Retrieves stock quote data, including current price, changes, or historical prices. Useful for real-time monitoring or historical analysis of stock performance.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"AAPL"`). Required.

**metric**: The quote metrics to retrieve. (default: `"price"`).

_Available metrics:_
- `"price"` (Current Price)
- `"change"` (Price Change)
- `"volume"` (Volume)
- `"priceAvg50"` (50 Day Average Price)
- `"priceAvg200"` (200 Day Average Price)
- `"yearHigh"` (52 Week High)
- `"yearLow"` (52 Week Low)
- `"full"` (Returns all quote data points in a table)
- `"history"` (Returns historical price data)

**fromDate**: Start date for history (`"YYYY-MM-DD"`, for history only).

**toDate**: End date for history (`"YYYY-MM-DD"`, for history only).

**showHeaders**: Include headers for `"full"` or `"history"` (default: `true`).

### Examples

**Input**: `=DIVIDENDDATA_QUOTE("AAPL", "price")`

**Output**: Current price.

**Input**: `=DIVIDENDDATA_QUOTE("MSFT", "priceAvg50")`

**Output**: The 50-day moving average price.

**Input**: `=DIVIDENDDATA_QUOTE("MSFT", "history", "2024-01-01", "2024-12-31", TRUE)`

**Output**: Historical price table with headers.

## 8) DIVIDENDDATA_QUOTE_BATCH
### Description
Retrieves batch quote data for multiple stocks. Efficient for monitoring prices or volumes across tickers.

### Parameters

**symbols**: Comma-separated tickers (e.g., `"AAPL,MSFT"`). Required.

**metrics**: Comma-separated metrics or `"all"`: `"price"`, `"change"`, `"volume"` (default: `"all"`).

**showHeaders**: Include headers and symbol column (defaults based on metrics).

### Examples

**Input**: `=DIVIDENDDATA_QUOTE_BATCH("AAPL,MSFT", "price,change")`

**Output**: Table with prices and changes.

**Input**: `=DIVIDENDDATA_QUOTE_BATCH("MSFT,KMB", "all", TRUE)`

**Output**: Full quotes table with headers.


## 9) DIVIDENDDATA_PROFILE
### Description
Retrieves company profile information. Useful for overview details like market cap, sector, or description.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: Specific profile metric or `"full"`.

_All available metrics:_
* `"symbol"` (Symbol)
* `"price"` (Price)
* `"marketcap"` (Market Cap)
* `"beta"` (Beta)
* `"lastdividend"` (Last Dividend)
* `"range"` (Range)
* `"change"` (Change)
* `"changepercentage"` (Change Percentage)
* `"volume"` (Volume)
* `"averagevolume"` (Average Volume)
* `"companyname"` (Company Name)
* `"currency"` (Currency)
* `"cik"` (Cik)
* `"isin"` (Isin)
* `"cusip"` (Cusip)
* `"exchangefullname"` (Exchange Full Name)
* `"exchange"` (Exchange)
* `"industry"` (Industry)
* `"website"` (Website)
* `"description"` (Description)
* `"ceo"` (Ceo)
* `"sector"` (Sector)
* `"country"` (Country)
* `"fulltimeemployees"` (Full Time Employees)
* `"phone"` (Phone)
* `"address"` (Address)
* `"city"` (City)
* `"state"` (State)
* `"zip"` (Zip)
* `"image"` (Image)
* `"ipodate"` (Ipo Date)
* `"defaultimage"` (Default Image)
* `"isetf"` (Is Etf)
* `"isactivelytrading"` (Is Actively Trading)
* `"isadr"` (Is ADR?)
* `"isfund"` (Is Fund?)
* `"full"` (Return Full Profile)

**showHeaders**: Include headers for `"full"` (default: `false`).

### Examples

**Input**: `=DIVIDENDDATA_PROFILE("MSFT", "marketcap")`

**Output**: Market capitalization.

**Input**: `=DIVIDENDDATA_PROFILE("AAPL", "full", TRUE)`

**Output**: Full profile table with headers.

## 10) DIVIDENDDATA_FUND
### Description
Retrieves data for ETFs or mutual funds. Useful for fund analysis, including holdings, expense ratios, or sector exposures.

### Parameters

**symbol**: Fund ticker symbol (e.g., `"SPY"`). Required.

**metric**: Fund metrics like `"holdings"`, `"expenseRatio"`, etc.

_All available metrics:_
* `"etfcompany"` (Etf Company)
* `"expenseratio"` (Expense Ratio)
* `"assetsundermanagement"` (Assets Under Management)
* `"holdings"` (Holdings)
* `"nav"` (Net Asset Value)
* `"navcurrency"` (NAV Currency)
* `"holdingscount"` (Holdings Count)
* `"countryweighting"` (Country Weighting)
* `"sectorslist"` (Sectors List)
* `"description"` (Description)
* `"assetclass"` (Asset Class)
* `"domicile"` (Domicile)
* `"website"` (Website)
* `"avgvolume"` (Avg Volume)
* `"inceptiondate"` (Inception Date)
* `"updatedat"` (Updated At)

**showHeaders**: Include headers for tables (default: `false`).

### Examples

**Input**: `=DIVIDENDDATA_FUND("SPY", "expenseratio")`

**Output**: Expense ratio.

**Input**: `=DIVIDENDDATA_FUND("SPY", "holdings", TRUE)`

**Output**: Holdings table with headers.

## 11) DIVIDENDDATA_SEGMENTS
### Description
Retrieves revenue segmentation data by product or geography. Useful for understanding revenue sources and diversification.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"AAPL"`). Required.

**metric**: Segmentation type: `"products"` or `"geographic"`. Required.

**period**: Period: `"annual"` or `"quarter"` (default: `"annual"`).

**year**: Specific year filter (default: `''`).

**showHeaders**: Include header row (default: `false`).

### Examples

**Input**: `=DIVIDENDDATA_SEGMENTS("AAPL", "products", "annual", "2024", TRUE)`

**Output**: Product segments table for 2024 with headers.

**Input**: `=DIVIDENDDATA_SEGMENTS("MSFT", "geographic", "quarter")`

**Output**: Quarterly geographic segments table.

## 12) DIVIDENDDATA_KPIS
### Description
Retrieves key performance indicators (KPIs) and segments from Fiscal.ai. Useful for advanced metrics like customer acquisition cost or churn rates.

### Parameters

**ticker**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**period**: `"annual"` or `"quarterly"` (default: `"annual"`).

**year**: Optional year filter (default: `''`).

**showheaders**: Include header row (default: `TRUE`).

### Examples

**Input**: `=DIVIDENDDATA_KPIS("MSFT", "annual", "2024", TRUE)`

**Output**: Annual KPIs table for 2024 with headers.

**Input**: `=DIVIDENDDATA_KPIS("AAPL", "quarterly")`

**Output**: Quarterly KPIs table without headers.

## 13) DIVIDENDDATA_COMMODITIES
### Description
Retrieves commodities data, such as prices or history. Useful for tracking commodity markets like oil or gold.

### Parameters

**symbol**: Commodity symbol (e.g., `"CLUSD"`). Required except for `"list"`.

_Full list of symbols:_
* `"ZQUSD"` (30 Day Fed Fund Futures)
* `"ALIUSD"` (Aluminum Futures)
* `"ZBUSD"` (30 Year U.S. Treasury Bond Futures)
* `"ZOUSX"` (Oat Futures)
* `"ESUSD"` (E-Mini S&P 500)
* `"ZMUSD"` (Soybean Meal Futures)
* `"GCUSD"` (Gold Futures)
* `"ZLUSX"` (Soybean Oil Futures)
* `"KEUSX"` (Wheat Futures)
* `"ZFUSD"` (Five-Year US Treasury Note)
* `"SILUSD"` (Micro Silver Futures)
* `"ZCUSX"` (Corn Futures)
* `"HEUSX"` (Lean Hogs Futures)
* `"PLUSD"` (Platinum)
* `"HGUSD"` (Copper)
* `"MGCUSD"` (Micro Gold Futures)
* `"SBUSX"` (Sugar)
* `"SIUSD"` (Silver Futures)
* `"CTUSX"` (Cotton)
* `"DXUSD"` (US Dollar)
* `"ZSUSX"` (Soybean Futures)
* `"LBUSD"` (Lumber Futures)
* `"LEUSX"` (Live Cattle Futures)
* `"NGUSD"` (Natural Gas)
* `"CLUSD"` (Crude Oil)
* `"OJUSX"` (Orange Juice)
* `"KCUSX"` (Coffee)
* `"PAUSD"` (Palladium)
* `"GFUSX"` (Feeder Cattle Futures)
* `"ZTUSD"` (2-Year T-Note Futures)
* `"ZRUSD"` (Rough Rice Futures)
* `"CCUSD"` (Cocoa)
* `"NQUSD"` (Nasdaq 100)
* `"ZNUSD"` (10-Year T-Note Futures)
* `"RTYUSD"` (Micro E-mini Russell 2000 Index Futures)
* `"BZUSD"` (Brent Crude Oil)
* `"DCUSD"` (Class III Milk Futures)
* `"YMUSD"` (Mini Dow Jones Industrial Average Index)
* `"RBUSD"` (Gasoline RBOB)
* `"HOUSD"` (Heating Oil)

**metric**: Metric: `"list"`, `"price"`, `"fullquote"`, `"history"`. Required.

**fromDate**: Start date for history.

**toDate**: End date for history.

**showHeaders**: Include headers for tables.

### Examples

**Input**: `=DIVIDENDDATA_COMMODITIES("CLUSD", "price")`

**Output**: Current oil price.

**Input**: `=DIVIDENDDATA_COMMODITIES(, "list", , , TRUE)`

**Output**: Commodities list table with headers.

## 14) DIVIDENDDATA_CRYPTO
### Description
Retrieves cryptocurrency data, such as prices or history. Useful for tracking crypto markets like Bitcoin or Ethereum.

### Parameters
**symbol**: Cryptocurrency symbol (e.g., "BTCUSD"). Required except for "list".

_Full list of symbols (examples, may vary):_
* `"BTCUSD"` (Bitcoin)
* `"ETHUSD"` (Ethereum)
* `"BNBUSD"` (Binance Coin)
* `"XRPUSD"` (Ripple)
* `"ADAUSD"` (Cardano)
* `"SOLUSD"` (Solana)
* `"DOGEUSD"` (Dogecoin)
* `"LINKUSD"` (Chainlink)

**metric**: Metric: `"list"`, `"price"`, `"fullquote"`, `"history"`. Required.

**fromDate**: Start date for history (`"YYYY-MM-DD"`).

**toDate**: End date for history (`"YYYY-MM-DD"`).

**showHeaders**: Include headers for tables (default: `true`).

### Examples:

**Input**: `=DIVIDENDDATA_CRYPTO("BTCUSD", "price")`

**Output**: Current Bitcoin price.

**Input**: `=DIVIDENDDATA_CRYPTO(, "list", , , TRUE)`

**Output**: Cryptocurrencies list table with headers.

**Input**: `=DIVIDENDDATA_CRYPTO("ETHUSD", "history", "2024-01-01", "2024-12-31", TRUE)`

**Output**: Historical Ethereum price table with headers.

## 15) DIVIDENDDATA_PRICE_TARGET
### Description
Retrieves analyst price targets for a stock. Returns summary metrics, consensus values, or news list. Useful for gauging market expectations and analyst opinions.

### Parameters
**symbol**: Stock ticker symbol (e.g., `"AAPL"`). Required.

**metric**: Metric: `"summary"`, `"high"`, `"low"`, `"consensus"`, `"median"`, `"news"` (default: `"summary"`).

**showHeaders**: Include headers for summary or news tables (defaults to true for summary/news, false otherwise).

### Examples:

**Input**: `=DIVIDENDDATA_PRICE_TARGET("AAPL", "consensus")`

**Output**: Consensus price target.

**Input**: `=DIVIDENDDATA_PRICE_TARGET("MSFT", "summary", TRUE)`

**Output**: Price target summary table with headers.

**Input**: `=DIVIDENDDATA_PRICE_TARGET("GOOGL", "news", TRUE)`

**Output**: Recent price target news table with headers.

## 16) DIVIDENDDATA_ESTIMATES
### Description
Retrieves analyst estimates for earnings and revenue. Returns a summary table of future estimates, specific metric history (estimates vs actuals), or single next-year values for quick watchlist building. Useful for forecasting future growth, valuation models, and gauging analyst expectations.

### Parameters
**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The estimate metric to retrieve (default: `"summary"`).

_Available metrics:_
* `"summary"` (Returns a full table including: Date, Avg EPS Estimate, EPS Growth, High/Low EPS Estimates, Avg Revenue Estimate, Revenue Growth, High/Low Revenue Estimates, and Analyst Counts).
* `"eps"` (Returns: Date, Avg EPS Estimate, EPS Growth).
* `"revenue"` (Returns: Date, Avg Revenue Estimate, Revenue Growth).
* `"next_year_eps"` (Returns a single value for the nearest future year's estimated EPS).
* `"next_year_revenue"` (Returns a single value for the nearest future year's estimated Revenue).
* `"next_year_eps_growth"` (Returns a single value for next year's estimated EPS growth %).
* `"next_year_revenue_growth"` (Returns a single value for next year's estimated Revenue growth %).

**period**: Period: `"annual"` or `"quarter"` (default: `"annual"`).

**showHeaders**: Include headers for tables (default: `true`).

### Examples:

**Input**: `=DIVIDENDDATA_ESTIMATES("MSFT", "summary")`

**Output**: A comprehensive table of future analyst estimates for earnings and revenue.

**Input**: `=DIVIDENDDATA_ESTIMATES("AAPL", "next_year_eps")`

**Output**: A single value representing the average EPS estimate for the next fiscal year (e.g., `7.25`).

**Input**: `=DIVIDENDDATA_ESTIMATES("NVDA", "revenue", "quarter")`

**Output**: A table showing quarterly revenue estimates and growth rates.

---

## Support

- **Help Center:** [help.dividenddata.com](https://help.dividenddata.com)
- **Email:** support@dividenddata.com
- **YouTube Tutorials:** [@DividendData](https://www.youtube.com/@DividendData)

---

<p align="center">
  <strong>Built for dividend investors, by dividend investors.</strong>
</p>
