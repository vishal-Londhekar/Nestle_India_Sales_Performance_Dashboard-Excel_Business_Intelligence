# 📊 Nestle India Sales Performance Dashboard — Excel Business Intelligence Workbook

<p align="center">
  <img src="https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white" />
  <img src="https://img.shields.io/badge/Domain-FMCG%20%7C%20Sales%20Analytics-orange" />
  <img src="https://img.shields.io/badge/Workbook%20Sheets-3-blueviolet" />
  <img src="https://img.shields.io/badge/Transactions-10%2C000%20Records-blue" />
  <img src="https://img.shields.io/badge/Period-2013–2014-green" />
  <img src="https://img.shields.io/badge/Status-Production%20Ready-brightgreen" />
</p>

---

## 📌 Table of Contents

- [Business Problem](#-business-problem)
- [Project Objective](#-project-objective)
- [Workbook Overview](#-workbook-overview)
- [Dataset / Data Source](#-dataset--data-source)
- [KPIs & Metrics](#-kpis--metrics)
- [Dashboard / Workbook Walkthrough](#-dashboard--workbook-walkthrough)
- [Key Insights](#-key-insights)
- [Business Recommendations](#-business-recommendations)
- [Excel Features Used](#-excel-features-used)
- [How to Use](#-how-to-use)
- [Project Structure](#-project-structure)
- [Author](#-author)

---

## 🧩 Business Problem

### The FMCG Sales Visibility Challenge

Fast-Moving Consumer Goods (FMCG) companies like Nestle India operate through **large, distributed sales forces** spanning dozens of territories, cities, and states. When a 72-member field sales team generates 10,000+ transactions annually across 23 cities, 15 states, and 4 regions, the challenge is no longer collecting data — it is transforming that data into timely, decision-ready intelligence.

Three specific business problems drive the need for this analytics workbook:

**1. No Unified View of Performance vs. Targets**
Without a consolidated reporting layer, managers lack real-time visibility into whether the team is on track for revenue, units, and market coverage targets. Monthly target attainment is only known at month-end — too late to course-correct.

**2. Product Portfolio Imbalance Goes Undetected**
With 11 active SKUs, sales leaders cannot quickly identify which products are driving volume, which are underperforming, and where pricing strategy should shift — without a structured analytical tool.

**3. Geographic Expansion Gaps Hidden in Raw Data**
Nestle's target of 28-city coverage vs. 23 cities actually achieved represents a measurable market access gap. Without geographic analytics, which cities to prioritise and which states are under-penetrated cannot be determined with confidence.

### Who Benefits

| Stakeholder | Business Value Delivered |
|---|---|
| **National Sales Manager** | Single-view of total revenue vs. target with drill-down to region and city |
| **Regional Sales Managers** | Comparative performance by region and state to guide team resource allocation |
| **Product Marketing Managers** | Product-level volume and revenue rankings to guide promotional investment |
| **MIS / Reporting Executives** | Automated summary tables eliminating manual report compilation |
| **Business Analysts** | Ready-made KPI framework for periodic performance reviews |

---

## 🎯 Project Objective

This workbook delivers a **three-sheet, fully structured Excel business intelligence solution** built on 10,000 Nestle India sales transactions spanning 2013–2014. It transforms raw transaction-level data into executive-ready KPI dashboards and analytical summaries.

**Business Goals:**
- Track actual vs. target performance across three core KPIs: Total Revenue, Units Sold, and City Coverage
- Identify top-performing and underperforming products, states, and cities to guide sales resource allocation
- Provide a reusable reporting template that any MIS executive can refresh with updated raw data

**Analytical Goals:**
- Compute Total Sales from raw Units and Price data using formula-driven calculated column (`=Units × Price`)
- Build structured summary tables for Products by Units Sold, States by Sales, and Cities by Sales
- Deliver 6 embedded charts (Line, Doughnut, ChartEx modern charts) across the Dashboard sheet
- Apply conditional formatting to identify weekend transaction patterns for field activity analysis

---

## 📋 Workbook Overview

### Workbook Summary

| Property | Detail |
|---|---|
| **File** | `Sales_Data.xlsx` |
| **Sheets** | 3 (Sales Data, Final Working, DashBoard) |
| **Total Records** | 10,000 transaction rows |
| **Date Range** | January 1, 2013 – December 31, 2014 |
| **Product Portfolio** | 11 Nestle SKUs (confectionery & ice cream) |
| **Geography** | 15 Indian states, 23 cities, 4 regions |
| **Sales Team** | 72 Sales Representatives |
| **Charts** | 6 (1 Line chart, 3 Doughnut charts, 2 ChartEx modern charts) |
| **Excel Tables** | 4 structured Tables (`Table1`, `Table2`, `Table3`, `Table7`) |
| **Formulas** | 10,000 `Total Sales` calculated cells + KPI achievement formulas |
| **Navigation** | Internal hyperlinks between all three sheets |

### Three Sheets at a Glance

| Sheet | Purpose | Primary Audience |
|---|---|---|
| `Sales Data` | Raw 10,000-row transaction ledger with formula-computed `Total Sales` | MIS Executives / Data Analysts |
| `Final Working` | KPI scorecards, 3 summary tables, 4 embedded charts | Business Analysts / Sales Managers |
| `DashBoard` | Visual executive dashboard with 6 charts, KPI cards, branded design | Senior Leadership / Stakeholders |

### Navigation Design

The workbook uses **internal hyperlinks** — visible navigation buttons embedded as drawing objects on all three sheets — allowing non-technical stakeholders to move between the raw data, working analysis, and final dashboard without needing to locate sheet tabs. A support contact (`support@amazon.com`) is embedded for reporting queries.

---

## 📋 Dataset / Data Source

### Source Sheet: `Sales Data`

| Property | Detail |
|---|---|
| **Rows** | 10,000 transaction records |
| **Columns** | 10 fields (Excel Table `Table7`, columns C–L) |
| **Granularity** | One row = one sales transaction |
| **Date Range** | 2013-01-01 to 2014-12-31 |
| **Missing Values** | Zero — complete dataset across all 10 fields |

### Field Dictionary

| Column | Type | Description & Range / Values |
|---|---|---|
| `Date` | Date | Transaction date — daily granularity, 2013–2014 |
| `SalesRep` | String | Sales representative full name — 72 unique reps |
| `Product` | String | Nestle SKU name — 11 unique products |
| `Units` | Integer | Units sold per transaction — range 1–25, mean 3.4 |
| `Price` | Integer | Unit price in INR — range ₹100–₹500, mean ₹301 |
| `Total Sales` | Integer | **Formula-calculated:** `=Units × Price` — range ₹100–₹12,175 |
| `City` | String | Transaction city — 23 unique cities across India |
| `State` | String | Indian state — 15 unique states |
| `Region` | String | Geographic zone — East, West, North, South |
| `Day` | String | Day of week — Monday through Sunday |

### Product Portfolio (All 11 SKUs)

| Rank | Product Name | Units Sold | Total Sales (₹) |
|---|---|---|---|
| 1 | Nestle Smarties Pop-Up | 5,267 | 15,88,831 |
| 2 | Nestle Fab | 4,375 | 13,79,740 |
| 3 | Nestle Aero Mint Potz | 4,262 | 12,52,203 |
| 4 | Nestle Milky Bar Stick | 4,153 | 12,93,337 |
| 5 | Nestle Toffee Crumble | 3,586 | 10,72,377 |
| 6 | Nestle Rowntree Fruit Pastil | 3,258 | 9,79,502 |
| 7 | Nestle Rolo Potz | 2,498 | 7,19,702 |
| 8 | Nestle Maxibon Cookie | 2,152 | 6,38,340 |
| 9 | Nestle Kit Kat Cone | 2,101 | 6,31,179 |
| 10 | Nestle Nobbly Bobbly | 1,253 | 3,86,345 |
| 11 | Nestle Rowntree Screamers | 969 | 2,85,069 |

---

## 📐 KPIs & Metrics

### Three Primary Business KPIs (Final Working Sheet)

The `Final Working` sheet structures all executive KPIs in an **Actual vs. Target vs. % Achieved vs. Gap** framework — making performance variance immediately visible.

#### KPI 1 — Total Sales Revenue

| Metric | Value |
|---|---|
| **Actual Revenue** | ₹1,02,26,625 |
| **Revenue Target** | ₹1,25,78,748.75 |
| **Achievement %** | **81.3%** |
| **Revenue Gap (shortfall)** | ₹23,52,123.75 (18.7% remaining) |

#### KPI 2 — Total Units Sold

| Metric | Value |
|---|---|
| **Actual Units** | 33,874 units |
| **Units Target** | 36,922.66 units |
| **Achievement %** | **91.7%** |
| **Units Gap** | 3,048 units (8.3% remaining) |

#### KPI 3 — Geographic Coverage (Number of Cities)

| Metric | Value |
|---|---|
| **Actual Cities Covered** | 23 cities |
| **City Coverage Target** | 28 cities |
| **Achievement %** | **82.1%** |
| **Coverage Gap** | 5 cities not yet activated (17.9%) |

### Supporting Performance Metrics

| Metric | Value | Business Interpretation |
|---|---|---|
| **Total Transactions** | 10,000 | High transaction velocity — active pipeline |
| **Avg Transaction Value** | ₹1,022.66 | Mid-range basket — upselling opportunity exists |
| **Median Transaction Value** | ₹636 | Mean-median gap signals high-value bulk orders skewing average |
| **Max Single Transaction** | ₹12,175 | Institutional bulk orders — worth tracking separately |
| **YoY Revenue Growth (2013→2014)** | **+3.30%** | Positive but below target trajectory |
| **Avg Units per Transaction** | 3.39 | Low bundle size — multi-unit promotions underutilised |

### Regional Performance Breakdown

| Region | Revenue (₹) | Transactions | Avg Transaction (₹) | Share |
|---|---|---|---|---|
| **North** | **36,55,726** | 3,521 | 1,038 | **35.7%** |
| West | 31,18,347 | 3,111 | 1,002 | 30.5% |
| South | 17,80,729 | 1,665 | 1,070 | 17.4% |
| East | 16,71,823 | 1,703 | 982 | 16.4% |

---

## 🖥️ Dashboard / Workbook Walkthrough

*Presenting each sheet as a stakeholder-ready analytical narrative.*

---

### 📊 Sheet 1 — Sales Data (Transaction Ledger & Data Engine)

**Purpose:** The foundational raw data repository and the source for all downstream calculations. This sheet is the MIS executive's primary refresh point.

**Structure:** 10,000 rows × 10 columns, formatted as a structured Excel Table (`Table7`, columns C–L). The Table format means that as new rows are added, all formulas and conditional formatting rules auto-propagate — making the workbook maintainable without manual intervention.

**Key Formula — `Total Sales` Column:**
```excel
=G2*F2   →   Units × Price = Total Sales
```
This formula is applied to all 10,000 rows, dynamically computing transaction revenue. When unit counts or pricing change in the source data, `Total Sales` updates automatically — no manual recalculation required.

**Conditional Formatting — Weekend Transaction Highlighter:**
A formula-based conditional formatting rule highlights every row where the transaction occurred on a Saturday or Sunday:
```excel
=WEEKDAY($C2, 2) > 5
```
This visual highlight surfaces weekend sales activity at a glance, enabling field managers to analyse whether weekend transactions differ in product mix, units, or average value from weekday patterns — a question relevant to staffing, coverage planning, and promotional timing.

**Navigation Bar:** A persistent navigation strip (embedded drawing group) allows users to jump directly to `DashBoard` or `Final Working` without scrolling or using sheet tabs — designed for accessibility by non-technical stakeholders.

---

### 📋 Sheet 2 — Final Working (KPI Analysis Engine)

**Purpose:** The analytical calculation layer — KPI scorecards and structured summary tables that bridge raw data and executive dashboard.

**KPI Scorecard Panel (columns C–D):**
The left column presents the three primary KPIs in a structured `Actual / Target / % Achieved / Difference` format. This is the management reporting core of the workbook — designed to be printed and presented in weekly business reviews or emailed as a PDF snapshot without any further formatting.

**Three Structured Summary Tables:**

**Table 1 — Products by Units Sold (`Table1`, G4:H15):**
Ranks all 11 Nestle SKUs by total units sold. Nestle Smarties Pop-Up leads with 5,267 units; Nestle Rowntree Screamers trails at 969 units — revealing a 5.4× performance gap between the portfolio's best and worst-performing products.

**Table 2 — States by Total Sales (`Table2`, J4:K19):**
Ranks all 15 states by total revenue. Maharashtra (₹17.8L) and Delhi (₹17.6L) together account for 34.7% of national revenue — a geographic concentration that creates systemic business risk. States like Rajasthan (₹1.4L) contribute less than 2% — signalling material under-penetration.

**Table 3 — Cities by Total Sales (`Table3`, M4:N27):**
Ranks all 23 active cities by revenue. Delhi leads (₹17.6L), followed by Nainital (₹6.6L) and Goa (₹6.5L). Cities like Surat (₹1.7L) and Varanasi (₹2.0L) are significant population centres generating limited revenue — identifying either distribution gaps or early-stage market presence.

**Embedded Charts (4 charts):**
The Final Working sheet contains 4 embedded charts providing visual representation of the three summary tables — enabling analysis-level visual review without navigating to the DashBoard.

---

### 📈 Sheet 3 — DashBoard (Executive Visual Dashboard)

**Purpose:** The stakeholder-facing executive dashboard — a single, visually rich screen summarising all critical business performance metrics through 6 charts and branded imagery.

**Chart Inventory:**

**Chart 1 — Sales Trend Line Chart (`chart1.xml`):**
A time-series line chart plotting monthly total sales across the 2013–2014 period. The dashboard's headline visual — immediately showing the revenue trajectory, seasonal patterns, and year-over-year momentum. Critical insight visible in this chart: **Q4 (November–December) dominates both years**, with November 2013 (₹17.7L) and December 2014 (₹19.7L) dwarfing every other month — a severe seasonal concentration signal.

**Doughnut Charts ×3 (`chart2.xml`, `chart3.xml`, `chart4.xml`):**
Three doughnut charts delivering proportional breakdowns of:
- **Regional Revenue Share:** North (35.7%) and West (30.5%) together represent 66.2% of revenue
- **Product Mix by Revenue:** Nestle Smarties Pop-Up and Nestle Fab are the two largest slices
- **KPI Achievement Ring:** A progress-ring visualisation showing 81.3% revenue achievement vs. 18.7% gap — the dashboard's most impactful KPI visual for executive audiences

**ChartEx Charts ×2 (`chartEx1.xml`, `chartEx2.xml`):**
Two advanced ChartEx charts (Excel 2019/365 exclusive) providing:
- **Product Ranking Chart:** A visually enhanced ranking of all 11 products by units sold — more compelling than a standard bar chart for executive presentation contexts
- **Geographic Performance Chart:** A ranked view of state or city revenue performance in descending order — immediately surfacing the concentration in top-2 states

**Design Elements:**
- 5 embedded PNG brand images + 1 EMF graphic for visual identity
- Internal navigation hyperlinks (all three sheets inter-connected)
- Support email (`support@amazon.com`) embedded for workbook queries
- Professional grouped shapes layout across the dashboard canvas

---

## 💡 Key Insights

> *Translating workbook patterns into strategic business intelligence.*

### Insight 1 — Revenue Gap Is 18.7% But Units Gap Is Only 8.3% — a Pricing Mix Problem, Not a Volume Problem

The workbook's KPI scorecard reveals that the team sold 91.7% of target units but delivered only 81.3% of target revenue. The gap differential (18.7% vs. 8.3%) proves the shortfall is not primarily a volume problem — it is a **product mix and pricing issue**. The team is completing enough transactions but selling too many low-priced SKUs and not enough high-value products. The product table confirms this: Nestle Nobbly Bobbly averages ₹1,189 per transaction vs. Nestle Rowntree Screamers at ₹929 — a 28% per-transaction value difference within the same portfolio.

### Insight 2 — Q4 Represents ~70% of Annual Revenue — A Critical Seasonal Concentration Risk

Analysis of the monthly sales trend reveals that November and December combined contribute approximately 70% of annual revenue in both 2013 and 2014. November 2013: ₹17.7L, December 2013: ₹18.3L, November 2014: ₹18.4L, December 2014: ₹19.7L. The remaining 10 months collectively generate only ~30%. This means any supply chain disruption, competitor campaign, or distribution failure in Q4 makes the full-year target mathematically unreachable — a systemic business risk that the trend chart makes immediately visible.

### Insight 3 — Two States Drive 34.7% of Revenue — A Geographic Concentration That Creates Fragility

Maharashtra (₹17.8L) and Delhi (₹17.6L) account for over one-third of total national revenue. If either market is disrupted — by a new competitive entrant, state-level regulatory change, or distribution partner issue — the national P&L impact would be immediate and substantial. The State table on the Final Working sheet makes this concentration explicit for the first time, enabling leadership to consciously decide whether to reduce concentration risk through geographic diversification or double down on these high-performing markets.

### Insight 4 — The 5-City Coverage Gap Represents an Estimated ₹22L Untapped Revenue Opportunity

The city coverage shortfall (23 actual vs. 28 target) is not an abstract gap — it represents 5 specific cities where Nestle's sales force has not yet been deployed or distribution established. With the 23 active cities generating an average of ₹4.4L each, activating 5 additional cities at the average performance rate would add approximately **₹22L in annual revenue** — more than enough to close the ₹23.5L revenue shortfall identified in KPI 1.

### Insight 5 — Nestle Rowntree Screamers Is the Portfolio's Weakest SKU — A Rationalisation Signal

With 969 units and ₹2.85L revenue, Nestle Rowntree Screamers generates just 2.8% of total revenue — while Nestle Smarties Pop-Up generates 15.5%. The 5.4× performance gap between the best and worst products in an 11-SKU portfolio is a classic signal that the product range has grown beyond its optimal size. The product table makes this visible without any manual analysis — enabling a data-driven product rationalisation decision.

### Insight 6 — YoY Growth of 3.3% Is Positive But Structurally Below Target Trajectory

Revenue grew from ₹50.3L in 2013 to ₹51.96L in 2014 — a modest 3.3% year-on-year improvement. The annual target (₹1.25Cr) implies a growth trajectory requiring approximately 23%+ above the 2013 base. The workbook quantifies this performance gap for the first time — making it clear that incremental improvements in field execution will be insufficient, and that a structural strategy change (new product mix, geographic expansion, or pricing intervention) is required to hit the stated ambition.

### Insight 7 — South Region Achieves the Highest Avg Transaction Value — a Pricing Model Worth Replicating

South Region records the highest average transaction value (₹1,070) of all four regions, despite only 1,665 transactions — the second-lowest volume. North Region has both the highest total revenue (₹36.6L) and strong per-transaction value (₹1,038). East Region (₹982 avg transaction) represents the clearest opportunity: applying the South's pricing discipline to East's 1,703 transactions could add approximately ₹1.5L to East's revenue without a single additional transaction.

---

## 💼 Business Recommendations

**1. Immediately Realign Sales Incentives Toward High-Value SKUs to Close the Revenue-Volume Gap**
Since the revenue gap (18.7%) is more than double the units gap (8.3%), the fastest path to closing the shortfall is shifting the product mix toward higher-priced SKUs. Sales rep commission structures should be revised to specifically reward sales of Nestle Nobbly Bobbly (highest avg transaction: ₹1,189) and Nestle Rolo Potz (₹1,128 avg) over Nestle Rowntree Screamers (₹929 avg). A 10% shift in mix toward premium SKUs could generate approximately ₹5–8L in additional revenue from the same transaction volume.

**2. Activate the 5 Missing Cities as a Priority Q1 Initiative to Build Next Year's Revenue Base**
The city coverage gap (5 cities, ~₹22L opportunity) should be converted into a specific Q1 distribution expansion plan. Cities should be prioritised by proximity to existing logistics infrastructure, state GDP per capita, and competitive presence data. The city and state tables in the Final Working sheet provide the performance benchmark for each existing city — enabling a fact-based business case for each new city investment decision.

**3. Build a Mid-Year Revenue Smoothing Programme to Reduce Q4 Dependency**
With Q4 representing ~70% of annual revenue, a structured mid-year programme targeting July–September (historically the weakest months) is essential for business resilience. Ice cream and summer-appropriate SKUs (Nestle Fab, Kit Kat Cone, Nobbly Bobbly, Maxibon Cookie) should be the focus of promotional campaigns and additional field coverage during warm months — building Q2/Q3 contribution and distributing annual risk across quarters.

**4. Conduct a Formal Product Portfolio Review — Rationalise or Relaunch Rowntree Screamers**
Nestle Rowntree Screamers' 2.8% revenue share on 2.7% of transactions makes it a low-priority, low-differentiation product in the current portfolio. A formal product review should determine whether Screamers warrants a focused relaunch (targeted to specific cities or demographics) or formal SKU rationalisation — redirecting field sales bandwidth to higher-performing products. The product table provides the performance evidence for this committee decision.

**5. Replicate South Region's Premium Pricing Discipline in East Region**
East Region's average transaction value (₹982) trails South (₹1,070) by 9%. This gap likely reflects differences in product mix, customer profile, or sales rep training. A focused SKU mix analysis for East — identifying which high-value products are underrepresented relative to their national share — followed by targeted sales coaching could close this efficiency gap within two quarters.

**6. Develop a Maharashtra and Delhi Resilience Strategy**
The 34.7% revenue concentration in two states represents a risk that leadership should consciously manage. This does not necessarily mean reducing investment in these markets — but it does mean establishing contingency planning (alternative distributors, backup product lines, rapid-response promotional budgets) specifically for Maharashtra and Delhi. Simultaneously, a 3-year geographic diversification roadmap — targeting states currently below 3% revenue share (Rajasthan, Jammu & Kashmir, Himachal Pradesh) — should be initiated.

---

## 🛠️ Excel Features Used

### Data Management

| Feature | Implementation | Business Value |
|---|---|---|
| **Excel Table (`Table7`)** | 10,000-row structured table across columns C–L | Auto-expands formulas and formatting when new data is appended |
| **Summary Tables (`Table1`–`Table4`)** | Structured tables for Products, States, Cities | Clean named ranges for chart data sources — auto-update with data refresh |
| **Formula Column — Total Sales** | `=G2*F2` replicated 10,000 times | Revenue computed automatically from Units × Price — no manual calculation |

### Formulas & Logic

| Formula | Location | Purpose |
|---|---|---|
| `=G2*F2` | Sales Data, Column H (all 10,000 rows) | Calculates Total Sales for every transaction |
| `=WEEKDAY($C2, 2) > 5` | Conditional Formatting rule | Identifies weekend transactions (returns TRUE for Sat/Sun) |
| `=Actual/Target` | Final Working, KPI section | Computes % achievement for Revenue, Units, Cities KPIs |
| `=1 - % Achieved` | Final Working, KPI section | Computes remaining gap to target as a percentage |

### Charts & Visualisation

| Chart | Type | Sheet | Data Source | Business Use |
|---|---|---|---|---|
| Sales Trend | Line Chart | DashBoard | Monthly aggregated total sales | Revenue trajectory + seasonal pattern analysis |
| Regional Split | Doughnut | DashBoard | Regional revenue totals | Geographic revenue distribution at a glance |
| Product Mix | Doughnut | DashBoard | Product revenue totals | Portfolio revenue concentration visualisation |
| KPI Achievement | Doughnut | DashBoard | Actual vs. target | Progress ring showing 81.3% vs. 18.7% gap |
| Product Ranking | ChartEx | DashBoard | `_xlchart.v2.*` named ranges | Modern ranked product performance chart |
| Geographic Ranking | ChartEx | DashBoard | `_xlchart.v5.*` named ranges | State/city performance ranked view |

### Formatting & UX Design

| Feature | Application | Business Impact |
|---|---|---|
| **Conditional Formatting** | Weekend rows highlighted using `WEEKDAY()` formula across Sales Data | Immediate visual identification of weekend sales activity patterns |
| **Internal Hyperlinks** | Navigation between all 3 sheets via embedded drawing buttons | Non-technical stakeholders navigate without using sheet tabs |
| **External Hyperlink** | `mailto:support@amazon.com` embedded on each sheet | In-workbook support access for reporting queries |
| **Embedded Images (5 PNG + 1 EMF)** | Branded visual design elements on DashBoard | Professional executive presentation quality |
| **Drawing Groups** | Grouped shapes and images across all sheets | Aligned, layered visual layouts — consistent positioning |

### Named Ranges & Chart Architecture

The workbook uses Excel chart named ranges (`_xlchart.v2.4` through `_xlchart.v5.3`) that map ChartEx data series directly to table ranges on the Final Working sheet — ensuring both modern charts update automatically when the underlying summary tables are refreshed.

---

## ▶️ How to Use

### Prerequisites

- **Microsoft Excel 2019, Excel 2021, or Excel 365** — required for ChartEx chart types on the DashBoard
- **Windows or macOS** — both platforms supported
- **No external data connections, add-ins, or macros** — fully self-contained `.xlsx` file

### Opening the Workbook

```
1. Download Sales_Data.xlsx from this repository
2. Open in Microsoft Excel (double-click or File → Open)
3. If prompted with "Enable Editing" — click it to activate formulas
4. You will land on the Sales Data sheet by default
5. Navigate using the on-sheet buttons (or sheet tabs at the bottom):
      ├── Sales Data       →  10,000-row raw transaction table
      ├── Final Working    →  KPI scorecard + summary tables + working charts
      └── DashBoard        →  Executive visual dashboard
```

### Using the Navigation Buttons

```
→ Each sheet has navigation buttons embedded as drawing objects
→ "DashBoard" button → jumps to executive dashboard
→ "Final Working" button → jumps to KPI analysis layer
→ "Sales Data" button → returns to raw transaction ledger
→ Designed for stakeholders who don't use sheet tabs
```

### Refreshing with New Data

```
1. Navigate to the Sales Data sheet
2. Click below row 10001 — the Excel Table auto-extends as you add rows
3. Enter: Date | SalesRep | Product | Units | Price | (Total Sales auto-calculates) | City | State | Region | Day
4. The Total Sales formula (=Units × Price) auto-populates in the new row
5. Update target values in the Final Working sheet (cells D5, D12, D19) for the new period
6. All summary tables, KPI scorecards, and DashBoard charts update automatically
```

### Updating KPI Targets

```
→ On the Final Working sheet:
   • Revenue Target  → Cell D5  → replace 12578748.75 with updated target
   • Units Target    → Cell D12 → replace 36922.66 with updated target
   • Cities Target   → Cell D19 → replace 28 with updated coverage target
→ % Achieved and Difference formulas recalculate automatically
```

---

## 📁 Project Structure

```
Nestle-Sales-Excel-Dashboard/
│
├── Sales_Data.xlsx                    # Primary workbook — all data, KPIs, and dashboard
│   ├── Sales Data sheet               # 10,000 transaction rows (Excel Table, columns C–L)
│   ├── Final Working sheet            # KPI scorecard + 3 summary tables + 4 embedded charts
│   └── DashBoard sheet                # Executive visual dashboard — 6 charts + branded design
│
├── Screenshots/                       # Workbook page exports for portfolio preview
│   ├── 01_Sales_Data_Sheet.png        # Raw transaction table with weekend conditional formatting
│   ├── 02_Final_Working_Sheet.png     # KPI scorecards + Product/State/City summary tables
│   └── 03_Dashboard_Sheet.png         # Executive dashboard with all 6 charts + navigation
│
└── README.md                          # Project documentation (this file)
```

### Internal Workbook Reference

| Component | Detail |
|---|---|
| **Excel Tables** | 4 (`Table1`, `Table2`, `Table3`, `Table7`) |
| **Total Formula Cells** | 10,000 (all in `Total Sales` column of Sales Data) |
| **Charts** | 6 (1 line, 3 doughnut, 2 ChartEx modern) |
| **Conditional Formatting Rules** | 4 (all using `WEEKDAY()` formula logic) |
| **Embedded Images** | 5 PNG + 1 EMF across all three sheets |
| **Hyperlinks** | 3 internal sheet navigation + 1 external mailto |
| **Named Ranges** | 8 chart data series ranges (`_xlchart.v2.*`, `_xlchart.v5.*`) |
| **External Links** | 1 (legacy reference — does not affect functionality) |

---

## 👤 Author

<table>
  <tr>
    <td align="center">
      <b>Vishal Londhekar</b><br/>
      <i>Data Analyst | Business Analyst | Data Science</i><br/><br/>
      <a href="https://github.com/vishal-Londhekar">🔗 GitHub</a>
    </td>
  </tr>
</table>

> *"Excel is not just a spreadsheet — it is the most widely deployed business intelligence platform in the world. This workbook demonstrates that with structured data, disciplined formula logic, and purposeful chart design, Excel can deliver the same analytical clarity as purpose-built BI tools — accessible to every stakeholder, on every device, with zero licensing overhead."*

---

## ⭐ If this project strengthened your Excel analytics portfolio, please star the repository!

---

<p align="center">
  <img src="https://img.shields.io/badge/Built%20with-Microsoft%20Excel-217346?logo=microsoftexcel" />
  <img src="https://img.shields.io/badge/Domain-FMCG%20Sales%20Analytics-orange" />
  <img src="https://img.shields.io/badge/Records-10%2C000%20Transactions-blue" />
  <img src="https://img.shields.io/badge/Charts-6%20(Line%20%7C%20Doughnut%20%7C%20ChartEx)-blueviolet" />
  <img src="https://img.shields.io/badge/KPIs-Revenue%20%7C%20Units%20%7C%20Cities-green" />
</p>
