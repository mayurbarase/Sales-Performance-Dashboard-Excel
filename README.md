📊 Interactive Sales Performance Dashboard (Excel)
A fully interactive sales dashboard built in Microsoft Excel that turns 1,000 rows of raw e-commerce orders into a clear, one-page view of business performance.
Pick a Region, Category, or Ship Mode with the slicers, and every KPI and chart updates instantly.
<img width="1790" height="880" alt="Screenshot 2026-09-19 182730" src="https://github.com/user-attachments/assets/d0518592-0a5e-4ffb-9116-b204515e3aa3" />

---
📌 Why I Built This
Raw transaction data is hard to read. Rows and columns don't tell you which region is pulling ahead, which category actually makes money, or whether sales are trending up or down.
I built this dashboard to answer questions like:
How much revenue and profit have we made?
Which regions sell the most, and which are the most profitable?
Which product categories drive sales and profit?
How is revenue split across shipping methods?
How do sales change month to month?
The goal was to practice the full analytics workflow in Excel: Raw Data → Analysis → Visualization → Interactive Dashboard.
---
✨ Key Features
5 KPI cards for Total Sales, Total Profit, Total Orders, Profit Margin, and Highest Sale
6 linked charts covering region, category, shipping, and time trends
3 slicers (Region, Category, Ship Mode) connected to all PivotTables and charts
Formula-driven analysis sheet for transparent, checkable KPI calculations
Clean BI-style layout with a blue and teal palette and minimal clutter
Validated numbers, cross-checked between formulas, PivotTables, and raw data
---
📈 Headline Numbers
KPI	Value
Total Sales	₹90,202.61
Total Profit	₹28,148.13
Total Orders	1,000
Profit Margin	31%
Highest Single Sale	₹619.75
---

📊 Dashboard Visuals
Visual	Chart Type	PivotTable Setup	What It Shows
Sales by Region	Column	Region → Sum of Total Revenue	Revenue comparison across regions
Profit by Region	Horizontal Bar	Region → Sum of Profit	Which regions are most profitable
Sales by Ship Mode	Doughnut	Ship Mode → Sum of Total Revenue	Share of revenue by shipping method
Profit by Category	Horizontal Bar	Category → Sum of Profit	Category profitability
Sales by Category	Treemap	Category → Sum of Total Revenue	Relative size of each category at a glance
Monthly Sales	Line	Order Date (grouped by Year, Quarter, Month) → Sum of Total Revenue	Sales trend over time

🎛️ Slicers
Three slicers control the whole dashboard: Region, Category, and Ship Mode. Select a region, for example, and the KPIs, category breakdown, shipping split, and monthly trend all update to match.
<img width="1770" height="873" alt="Screenshot 2026-09-19 182838" src="https://github.com/user-attachments/assets/bedd79bf-06d1-4dd7-ad59-bdb18627fe22" />
<img width="1791" height="876" alt="Screenshot 2026-09-19 183159" src="https://github.com/user-attachments/assets/d417e384-ec2f-4360-8c76-fe37fb7bc0a4" />

---
🗂️ Dataset
The data is an e-commerce orders dataset with 1,000 orders, sourced from TrumpExcel's E-Commerce Orders Dataset.
Column	Description
Order ID	Unique identifier for each order
Order Date	Date the order was placed
Customer Name	Name of the customer
Region	Geographic region of the order
Product Name	Name of the product
Category	Product category
Quantity	Units purchased
Unit Price	Price per unit
Discount %	Discount applied to the order
Total Revenue	Revenue generated from the order
Profit	Profit generated from the order
Ship Mode	Shipping method selected
---
🧮 How the Numbers Are Calculated
The `Analysis` sheet holds all KPI formulas, written with structured references to the `SalesData` table.
```excel
Total Sales           =SUM(SalesData[Total Revenue])
Total Profit          =SUM(SalesData[Profit])
Total Orders          =COUNTA(SalesData[Order ID])
Average Order Value   =AVERAGE(SalesData[Total Revenue])
Total Units Sold      =SUM(SalesData[Quantity])
Highest Sale          =MAX(SalesData[Total Revenue])
Lowest Sale           =MIN(SalesData[Total Revenue])
Profit Margin         =B3/B2    // Total Profit / Total Sales
```
Keeping the calculations on their own sheet makes the workbook easier to maintain and gives me a place to double-check the dashboard.
---
📑 Workbook Structure
```text
Interactive_Sales_Performance_Dashboard.xlsx
│
├── Dashboard      → KPI cards, slicers, and charts (the sheet users interact with)
├── Analysis       → Formula-based KPI calculations
├── PivotTables    → PivotTables that power every chart
└── Raw Data       → Original dataset stored as an Excel Table named "SalesData"
```
---
🔍 Data Validation
To make sure the dashboard can be trusted, I cross-checked the results:
Total Sales on the Analysis sheet matches the PivotTable grand total
Total Profit matches the Profit PivotTable
Total Orders matches the number of records in the raw data
Average Order Value equals Total Sales ÷ Total Orders
Profit Margin equals Total Profit ÷ Total Sales
Highest and Lowest Sale values match the raw data
---
🎨 Design Approach
I wanted this to look like a real BI report, not a pile of charts on a worksheet:
Clear header with a last-updated date
KPI cards at the top so key numbers are seen first
A dedicated slicer panel
Consistent chart sizing, fonts, and colors
White chart panels on a light background
---
🛠️ Tools & Skills
Microsoft Excel · Excel Tables · Structured References · PivotTables · PivotCharts · Slicers · Date Grouping · KPI Development · Dashboard Design · Data Visualization
Functions used: `SUM`, `AVERAGE`, `COUNTA`, `MAX`, `MIN`
> No VBA/macros or Power Query were used. Everything is built with native Excel features.
---
💡 What I Learned
This project helped me see how Excel's features fit together in a real analytics workflow, not just as isolated tricks. The biggest lesson: a good dashboard isn't about adding charts. It's about deciding what to measure, how to present it, and how someone will interact with it.
Along the way I practiced structuring data as Excel Tables, writing structured-reference formulas, choosing the right chart for each question, connecting multiple PivotCharts through slicers, and validating results before presenting them.
---
🚀 Future Improvements
Deeper time-based analysis (year-over-year, growth rates)
Customer-level and product-level performance views
Sales targets and variance tracking
Additional profitability metrics
Rebuilding the dashboard in Power BI
Connecting to a regularly refreshed data source
---
📁 Repository Structure
```text
Sales-Performance-Dashboard-Excel/
│
├── Interactive_Sales_Performance_Dashboard.xlsx
├── README.md
└── images/
    ├── dashboard-full.png
    ├── dashboard-region-filter.png
    └── dashboard-shipmode-filter.png
```
---
▶️ How to Use
Download `Interactive_Sales_Performance_Dashboard.xlsx`.
Open it in Microsoft Excel (desktop version recommended, since slicers and Treemap charts work best there).
Go to the Dashboard sheet and click the slicers to explore the data.
---
👋 About
I built this as part of my hands-on journey into Data Analytics and Business Intelligence. Instead of only working through theory, I wanted a project that looks like the reporting work done in real businesses.
If you have feedback or ideas, I'd love to hear them.
GitHub: your-username
LinkedIn: your-name
---
Status: Completed ✅
⭐ If you found this project useful, consider giving the repo a star!
