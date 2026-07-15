# Sales Analysis — SQL + Power BI

> **AdventureWorks DW 2019** · Internet Sales · 2019–2021  
> 📁 Repository: [vatsalgajera-tech/Sales-Analysis-Dashboard](https://github.com/vatsalgajera-tech/Sales-Analysis-Dashboard)

Interactive sales reporting pipeline built on T-SQL data cleaning and Power BI dashboards.

---

## Project Structure

```
SalesAnalysis_SQL_PowerBI/
│
├── sql/                                        # T-SQL cleaning scripts
│   ├── DIM_Calendar_Clean.sql                  (743 B)
│   ├── DIM_Customer_Clean.sql                  (1.3 KB)
│   ├── DIM_Product_Clean.sql                   (1.6 KB)
│   └── FACT_InternetSales_Clean.sql            (866 B)
│
├── data/
│   ├── raw/                                    # CSV exports from SQL Server
│   │   ├── DIM_Calendar_Export.csv             (53 KB)
│   │   ├── DIM_Customer_Export.csv             (1.1 MB)
│   │   ├── DIM_Product_Export.csv              (98 KB)
│   │   └── FACT_InternetSales_Export.csv       (2.9 MB)
│   └── budget/
│       └── Sent Over Data - SalesBudget.xlsx   (8.8 KB)
│
├── reports/
│   └── Sales Report.pbix                       (2 MB)   ← Power BI report
│
├── scripts/
│   └── generate_pbix.py                        (16.9 KB) ← Auto-generate .pbix
│
├── .gitignore
└── README.md
```

---

## Quick Start

### 1. Restore the Database
Restore **AdventureWorksDW2019** in SQL Server Express.  
Guide: [Microsoft Docs → AdventureWorks install](https://docs.microsoft.com/en-us/sql/samples/adventureworks-install-configure)

Update to recent years using:  
[Update_AdventureWorksDW_Data.sql](https://github.com/techtalkcorner/SampleDemoFiles/blob/master/Database/AdventureWorks/Update_AdventureWorksDW_Data.sql)

### 2. Run SQL Cleaning Scripts
Execute each script in `sql/` in SSMS and **export results as CSV** into `data/raw/`:

| Script | Output CSV |
|---|---|
| `DIM_Calendar_Clean.sql`       | `DIM_Calendar_Export.csv`       |
| `DIM_Customer_Clean.sql`       | `DIM_Customer_Export.csv`       |
| `DIM_Product_Clean.sql`        | `DIM_Product_Export.csv`        |
| `FACT_InternetSales_Clean.sql` | `FACT_InternetSales_Export.csv` |

### 3. Generate a new .pbix (Python)
```bash
python -m pip install pandas openpyxl
python scripts/generate_pbix.py
```
This reads all CSVs + budget XLSX and produces a `.pbix` file in `reports/`.

### 4. Open in Power BI Desktop
- Open `reports/Sales Report.pbix`
- On first open: **Transform Data → Data Source Settings** → point each table to your local `data/raw/` CSVs

---

## Data Model

```
DIM_Calendar ◄──── FACT_InternetSales ────► DIM_Customer
                          │
                          └─────────────► DIM_Product

FACT_Budget (standalone — linked via Date column)
```

| Table | Type | Rows |
|---|---|---|
| `FACT_InternetSales` | Fact   | 58,168 |
| `DIM_Customer`       | Dim    | 18,484 |
| `DIM_Product`        | Dim    | 606    |
| `DIM_Calendar`       | Dim    | 1,096  |
| `FACT_Budget`        | Budget | 18     |

### Relationships

| From | Column | To | Column |
|---|---|---|---|
| `FACT_InternetSales` | `CustomerKey`   | `DIM_Customer` | `CustomerKey` |
| `FACT_InternetSales` | `ProductKey`    | `DIM_Product`  | `ProductKey`  |
| `FACT_InternetSales` | `OrderDateKey`  | `DIM_Calendar` | `DateKey`     |
| `FACT_InternetSales` | `DueDateKey`    | `DIM_Calendar` | `DateKey` *(inactive)* |
| `FACT_InternetSales` | `ShipDateKey`   | `DIM_Calendar` | `DateKey` *(inactive)* |

---

## Dashboard Pages

| Page | Visuals |
|---|---|
| **Sales Overview**  | Revenue KPI · Budget KPI · Monthly line chart · Category donut · Top 10 Products bar · Top 10 Customers bar · City map |
| **Customer Detail** | Customer table · Gender pie chart · Customers by city bar |
| **Product Detail**  | Revenue by sub-category bar · Product table · Category pie chart |

---

## DAX Measures

| Measure | Expression |
|---|---|
| `Total Sales` | `SUM(FACT_InternetSales[SalesAmount])` |
| `Total Orders` | `DISTINCTCOUNT(FACT_InternetSales[SalesOrderNumber])` |
| `Sales vs Budget` | `[Total Sales] - SUM(FACT_Budget[Budget])` |
| `Budget Achievement %` | `DIVIDE([Total Sales], SUM(FACT_Budget[Budget]), 0)` |
| `Avg Sales per Customer` | `DIVIDE([Total Sales], DISTINCTCOUNT(FACT_InternetSales[CustomerKey]), 0)` |

---

## Key Metrics (2021)

| Metric | Value |
|---|---|
| Total Revenue | **$22.24M** |
| Total Orders  | **58,168**  |
| Total Customers | **18,484** |

---

## Tools

| Tool | Purpose |
|---|---|
| SQL Server Express + SSMS | Run T-SQL cleaning scripts |
| Power BI Desktop | Build & view the `.pbix` report |
| Python 3.x + pandas + openpyxl | Auto-generate `.pbix` from CSV data |

---

## Contributor

<table>
  <tr>
    <td align="center">
      <a href="https://github.com/vatsalgajera-tech">
        <img src="https://github.com/vatsalgajera-tech.png" width="80" style="border-radius:50%"/><br/>
        <b>Vatsal Gajera</b>
      </a><br/>
      <sub>SQL · Power BI · Python</sub>
    </td>
  </tr>
</table>

[![GitHub](https://img.shields.io/badge/GitHub-vatsalgajera--tech-181717?style=flat&logo=github)](https://github.com/vatsalgajera-tech)
[![Repo](https://img.shields.io/badge/Repo-Sales--Analysis--Dashboard-blue?style=flat&logo=github)](https://github.com/vatsalgajera-tech/Sales-Analysis-Dashboard)

---

## Acknowledgements

- Dataset: [AdventureWorks DW 2019](https://docs.microsoft.com/en-us/sql/samples/adventureworks-install-configure) — Microsoft Sample Database
- Original project structure inspired by the Power BI community
