"""
generate_pbix.py
================
Generates a fully structured Power BI (.pbix) file from the AdventureWorks
Sales Analysis CSV exports.

A .pbix file is a ZIP archive containing:
  - [Content_Types].xml
  - Version
  - Settings
  - SecurityBindings
  - DataModel          (schema + relationships as JSON)
  - Report/Layout      (JSON visual layout, UTF-16-LE encoded)
  - DiagramState       (data model diagram positions)
  - Metadata           (JSON)

Requirements:
    pip install pandas openpyxl

Usage:
    python scripts/generate_pbix.py
"""

import os
import json
import zipfile
import datetime
import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_RAW    = os.path.join(BASE_DIR, "data", "raw")
DATA_BUDGET = os.path.join(BASE_DIR, "data", "budget")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
OUTPUT_PBIX = os.path.join(REPORTS_DIR, "Sales_Analysis_Generated.pbix")

os.makedirs(REPORTS_DIR, exist_ok=True)

# ── 1. Load Data ──────────────────────────────────────────────────────────────
print("Loading data files...")

df_calendar = pd.read_csv(os.path.join(DATA_RAW, "DIM_Calendar_Export.csv"))
df_customer = pd.read_csv(os.path.join(DATA_RAW, "DIM_Customer_Export.csv"))
df_product  = pd.read_csv(os.path.join(DATA_RAW, "DIM_Product_Export.csv"))
df_sales    = pd.read_csv(os.path.join(DATA_RAW, "FACT_InternetSales_Export.csv"))

budget_file = os.path.join(DATA_BUDGET, "Sent Over Data - SalesBudget.xlsx")
df_budget   = pd.read_excel(budget_file) if os.path.exists(budget_file) else pd.DataFrame()

print(f"  Calendar : {len(df_calendar):,} rows")
print(f"  Customers: {len(df_customer):,} rows")
print(f"  Products : {len(df_product):,} rows")
print(f"  Sales    : {len(df_sales):,} rows")
print(f"  Budget   : {len(df_budget):,} rows")

# ── 2. Key Metrics ────────────────────────────────────────────────────────────
sales_col = next(
    (c for c in df_sales.columns if c == "SalesAmount"), None
) or next(
    (c for c in df_sales.columns
     if ("amount" in c.lower()) and df_sales[c].dtype in ["float64","int64"]), None
)
total_sales  = float(df_sales[sales_col].sum()) if sales_col else 0.0
total_orders = len(df_sales)
total_cust   = int(df_customer["CustomerKey"].nunique()) if "CustomerKey" in df_customer.columns else len(df_customer)

print(f"\nKey Metrics:")
print(f"  Total Sales    : ${total_sales:,.2f}")
print(f"  Total Orders   : {total_orders:,}")
print(f"  Total Customers: {total_cust:,}")


# ── 3. Helpers ────────────────────────────────────────────────────────────────
TYPE_MAP = {
    "int64": "Int64", "int32": "Int64", "float64": "Double",
    "float32": "Double", "object": "String", "bool": "Boolean",
    "datetime64[ns]": "DateTime"
}


def df_to_columns(df):
    return [
        {"name": col, "dataType": TYPE_MAP.get(str(dtype), "String")}
        for col, dtype in df.dtypes.items()
    ]


def make_visual(vis_type, x, y, w, h, title, table, col, v_id):
    return {
        "id": v_id,
        "position": {"x": x, "y": y, "z": 0, "width": w, "height": h, "tabOrder": v_id},
        "visual": {
            "visualType": vis_type,
            "title": {"show": True, "text": title},
            "projections": {
                "Category": [{"queryRef": f"{table}.{col}"}],
                "Y":        [{"queryRef": f"{table}.SalesAmount"}],
            },
            "prototypeQuery": {
                "Version": 2,
                "From":    [{"Name": "s", "Entity": table, "Type": 0}],
                "Select": [
                    {"Column": {"Expression": {"SourceRef": {"Source": "s"}},
                                "Property": col},
                     "Name": f"s.{col}"},
                    {"Aggregation": {
                        "Expression": {"Column": {
                            "Expression": {"SourceRef": {"Source": "s"}},
                            "Property": "SalesAmount"}},
                        "Function": 0},
                     "Name": "Sum(SalesAmount)"},
                ],
                "OrderBy": [{"Direction": 2, "Expression": {
                    "Aggregation": {"Expression": {"Column": {
                        "Expression": {"SourceRef": {"Source": "s"}},
                        "Property": "SalesAmount"}}, "Function": 0}}}],
            },
            "objects": {
                "background": [{"properties": {"color": {"solid": {"color": "#1a2035"}}}}],
            },
        },
    }


# ── 4. Report Layout ──────────────────────────────────────────────────────────
print("\nBuilding report layout...")

pages = [
    {
        "id": "ReportSection1",
        "name": "Sales Overview",
        "displayName": "Sales Overview",
        "width": 1280,
        "height": 720,
        "background": {"color": {"solid": {"color": "#0f1117"}}},
        "visualContainers": [
            make_visual("card",       20,  20, 280, 100,
                        f"Total Revenue  ${total_sales/1e6:.2f}M",
                        "FACT_InternetSales", "SalesAmount", 1),
            make_visual("card",       320, 20, 280, 100,
                        f"Total Orders  {total_orders:,}",
                        "FACT_InternetSales", "SalesOrderNumber", 2),
            make_visual("card",       620, 20, 280, 100,
                        f"Customers  {total_cust:,}",
                        "DIM_Customer", "CustomerKey", 3),
            make_visual("lineChart",  20, 140, 820, 260,
                        "Revenue vs Budget - Monthly",
                        "FACT_InternetSales", "SalesAmount", 4),
            make_visual("donutChart", 860, 140, 400, 260,
                        "Sales by Category",
                        "DIM_Product", "Product Category", 5),
            make_visual("barChart",   20, 420, 400, 280,
                        "Top 10 Products by Revenue",
                        "DIM_Product", "Product Name", 6),
            make_visual("barChart",   440, 420, 400, 280,
                        "Top 10 Customers",
                        "DIM_Customer", "Full Name", 7),
            make_visual("map",        860, 420, 400, 280,
                        "Sales by City",
                        "DIM_Customer", "Customer City", 8),
        ],
    },
    {
        "id": "ReportSection2",
        "name": "Customer Detail",
        "displayName": "Customer Detail",
        "width": 1280,
        "height": 720,
        "background": {"color": {"solid": {"color": "#0f1117"}}},
        "visualContainers": [
            make_visual("tableEx",  20,  20, 820, 680,
                        "Customer Details",
                        "DIM_Customer", "Full Name", 9),
            make_visual("pieChart", 860, 20, 400, 320,
                        "Gender Distribution",
                        "DIM_Customer", "Gender", 10),
            make_visual("barChart", 860, 360, 400, 340,
                        "Top Cities by Customer Count",
                        "DIM_Customer", "Customer City", 11),
        ],
    },
    {
        "id": "ReportSection3",
        "name": "Product Detail",
        "displayName": "Product Detail",
        "width": 1280,
        "height": 720,
        "background": {"color": {"solid": {"color": "#0f1117"}}},
        "visualContainers": [
            make_visual("barChart", 20,  20, 820, 320,
                        "Revenue by Sub-Category",
                        "DIM_Product", "Sub Category", 12),
            make_visual("tableEx",  20, 360, 820, 340,
                        "Product Performance",
                        "DIM_Product", "Product Name", 13),
            make_visual("pieChart", 860, 20, 400, 680,
                        "Sales Share by Category",
                        "DIM_Product", "Product Category", 14),
        ],
    },
]

theme_config = {
    "name": "AdventureWorks Dark",
    "dataColors": [
        "#3b82f6", "#8b5cf6", "#f59e0b", "#10b981",
        "#06b6d4", "#ef4444", "#f97316", "#ec4899",
    ],
    "background":   "#0f1117",
    "foreground":   "#f0f4ff",
    "tableAccent":  "#3b82f6",
}

layout = {
    "id": 0,
    "resourcePackages": [],
    "sections": pages,
    "config": json.dumps({
        "version": "5.48",
        "themeCollection": {
            "baseTheme": {
                "name":      "AdventureWorks Dark",
                "version":   "5.48",
                "themeJson": json.dumps(theme_config),
            }
        },
    }),
}

layout_bytes = json.dumps(layout, indent=2, ensure_ascii=False).encode("utf-16-le")

# ── 5. DataModel (schema + relationships + measures) ─────────────────────────
data_model = {
    "version": "1.0",
    "model": {
        "tables": [
            {
                "name":     "FACT_InternetSales",
                "columns":  df_to_columns(df_sales),
                "measures": [
                    {
                        "name":       "Total Sales",
                        "expression": "SUM(FACT_InternetSales[SalesAmount])",
                        "formatString": "$#,0.00",
                    },
                    {
                        "name":       "Total Orders",
                        "expression": "DISTINCTCOUNT(FACT_InternetSales[SalesOrderNumber])",
                        "formatString": "#,0",
                    },
                    {
                        "name":       "Sales vs Budget",
                        "expression": "[Total Sales] - SUM(FACT_Budget[Budget])",
                        "formatString": "$#,0.00",
                    },
                    {
                        "name":       "Budget Achievement %",
                        "expression": "DIVIDE([Total Sales], SUM(FACT_Budget[Budget]), 0)",
                        "formatString": "0.00%",
                    },
                    {
                        "name": "Avg Sales per Customer",
                        "expression": "DIVIDE([Total Sales], DISTINCTCOUNT(FACT_InternetSales[CustomerKey]), 0)",
                        "formatString": "$#,0.00",
                    },
                ],
            },
            {"name": "DIM_Customer", "columns": df_to_columns(df_customer)},
            {"name": "DIM_Product",  "columns": df_to_columns(df_product)},
            {"name": "DIM_Calendar", "columns": df_to_columns(df_calendar)},
            {
                "name":    "FACT_Budget",
                "columns": df_to_columns(df_budget) if len(df_budget) else [],
            },
        ],
        "relationships": [
            {
                "name":                   "Sales_Customer",
                "fromTable":              "FACT_InternetSales",
                "fromColumn":             "CustomerKey",
                "toTable":                "DIM_Customer",
                "toColumn":               "CustomerKey",
                "crossFilteringBehavior": "oneDirection",
            },
            {
                "name":                   "Sales_Product",
                "fromTable":              "FACT_InternetSales",
                "fromColumn":             "ProductKey",
                "toTable":                "DIM_Product",
                "toColumn":               "ProductKey",
                "crossFilteringBehavior": "oneDirection",
            },
            {
                "name":                   "Sales_Calendar_OrderDate",
                "fromTable":              "FACT_InternetSales",
                "fromColumn":             "OrderDateKey",
                "toTable":                "DIM_Calendar",
                "toColumn":               "DateKey",
                "crossFilteringBehavior": "oneDirection",
            },
            {
                "name":                   "Sales_Calendar_DueDate",
                "fromTable":              "FACT_InternetSales",
                "fromColumn":             "DueDateKey",
                "toTable":                "DIM_Calendar",
                "toColumn":               "DateKey",
                "crossFilteringBehavior": "oneDirection",
                "isActive":               False,
            },
            {
                "name":                   "Sales_Calendar_ShipDate",
                "fromTable":              "FACT_InternetSales",
                "fromColumn":             "ShipDateKey",
                "toTable":                "DIM_Calendar",
                "toColumn":               "DateKey",
                "crossFilteringBehavior": "oneDirection",
                "isActive":               False,
            },
        ],
    },
}

# ── 6. DiagramState ───────────────────────────────────────────────────────────
diagram_state = {
    "version": 1,
    "nodes": [
        {"NodeIndex": 0, "Name": "FACT_InternetSales",
         "Location": {"x": 420, "y": 220}},
        {"NodeIndex": 1, "Name": "DIM_Customer",
         "Location": {"x": 80,  "y": 80}},
        {"NodeIndex": 2, "Name": "DIM_Product",
         "Location": {"x": 80,  "y": 400}},
        {"NodeIndex": 3, "Name": "DIM_Calendar",
         "Location": {"x": 760, "y": 80}},
        {"NodeIndex": 4, "Name": "FACT_Budget",
         "Location": {"x": 760, "y": 400}},
    ],
    "links": [
        {"SourceNodeIndex": 0, "TargetNodeIndex": 1},
        {"SourceNodeIndex": 0, "TargetNodeIndex": 2},
        {"SourceNodeIndex": 0, "TargetNodeIndex": 3},
    ],
}

# ── 7. Remaining entries ──────────────────────────────────────────────────────
content_types = (
    '<?xml version="1.0" encoding="utf-8"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    '  <Default Extension="json" ContentType="application/json"/>\n'
    '  <Default Extension="xml"  ContentType="application/xml"/>\n'
    '  <Default Extension="png"  ContentType="image/png"/>\n'
    '  <Override PartName="/DataModel"        ContentType="application/x-json"/>\n'
    '  <Override PartName="/Report/Layout"    ContentType="application/json"/>\n'
    '  <Override PartName="/Settings"         ContentType="application/json"/>\n'
    '  <Override PartName="/Metadata"         ContentType="application/json"/>\n'
    '  <Override PartName="/Version"          ContentType="application/json"/>\n'
    '  <Override PartName="/SecurityBindings" ContentType="application/json"/>\n'
    '  <Override PartName="/DiagramState"     ContentType="application/json"/>\n'
    '</Types>'
)

metadata = {
    "version": "4.0",
    "culture": "en-US",
    "created": datetime.datetime.now().isoformat(),
    "model": {
        "name": "Sales Analysis - AdventureWorks",
        "tables": [
            {"name": "FACT_InternetSales", "rows": total_orders},
            {"name": "DIM_Customer",       "rows": len(df_customer)},
            {"name": "DIM_Product",        "rows": len(df_product)},
            {"name": "DIM_Calendar",       "rows": len(df_calendar)},
            {"name": "FACT_Budget",        "rows": len(df_budget)},
        ],
    },
}

settings = {
    "Version":                2,
    "QueryGroupsEnabled":     True,
    "AutoRecoverEnabled":     True,
    "AllowChangeFilterTypes": True,
    "PersistReportTheme":     True,
    "UseNewFilterPane":       True,
}

version = {"version": "5.48.0"}

# ── 8. Assemble ZIP ───────────────────────────────────────────────────────────
print(f"\nAssembling .pbix -> {OUTPUT_PBIX}")

with zipfile.ZipFile(OUTPUT_PBIX, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
    zf.writestr("[Content_Types].xml", content_types)
    zf.writestr("Version",             json.dumps(version))
    zf.writestr("Settings",            json.dumps(settings,      indent=2))
    zf.writestr("SecurityBindings",    json.dumps({}))
    zf.writestr("Metadata",            json.dumps(metadata,      indent=2, ensure_ascii=False))
    zf.writestr("DataModel",           json.dumps(data_model,    indent=2, ensure_ascii=False))
    zf.writestr("DiagramState",        json.dumps(diagram_state, indent=2))
    zf.writestr("Report/Layout",       layout_bytes)

size_kb = os.path.getsize(OUTPUT_PBIX) / 1024
print(f"  Created : {OUTPUT_PBIX}")
print(f"  Size    : {size_kb:.1f} KB")
print("\nDone! Open 'reports/Sales_Analysis_Generated.pbix' in Power BI Desktop.")
print("On first open -> Transform Data -> set file paths to data/raw/ CSVs.")
