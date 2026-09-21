# Veda Technology - Task 14: Excel Data Analysis

[![GitHub Repo](https://img.shields.io/badge/GitHub-veda--technology--task--14--basic-181717?logo=github)](https://github.com/KrishnaKr87/veda-technology-task-14-basic)
[![Dataset](https://img.shields.io/badge/Dataset-10%2C000%20Records-blue.svg)](file:///./Veda_Technology_Excel_Data_Analysis_Dataset.xlsx)
[![Analysis](https://img.shields.io/badge/Analysis-Excel%20%2B%20Python-success.svg)]()
[![Status](https://img.shields.io/badge/Status-Completed-brightgreen.svg)]()

This repository contains the dataset, analysis model, executive summary, and verification scripts for **Task 14 (Excel Data Analysis - Basic & Intermediate)** at **Veda Technology**.

---

## 📌 Project Overview

The objective of this project is to perform end-to-end retail sales and profitability analysis using Microsoft Excel and Python on an extensive enterprise transactional dataset spanning 2024 to 2025. 

Key objectives include:
- Cleaning and validating transaction data across 10,000 order records.
- Calculating core business metrics (Total Revenue, Net Profit, Profit Margins, Average Order Values).
- Creating multi-dimensional Pivot Tables to segment sales performance across product categories, geographic regions, customer segments, and monthly trends.
- Programmatically verifying analytical accuracy with Python (`pandas` & `openpyxl`).

---

## 📊 Dataset Structure

The primary data source is [`Veda_Technology_Excel_Data_Analysis_Dataset.xlsx`](./Veda_Technology_Excel_Data_Analysis_Dataset.xlsx), which contains three structured sheets:

| Sheet Name | Description | Dimensions |
| :--- | :--- | :--- |
| **`Data`** | Raw transactional records containing customer, shipping, geographic, product, sales, discount, and profit fields. | 10,000 rows × 21 columns |
| **`Formulas`** | Executive KPI calculations utilizing native Excel formulas. | Summary Metrics |
| **`PivotTable`** | Pivot table summaries including category breakdowns, regional profitability, monthly timelines, and top products. | Cross-tabulated Views |

### Schema Attributes (21 Columns)
- **Order & Shipping**: `Row ID`, `Order ID`, `Order Date`, `Ship Date`, `Ship Mode`
- **Customer Information**: `Customer ID`, `Customer Name`, `Segment` (Consumer, Corporate, Home Office)
- **Geographic Details**: `Country`, `City`, `State`, `Postal Code`, `Region` (Central, East, South, West)
- **Product Information**: `Product ID`, `Category`, `Sub-Category`, `Product Name`
- **Financial Metrics**: `Sales`, `Quantity`, `Discount`, `Profit`

---

## 📈 Key Performance Indicators (KPIs)

Summary of key business figures calculated across all 10,000 order lines:

| Metric | Excel Formula Equivalent | Value |
| :--- | :--- | :--- |
| **Total Revenue / Sales** | `=SUM(Data!R:R)` | **2,178,766.07** |
| **Total Net Profit** | `=SUM(Data!U:U)` | **161,597.26** |
| **Total Units Sold** | `=SUM(Data!S:S)` | **54,449 units** |
| **Total Order Lines** | `=COUNTA(Data!B2:B10001)` | **10,000** |
| **Average Order Line Value** | `=AVERAGE(Data!R:R)` | **217.88** |
| **Overall Profit Margin** | `=Total_Profit / Total_Sales` | **7.42%** |

---

## 🔍 Detailed Analytical Findings

### 1. Performance by Category
The product catalog is divided into three core categories:
- **Technology**: Dominates revenue generation with **1,133,363.93 (52.0%)** and **81,965.39** in net profit (7.23% margin).
- **Furniture**: Generates **580,358.71 (26.6%)** in sales and **43,811.48** in profit (7.55% margin).
- **Office Supplies**: Drives the highest sales volume (**23,333 units sold**) with **465,043.43 (21.3%)** in sales and a healthy **7.70% margin** ($35,820.39 profit).

### 2. Regional Breakdown
Performance is balanced across all four US geographic regions:
- **West**: **558,828.92** Sales | **40,703.90** Profit (7.28% margin)
- **South**: **547,523.24** Sales | **39,952.40** Profit (7.30% margin)
- **Central**: **547,289.62** Sales | **42,533.18** Profit (**7.77% margin** - highest profitability)
- **East**: **525,124.29** Sales | **38,407.78** Profit (7.31% margin)

### 3. Customer Segments
- **Consumer**: Represents the largest buyer base, accounting for **51.2% (1,115,350.30)** of total sales and **81,762.33** in profit.
- **Corporate**: Accounts for **27.7% (603,849.44)** of total sales and **44,589.10** in profit.
- **Home Office**: Accounts for **21.1% (459,566.33)** of total sales and **35,245.83** in profit.

### 4. Top 5 Revenue-Generating Products
1. **Machines Product 0200**: 4,021.59
2. **Copiers Product 0556**: 3,784.62
3. **Phones Product 0594**: 3,683.74
4. **Copiers Product 0363**: 3,536.75
5. **Phones Product 0712**: 3,488.69

---

## 🛠️ Excel Skills & Functions Demonstrated

1. **Statistical & Mathematical Functions**: `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `SUMIFS`, `COUNTIFS`
2. **Lookup & Reference**: `XLOOKUP`, `VLOOKUP`, `INDEX-MATCH`
3. **Data Summarization**: Pivot Tables with multidimensional row groupings, calculated values, and percentage share representations.
4. **Data Cleaning & Formatting**: Currency formatting, date normalization, error checking (`IFERROR`), and missing-value elimination.

---

## 🐍 Python Verification & Audit

A companion script [`analysis.py`](./analysis.py) is provided to verify the mathematical integrity of the Excel model:

```bash
# Run the audit and verification script
python analysis.py
```

Output:
```text
=================================================================
   VEDA TECHNOLOGY - TASK 14: EXCEL DATA ANALYSIS
=================================================================
[+] Loaded workbook: Veda_Technology_Excel_Data_Analysis_Dataset.xlsx
[+] Available sheets: ['Data', 'Formulas', 'PivotTable']
[+] Total Records: 10,000 | Missing Values: 0
[+] Total Sales  : 2,178,766.07
[+] Total Profit : 161,597.26
[+] Margin       : 7.42%
=================================================================
```

---

## 📁 Repository Structure

```text
veda-technology-task-14-basic/
├── .gitignore                                     # Git ignore rules for Excel lock files & Python cache
├── Veda_Technology_Excel_Data_Analysis_Dataset.xlsx # Core dataset with Data, Formulas & PivotTable sheets
├── analysis.py                                    # Python data audit and verification script
└── README.md                                      # Documentation & executive summary report
```

---

## 👤 Author

- **Intern**: Krishna Kumar ([@KrishnaKr87](https://github.com/KrishnaKr87))
- **Organization**: Veda Technology (VEDA TDC COMPANY)
- **Role**: Data Analytics Intern
- **Repository**: [veda-technology-task-14-basic](https://github.com/KrishnaKr87/veda-technology-task-14-basic)
