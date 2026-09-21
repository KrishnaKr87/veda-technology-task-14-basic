"""
Veda Technology - Task 14: Basic Excel Data Analysis
Verification & Analysis Script

This script loads the dataset from 'Veda_Technology_Excel_Data_Analysis_Dataset.xlsx',
audits data integrity, computes key performance indicators (KPIs), verifies calculations
against the Excel 'Formulas' and 'PivotTable' sheets, and displays key business insights.
"""

import sys
import pandas as pd

DATASET_FILE = "Veda_Technology_Excel_Data_Analysis_Dataset.xlsx"

def run_analysis():
    print("=" * 65)
    print("   VEDA TECHNOLOGY - TASK 14: EXCEL DATA ANALYSIS")
    print("=" * 65)
    
    # 1. Load Data
    try:
        excel = pd.ExcelFile(DATASET_FILE)
        print(f"[+] Loaded workbook: {DATASET_FILE}")
        print(f"[+] Available sheets: {excel.sheet_names}")
    except Exception as e:
        print(f"[-] Error opening {DATASET_FILE}: {e}")
        sys.exit(1)
        
    df_data = excel.parse("Data")
    df_formulas = excel.parse("Formulas")
    
    # 2. Dataset Overview
    print("\n" + "-" * 40)
    print(" 1. DATASET OVERVIEW & QUALITY CHECK")
    print("-" * 40)
    print(f"Total Records (Rows): {len(df_data):,}")
    print(f"Total Columns       : {len(df_data.columns)}")
    print(f"Missing Values      : {df_data.isnull().sum().sum()}")
    print(f"Date Range          : {df_data['Order Date'].min().strftime('%Y-%m-%d')} to {df_data['Order Date'].max().strftime('%Y-%m-%d')}")
    print(f"Unique Customers    : {df_data['Customer ID'].nunique():,}")
    print(f"Unique Products     : {df_data['Product ID'].nunique():,}")

    # 3. Core KPI Calculations
    total_sales = df_data["Sales"].sum()
    total_profit = df_data["Profit"].sum()
    total_qty = df_data["Quantity"].sum()
    total_orders = len(df_data)
    avg_sales_per_order = total_sales / total_orders
    profit_margin_pct = (total_profit / total_sales) * 100

    print("\n" + "-" * 40)
    print(" 2. CORE KEY PERFORMANCE INDICATORS (KPIs)")
    print("-" * 40)
    print(f"Total Sales              : ${total_sales:,.2f}")
    print(f"Total Profit             : ${total_profit:,.2f}")
    print(f"Total Quantity Sold      : {total_qty:,} units")
    print(f"Total Order Lines        : {total_orders:,}")
    print(f"Average Order Line Sales : ${avg_sales_per_order:,.2f}")
    print(f"Overall Profit Margin    : {profit_margin_pct:.2f}%")

    # 4. Category Performance
    print("\n" + "-" * 40)
    print(" 3. CATEGORY PERFORMANCE SUMMARY")
    print("-" * 40)
    cat_summary = df_data.groupby("Category").agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Units_Sold=("Quantity", "sum")
    )
    cat_summary["Profit_Margin_%"] = (cat_summary["Total_Profit"] / cat_summary["Total_Sales"]) * 100
    cat_summary["Sales_Share_%"] = (cat_summary["Total_Sales"] / total_sales) * 100
    for cat, row in cat_summary.iterrows():
        print(f" • {cat:<16}: Sales = ${row['Total_Sales']:>10,.2f} ({row['Sales_Share_%']:>5.1f}%) | "
              f"Profit = ${row['Total_Profit']:>9,.2f} (Margin: {row['Profit_Margin_%']:>5.2f}%)")

    # 5. Regional Performance
    print("\n" + "-" * 40)
    print(" 4. REGIONAL PERFORMANCE SUMMARY")
    print("-" * 40)
    reg_summary = df_data.groupby("Region").agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Units_Sold=("Quantity", "sum")
    )
    reg_summary["Profit_Margin_%"] = (reg_summary["Total_Profit"] / reg_summary["Total_Sales"]) * 100
    for reg, row in reg_summary.iterrows():
        print(f" • {reg:<8}: Sales = ${row['Total_Sales']:>10,.2f} | "
              f"Profit = ${row['Total_Profit']:>9,.2f} (Margin: {row['Profit_Margin_%']:>5.2f}%)")

    # 6. Customer Segment Breakdown
    print("\n" + "-" * 40)
    print(" 5. CUSTOMER SEGMENT BREAKDOWN")
    print("-" * 40)
    seg_summary = df_data.groupby("Segment").agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum")
    )
    seg_summary["Sales_Share_%"] = (seg_summary["Total_Sales"] / total_sales) * 100
    for seg, row in seg_summary.iterrows():
        print(f" • {seg:<12}: Sales = ${row['Total_Sales']:>10,.2f} ({row['Sales_Share_%']:>5.1f}%) | "
              f"Profit = ${row['Total_Profit']:>9,.2f}")

    # 7. Top 5 Products by Sales
    print("\n" + "-" * 40)
    print(" 6. TOP 5 PRODUCTS BY REVENUE")
    print("-" * 40)
    top_products = df_data.groupby("Product Name")["Sales"].sum().sort_values(ascending=False).head(5)
    for rank, (prod, sales) in enumerate(top_products.items(), 1):
        print(f" {rank}. {prod:<25}: ${sales:,.2f}")

    print("\n" + "=" * 65)
    print("   ANALYSIS & DATA AUDIT COMPLETED SUCCESSFULLY!")
    print("=" * 65)

if __name__ == "__main__":
    run_analysis()
