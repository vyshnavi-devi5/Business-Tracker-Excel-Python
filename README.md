# business_tracker.py
# Business Tracker - Sales, Expenses & Profit Analyzer
# Analyzes Excel data to calculate profits, summaries and charts

import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime

# ────────────────────────────────────────────────
# CONFIGURATION
# ────────────────────────────────────────────────
EXCEL_FILE = 'business_tracker.xlsx'

# ────────────────────────────────────────────────
# 1. Welcome & File Check
# ────────────────────────────────────────────────
today = datetime.now().strftime("%d %B %Y")
print("\n" + "="*60)
print(f"     BUSINESS TRACKER REPORT  -  {today}")
print("="*60)

print(f"Current folder: {os.getcwd()}")
print(f"Looking for file: {EXCEL_FILE}")

if not os.path.exists(EXCEL_FILE):
    print(f"\nERROR: File '{EXCEL_FILE}' not found in this folder!")
    print("Please make sure the Excel file is in the same folder as this script.")
    print("Files currently here:", os.listdir('.'))
    exit()

print("File found. Loading data...\n")

# ────────────────────────────────────────────────
# 2. Load Excel sheets (lowercase names as per your file)
# ────────────────────────────────────────────────
try:
    data = pd.read_excel(EXCEL_FILE, sheet_name=None)
    sales_df    = data['sales']
    expenses_df = data['expenses']
    products_df = data['products']
except KeyError as e:
    print(f"ERROR: Sheet {e} not found in the Excel file.")
    print("Available sheets:", list(data.keys()))
    print("Please check sheet names (case-sensitive) in your Excel file.")
    exit()

print(f"Sales rows:    {len(sales_df):>4}")
print(f"Expenses rows: {len(expenses_df):>4}")
print(f"Products rows: {len(products_df):>4}\n")

# ────────────────────────────────────────────────
# 3. Data Preparation
# ────────────────────────────────────────────────
# Convert dates
sales_df['Date']    = pd.to_datetime(sales_df['Date'])
expenses_df['Date'] = pd.to_datetime(expenses_df['Date'])

# Add month column
sales_df['Month']    = sales_df['Date'].dt.to_period('M')
expenses_df['Month'] = expenses_df['Date'].dt.to_period('M')

# Merge cost price into sales
sales_df = sales_df.merge(
    products_df[['Product', 'Cost_Price']],
    on='Product',
    how='left'
)

# Calculate profit per row
sales_df['Cost']  = sales_df['Quantity'] * sales_df['Cost_Price'].fillna(0)
sales_df['Profit'] = sales_df['Total_Sales'] - sales_df['Cost']

print("Sample data with profit calculated (first 5 rows):")
print(sales_df[['Date', 'Product', 'Quantity', 'Price_per_Unit', 'Cost_Price', 'Total_Sales', 'Profit']].head())
print()

# ────────────────────────────────────────────────
# 4. Overall Summary
# ────────────────────────────────────────────────
total_sales    = sales_df['Total_Sales'].sum()
total_cost     = sales_df['Cost'].sum()
gross_profit   = sales_df['Profit'].sum()
total_expenses = expenses_df['Amount'].sum()
net_profit     = gross_profit - total_expenses

print("="*60)
print("          OVERALL SUMMARY")
print("="*60)
print(f"Total Sales            : ₹{total_sales:>12,.0f}")
print(f"Total Cost of Goods    : ₹{total_cost:>12,.0f}")
print(f"Gross Profit           : ₹{gross_profit:>12,.0f}")
print(f"Total Expenses         : ₹{total_expenses:>12,.0f}")
print(f"Net Profit             : ₹{net_profit:>12,.0f}")
print("="*60)
print()

# ────────────────────────────────────────────────
# 5. Monthly Summary
# ────────────────────────────────────────────────
monthly = sales_df.groupby('Month').agg(
    Sales=('Total_Sales', 'sum'),
    Gross_Profit=('Profit', 'sum')
)

monthly_exp = expenses_df.groupby('Month')['Amount'].sum().rename('Expenses')

monthly_summary = monthly.join(monthly_exp, how='outer').fillna(0)
monthly_summary['Net_Profit'] = monthly_summary['Gross_Profit'] - monthly_summary['Expenses']

print("          MONTHLY SUMMARY")
print("="*60)
print(monthly_summary.round(0).astype(int))
print("="*60)
print()

# ────────────────────────────────────────────────
# 6. Charts
# ────────────────────────────────────────────────

# Chart 1: Monthly Net Profit
plt.figure(figsize=(10, 5))
monthly_summary['Net_Profit'].plot(kind='bar', color='royalblue')
plt.title('Monthly Net Profit', fontsize=14, pad=15)
plt.ylabel('Amount (₹)')
plt.xlabel('Month')
plt.xticks(rotation=45)
plt.grid(axis='y', alpha=0.3, linestyle='--')
plt.tight_layout()
plt.savefig('monthly_net_profit.png', dpi=300, bbox_inches='tight')
plt.show()

# Chart 2: Expenses by Type (if column exists)
if 'Expense_Type' in expenses_df.columns:
    exp_by_type = expenses_df.groupby('Expense_Type')['Amount'].sum()
    plt.figure(figsize=(8, 8))
    exp_by_type.plot(kind='pie', autopct='%1.1f%%', startangle=90, shadow=True)
    plt.title('Expenses by Category', fontsize=14, pad=15)
    plt.ylabel('')
    plt.savefig('expenses_by_type.png', dpi=300, bbox_inches='tight')
    plt.show()

# Chart 3: Top Products
top_products = sales_df.groupby('Product')['Total_Sales'].sum().nlargest(8)
plt.figure(figsize=(10, 5))
top_products.plot(kind='bar', color='teal')
plt.title('Top 8 Products by Sales', fontsize=14, pad=15)
plt.ylabel('Total Sales (₹)')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('top_products.png', dpi=300, bbox_inches='tight')
plt.show()

print("\nCharts saved as PNG files in the current folder.")
print("All done! ✓")
