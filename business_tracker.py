# business_tracker.py - FIXED VERSION with lowercase sheet names

import pandas as pd
import matplotlib.pyplot as plt
import os

# File path (already correct in your case)
file_path = 'business_tracker.xlsx'

print("Current working folder:", os.getcwd())
print("Files in this folder:", os.listdir('.'))
print(f"\nTrying to read file: {file_path}")

if not os.path.exists(file_path):
    print(f"\nERROR: File '{file_path}' NOT FOUND!")
    exit()

print("File found! Loading data...\n")

# Load all sheets
data = pd.read_excel(file_path, sheet_name=None)

print("Sheets found in Excel:", list(data.keys()))

# Use the ACTUAL lowercase sheet names
try:
    sales_df    = data['sales']
    expenses_df = data['expenses']
    products_df = data['products']
except KeyError as e:
    print(f"\nERROR: Sheet {e} still not found! Double-check sheet names.")
    exit()

print(f"\nSales rows: {sales_df.shape[0]}")
print(f"Expenses rows: {expenses_df.shape[0]}")
print(f"Products rows: {products_df.shape[0]}")

# ── Data Preparation ────────────────────────────────────────────────
sales_df['Date'] = pd.to_datetime(sales_df['Date'])
expenses_df['Date'] = pd.to_datetime(expenses_df['Date'])

sales_df['Month'] = sales_df['Date'].dt.to_period('M')
expenses_df['Month'] = expenses_df['Date'].dt.to_period('M')

# Merge Cost_Price from products sheet (matching on 'Product' column)
sales_df = sales_df.merge(
    products_df[['Product', 'Cost_Price']],
    on='Product',
    how='left'
)

# Calculate profit per row
sales_df['Profit'] = sales_df['Total_Sales'] - (sales_df['Quantity'] * sales_df['Cost_Price'].fillna(0))

print("\nSales with Profit (first 5 rows):")
print(sales_df[['Date', 'Product', 'Quantity', 'Price_per_Unit', 'Cost_Price', 'Total_Sales', 'Profit']].head())

# ── Overall Summary ─────────────────────────────────────────────────
total_sales    = sales_df['Total_Sales'].sum()
total_cost     = (sales_df['Quantity'] * sales_df['Cost_Price'].fillna(0)).sum()
gross_profit   = sales_df['Profit'].sum()
total_expenses = expenses_df['Amount'].sum()
net_profit     = gross_profit - total_expenses

print("\n" + "="*50)
print("          OVERALL BUSINESS SUMMARY")
print("="*50)
print(f"Total Sales            : ₹{total_sales:,.0f}")
print(f"Total Cost of Goods    : ₹{total_cost:,.0f}")
print(f"Gross Profit           : ₹{gross_profit:,.0f}")
print(f"Total Expenses         : ₹{total_expenses:,.0f}")
print(f"Net Profit             : ₹{net_profit:,.0f}")
print("="*50)

# ── Monthly Summary ─────────────────────────────────────────────────
monthly = sales_df.groupby('Month').agg(
    Sales=('Total_Sales', 'sum'),
    Gross_Profit=('Profit', 'sum')
)

monthly_exp = expenses_df.groupby('Month')['Amount'].sum().rename('Expenses')

monthly_summary = monthly.join(monthly_exp, how='outer').fillna(0)
monthly_summary['Net_Profit'] = monthly_summary['Gross_Profit'] - monthly_summary['Expenses']

print("\n" + "="*50)
print("            MONTHLY SUMMARY")
print("="*50)
print(monthly_summary)
print("="*50)

# ── Charts ──────────────────────────────────────────────────────────
# 1. Monthly Net Profit Bar Chart
plt.figure(figsize=(10, 5))
monthly_summary['Net_Profit'].plot(kind='bar', color='cornflowerblue')
plt.title('Monthly Net Profit', fontsize=14)
plt.ylabel('Amount (₹)')
plt.xlabel('Month')
plt.xticks(rotation=45)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.show()

# 2. Expenses by Type (Pie Chart)
if 'Expense_Type' in expenses_df.columns:
    exp_by_type = expenses_df.groupby('Expense_Type')['Amount'].sum()
    plt.figure(figsize=(8, 8))
    exp_by_type.plot(kind='pie', autopct='%1.1f%%', startangle=90)
    plt.title('Expenses by Type')
    plt.ylabel('')
    plt.show()

# 3. Top Products by Sales
top_prod = sales_df.groupby('Product')['Total_Sales'].sum().nlargest(8)
plt.figure(figsize=(10, 5))
top_prod.plot(kind='bar', color='teal')
plt.title('Top Products by Sales')
plt.ylabel('Total Sales (₹)')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

print("\nDone! Check the printed numbers and pop-up charts.")
