#!/usr/bin/env python3
"""
Script to analyze American Express statement data.
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from datetime import datetime

def analyze_statement(filepath):
    """Analyze an AMEX statement Excel file and generate reports."""
    print(f"Analyzing file: {filepath}")
    
    # Create a directory for the analysis
    analysis_dir = os.path.join(os.path.dirname(filepath), 'analysis')
    os.makedirs(analysis_dir, exist_ok=True)
    
    # Load the data
    df = pd.read_excel(filepath)
    
    # Find actual transaction rows (skip header info)
    df = df[df['Transaction Details'].str.contains(r'^\d{2}/\d{2}/\d{4}$', na=False)]
    
    # Clean the data
    df.rename(columns={'Transaction Details': 'Date', 
                      'Business Gold Card / May 06, 2025 to Jun 06, 2025': 'Receipt',
                      'Unnamed: 2': 'Description',
                      'Unnamed: 3': 'Amount',
                      'Unnamed: 4': 'Extended_Details',
                      'Unnamed: 5': 'Statement_Description',
                      'Unnamed: 6': 'Address',
                      'Unnamed: 7': 'City_State',
                      'Unnamed: 8': 'Zip',
                      'Unnamed: 9': 'Country',
                      'Unnamed: 10': 'Reference',
                      'Unnamed: 11': 'Category'}, inplace=True)
    
    # Convert date to datetime
    df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')
    
    # Replace NaN in Country with 'N/A'
    df['Country'] = df['Country'].fillna('N/A')
    
    # Split the Category field into main category and subcategory
    df['Main_Category'] = df['Category'].str.split('-').str[0]
    df['Sub_Category'] = df['Category'].str.split('-').str[1]
    
    # Replace NaN with 'Uncategorized'
    df['Main_Category'] = df['Main_Category'].fillna('Uncategorized')
    df['Sub_Category'] = df['Sub_Category'].fillna('Uncategorized')
    
    # Convert Amount to numeric
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    
    # Filter out negative amounts (payments/credits) for spending analysis
    spending_df = df[df['Amount'] > 0]
    
    # Save filtered data to CSV for reference
    df.to_csv(os.path.join(analysis_dir, 'all_transactions.csv'), index=False)
    spending_df.to_csv(os.path.join(analysis_dir, 'spending_only.csv'), index=False)
    
    # Country spending breakdown
    country_spending = spending_df.groupby('Country')['Amount'].agg(['sum', 'count']).reset_index()
    country_spending.columns = ['Country', 'Total_Amount', 'Transaction_Count']
    country_spending.to_csv(os.path.join(analysis_dir, 'country_spending.csv'), index=False)
    
    # Category spending breakdown
    category_spending = spending_df.groupby('Main_Category')['Amount'].agg(['sum', 'count']).reset_index()
    category_spending.columns = ['Category', 'Total_Amount', 'Transaction_Count'] 
    category_spending.to_csv(os.path.join(analysis_dir, 'category_spending.csv'), index=False)
    
    # Country and Category combined
    country_cat_spending = spending_df.groupby(['Country', 'Main_Category'])['Amount'].agg(['sum', 'count']).reset_index()
    country_cat_spending.columns = ['Country', 'Category', 'Total_Amount', 'Transaction_Count']
    country_cat_spending.to_csv(os.path.join(analysis_dir, 'country_category_spending.csv'), index=False)
    
    # Generate visualizations
    plt.figure(figsize=(10, 6))
    cs_sorted = country_spending.sort_values('Total_Amount', ascending=False)
    plt.bar(cs_sorted['Country'], cs_sorted['Total_Amount'], color='skyblue')
    plt.title('Spending by Country')
    plt.ylabel('Amount ($)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(analysis_dir, 'country_spending.png'))
    
    plt.figure(figsize=(10, 6))
    cat_sorted = category_spending.sort_values('Total_Amount', ascending=False)
    plt.bar(cat_sorted['Category'], cat_sorted['Total_Amount'], color='lightgreen')
    plt.title('Spending by Category')
    plt.ylabel('Amount ($)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(analysis_dir, 'category_spending.png'))
    
    # Generate text report
    with open(os.path.join(analysis_dir, 'summary_report.txt'), 'w') as f:
        f.write("Business Gold Card Transaction Analysis\n")
        f.write("====================================\n\n")
        f.write(f"Statement Period: {df['Date'].min().strftime('%B %d, %Y')} to {df['Date'].max().strftime('%B %d, %Y')}\n\n")
        f.write(f"Total Transactions: {len(df)}\n")
        f.write(f"Total Spent: ${spending_df['Amount'].sum():.2f}\n")
        f.write(f"Total Payments/Credits: ${df[df['Amount'] < 0]['Amount'].sum() * -1:.2f}\n")
        f.write(f"Net Balance: ${df['Amount'].sum():.2f}\n\n")
        
        f.write("Spending by Country\n")
        f.write("-----------------\n")
        for _, row in cs_sorted.iterrows():
            f.write(f"{row['Country']}: ${row['Total_Amount']:.2f} ({row['Transaction_Count']} transactions)\n")
        
        f.write("\nSpending by Category\n")
        f.write("------------------\n")
        for _, row in cat_sorted.iterrows():
            f.write(f"{row['Category']}: ${row['Total_Amount']:.2f} ({row['Transaction_Count']} transactions)\n")
        
        f.write("\nDetailed Country/Category Breakdown\n")
        f.write("--------------------------------\n")
        cc_sorted = country_cat_spending.sort_values(['Country', 'Total_Amount'], ascending=[True, False])
        current_country = None
        for _, row in cc_sorted.iterrows():
            if current_country != row['Country']:
                current_country = row['Country']
                f.write(f"\n{current_country}\n")
            f.write(f"  {row['Category']}: ${row['Total_Amount']:.2f} ({row['Transaction_Count']} transactions)\n")
    
    # Print summary to console
    print("\nAnalysis complete! Files saved to:", analysis_dir)
    print("\nSummary:")
    print(f"Total Spent: ${spending_df['Amount'].sum():.2f}")
    print(f"Total Payments/Credits: ${df[df['Amount'] < 0]['Amount'].sum() * -1:.2f}")
    print(f"Net Balance: ${df['Amount'].sum():.2f}")
    
    print("\nTop Countries by Spending:")
    print(cs_sorted.head().to_string(index=False))
    
    print("\nTop Categories by Spending:")
    print(cat_sorted.head().to_string(index=False))
    
    return analysis_dir

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Analyze American Express statement data")
    parser.add_argument("filepath", help="Path to the Excel file containing statement data")
    
    args = parser.parse_args()
    
    analyze_statement(args.filepath)