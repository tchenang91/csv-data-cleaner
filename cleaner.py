"""CSV Data Cleaner - Automatic cleaning and standardization of CSV files."""
import pandas as pd
import argparse, sys
from pathlib import Path

def clean_csv(filepath, output=None, report=False):
    print(f"\n🧹 CSV Data Cleaner")
    print(f"   Datei: {filepath}")
    
    df = pd.read_csv(filepath, encoding='utf-8', on_bad_lines='warn')
    original_rows = len(df)
    original_cols = len(df.columns)
    print(f"   Geladen: {original_rows} Zeilen, {original_cols} Spalten")
    
    stats = {}
    
    # 1. Remove duplicates
    dupes = df.duplicated().sum()
    df = df.drop_duplicates()
    stats['Duplikate entfernt'] = dupes
    
    # 2. Strip whitespace from string columns
    for col in df.select_dtypes(include='object'):
        df[col] = df[col].str.strip()
    
    # 3. Handle missing values
    null_before = df.isnull().sum().sum()
    for col in df.select_dtypes(include='number'):
        df[col] = df[col].fillna(df[col].median())
    for col in df.select_dtypes(include='object'):
        df[col] = df[col].fillna('N/A')
    stats['Fehlende Werte behandelt'] = null_before
    
    # 4. Standardize column names
    df.columns = [c.strip().lower().replace(' ', '_').replace('-', '_') for c in df.columns]
    
    # 5. Remove empty rows/columns
    empty_cols = [c for c in df.columns if df[c].nunique() <= 1 and df[c].iloc[0] in ['N/A', '', None]]
    df = df.drop(columns=empty_cols)
    stats['Leere Spalten entfernt'] = len(empty_cols)
    
    final_rows = len(df)
    stats['Zeilen vorher'] = original_rows
    stats['Zeilen nachher'] = final_rows
    
    out_path = output or filepath.replace('.csv', '_cleaned.csv')
    df.to_csv(out_path, index=False, encoding='utf-8')
    print(f"   ✓ Gespeichert: {out_path}")
    
    if report:
        print(f"\n   📋 Bereinigungsbericht:")
        for k, v in stats.items():
            print(f"      {k}: {v}")
    
    return df

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='CSV Data Cleaner')
    parser.add_argument('input', help='Input CSV file')
    parser.add_argument('--output', '-o', help='Output file path')
    parser.add_argument('--report', '-r', action='store_true', help='Show cleaning report')
    args = parser.parse_args()
    clean_csv(args.input, args.output, args.report)
