import pandas as pd
import numpy as np
import re

# LOAD DATASET
file_path = "Admissions_register cleaned.xlsx"  # Corrected path based on available file

df = pd.read_excel(file_path)  # Load the Excel file into a DataFrame

# Explore the dataset
print("Data loaded successfully. Shape:", df.shape)
print("Columns:", list(df.columns))
print("First 5 rows:")
print(df.head())

# Data Cleansing
print("\n--- Data Cleansing ---")

# handle missing values
numerical_cols = ['program_duration', 'length_of_stay_days']
categorical_cols = ['education_level', 'ethnicity', 'referral_source', 'primary_substance_of_use', 'second_most_frequently_used_substance', 'sources_of_payment', 'employment']

for col in numerical_cols:
    if col in df.columns:
        median_val = df[col].median()
        df[col].fillna(median_val, inplace=True)
        print(f"Filled missing values in {col} with median: {median_val}")

for col in categorical_cols:
    if col in df.columns:
        mode_val = df[col].mode()
        if not mode_val.empty:
            df[col].fillna(mode_val[0], inplace=True)
            print(f"Filled missing values in {col} with mode: {mode_val[0]}")
        else:
            df[col].fillna('Unknown', inplace=True)
            print(f"Filled missing values in {col} with 'Unknown'")

# Convert date columns to datetime
date_cols = ['month', 'date', 'discharge_date']
for col in date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        print(f"Converted {col} to datetime")

# Check for duplicates
duplicates = df.duplicated().sum()
if duplicates > 0:
    df.drop_duplicates(inplace=True)
    print(f"Removed {duplicates} duplicate rows")
else:
    print("No duplicate rows found")

# Ensure data types
df['program_duration'] = df['program_duration'].astype(int, errors='ignore')
df['length_of_stay_days'] = df['length_of_stay_days'].astype(float, errors='ignore')
print("Ensured appropriate data types")

print("Cleaned data shape:", df.shape)

# Further Analysis
print("\n--- Further Analysis ---")

# Summary statistics for numerical columns
numerical_summary = df[numerical_cols].describe()
print("Summary statistics for numerical columns:")
print(numerical_summary)

# Value counts for categorical columns
for col in categorical_cols:
    if col in df.columns:
        print(f"\nValue counts for {col}:")
        print(df[col].value_counts())

# Trends: Admissions by year, month, gender, ethnicity
if 'admission_year' in df.columns:
    print("\nAdmissions by year:")
    print(df['admission_year'].value_counts().sort_index())

if 'admission_month' in df.columns:
    print("\nAdmissions by month:")
    print(df['admission_month'].value_counts())

if 'gender' in df.columns:
    print("\nAdmissions by gender:")
    print(df['gender'].value_counts())

if 'ethnicity' in df.columns:
    print("\nAdmissions by ethnicity:")
    print(df['ethnicity'].value_counts())

# Correlations
correlation_cols = numerical_cols + ['program_duration']  # Add more if needed
correlation_matrix = df[correlation_cols].corr()
print("\nCorrelation matrix:")
print(correlation_matrix)

# Save cleaned data
df.to_excel('cleaned_admissions.xlsx', index=False)
print("\nCleaned data saved to 'cleaned_admissions.xlsx'")

# Save analysis results to text file
with open('analysis_summary.txt', 'w') as f:
    f.write("Summary Statistics:\n")
    f.write(str(numerical_summary))
    f.write("\n\nValue Counts:\n")
    
    for col in categorical_cols:
        if col in df.columns:
            f.write(f"\n{col}:\n")
            f.write(str(df[col].value_counts()))
    f.write("\n\nCorrelations:\n")
    f.write(str(correlation_matrix))
print("Analysis summary saved to 'analysis_summary.txt'")

# R

