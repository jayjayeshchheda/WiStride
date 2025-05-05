# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 11:17:02 2025

@author: Jay Jayesh Chheda

Title: Exim Data Cleaner
"""

# %%

"""
Stage-1: Cleaning Exim Data
"""

import pandas as pd
import numpy as np
import statsmodels.api as sm
import matplotlib.pyplot as plt
import shutil
import os  # Helps work with file paths
import glob  # Helps to get all files matching a pattern

# write the path of the folder containing input files
source_folder = "D:\Python\EXIM_Cleaner"

output_path = os.path.join(source_folder, "Final Output.xlsx")

# Delete the file if it already exists
if os.path.exists(output_path):
    os.remove(output_path)
    print("Deleted file: Final Output.xlsx")

try:
    # Loop through items in the directory
    for item in os.listdir(source_folder):
        item_path = os.path.join(source_folder, item)
        # Check if the item is a directory
        if os.path.isdir(item_path):
            shutil.rmtree(item_path)
            print(f"Deleted folder: {item_path}")
except Exception as e:
    print(f"Error: {e}")

# Get a list of all Excel files in that folder
file_list = glob.glob(os.path.join(source_folder, "*.xlsx"))

# Words to filter from 'PRODUCT' column
unwanted_words = [
    'PALLET', 'PE L L E T S', 'PELLE', 'STERIC ACID', 'SAMPLE', 'FREE',
    'GRANULES', 'IMPURITIES', 'STAGE'
]

pattern = '|'.join(unwanted_words)  # Create pattern for filtering

# Loop through each file
for file_path in file_list:
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # replace spaces with underscores in the column headers
    df.columns = df.columns.str.replace(' ', '_')
    
    # Convert PRODUCT column to uppercase (for consistent matching)
    df['PRODUCT'] = df['PRODUCT'].astype(str).str.upper()
    
    # Get number of rows before filtering
    rows_before = len(df)

    # Convert to numeric (ignore non-numeric values)
    df['QUANTITY'] = pd.to_numeric(df['QUANTITY'], errors='coerce')

    # Convert to numeric (ignore non-numeric values)
    df['UNIT_RATE'] = pd.to_numeric(df['UNIT_RATE'], errors='coerce')
    
    """ Drop rows with NaN in QUANTITY or QUANTITY <= 0"""
    df = df.dropna(subset=['QUANTITY'])
    df = df[df['QUANTITY'] > 0]

    """ Drop rows with NaN in UNIT_RATE or UNIT_RATE <= 0"""
    df = df.dropna(subset=['UNIT_RATE'])
    df = df[df['UNIT_RATE'] > 0]

    """ Unit Conversion of GMS & TON to KGS and removal of other units"""

    # Convert quantities from grams (GMS) to kilograms (KGS)
    df.loc[df['UNIT'] == 'GMS', 'QUANTITY'] = df.loc[df['UNIT'] == 'GMS', 'QUANTITY'] / 1000
    
    # Since we are dividing quantity by 1000, we must multiply the unit rate by 1000 to maintain correctness
    df.loc[df['UNIT'] == 'GMS', 'UNIT_RATE'] = df.loc[df['UNIT'] == 'GMS', 'UNIT_RATE'] * 1000
    
    # Update the unit from GMS to KGS
    df.loc[df['UNIT'] == 'GMS', 'UNIT'] = 'KGS'
        
    # Convert quantities from tons (TON) to kilograms (KGS)
    df.loc[df['UNIT'] == 'TON', 'QUANTITY'] = df.loc[df['UNIT'] == 'TON', 'QUANTITY'] * 1000
    
    # Since we're multiplying quantity by 1000, we divide the unit rate by 1000 to keep it consistent
    df.loc[df['UNIT'] == 'TON', 'UNIT_RATE'] = df.loc[df['UNIT'] == 'TON', 'UNIT_RATE'] / 1000
    
    # Update the unit from TON to KGS
    df.loc[df['UNIT'] == 'TON', 'UNIT'] = 'KGS'
        
    # Finally, keep only the rows where the unit is now KGS & currency is USD (remove any other units)
    df = df[df['UNIT'] == 'KGS']    
    df = df[df['CURRENCY'] == 'USD']
    
    """ Remove rows that contain any of the unwanted words"""
        
    df_filtered = df[~df['PRODUCT'].str.contains(pattern, na=False)]
    
    region = pd.read_csv(r"D:\WisOnGo FTP Server\Google Drive\G-Drive\OneDrive\Python_Data\default_python_directory\Files in use\updated_region.csv")
        
    # Check if 'origin' column exists in df_filtered_exim and rename 'destination' column if it does
    if 'ORIGIN' in df_filtered.columns:
        region.rename(columns={'destination': 'ORIGIN'}, inplace=True)
        region['ORIGIN'] = region['ORIGIN'].str.upper()
        df_filtered.loc[:, 'ORIGIN'] = df_filtered['ORIGIN'].str.upper()
        df_filtered = pd.merge(df_filtered, region, on='ORIGIN', how='left', suffixes=('', '_new'))
    else:
        region.rename(columns={'destination': 'DESTINATION'}, inplace=True)
        region['DESTINATION'] = region['DESTINATION'].str.upper()
        df_filtered.loc[:, 'DESTINATION'] = df_filtered['DESTINATION'].str.upper()
        df_filtered = pd.merge(df_filtered, region, on='DESTINATION', how='left', suffixes=('', '_new'))
        
    # Get number of rows after filtering
    rows_after = len(df_filtered)

    # Get the directory where the original file is stored
    base_dir = os.path.dirname(file_path)

    # Create a new folder named 'Stage-1_filtered_files' inside the source directory
    output_dir = os.path.join(base_dir, "Stage-1_filtered_files")
    os.makedirs(output_dir, exist_ok=True)  # Create the folder if it doesn't exist    
    
    # Create output file name by adding 'filtered_' prefix
    output_path = os.path.join(output_dir, "filtered_" + os.path.basename(file_path))

    # Save the filtered data
    df_filtered.to_excel(output_path, index=False)
    
    print(f"File: {os.path.basename(file_path)}")
    print(f"Line items before filtering: {rows_before}")
    print(f"Line items after filtering:  {rows_after}")
    print(f"Processed and saved: {output_path}\n")
# %%
    
"""
Stage-2: Volume Cleaning
"""

# Define the folder containing filtered files
filtered_folder = "D:\Python\EXIM_Cleaner\Stage-1_filtered_files"

# Get a list of all Excel files in that folder
filtered_files = glob.glob(os.path.join(filtered_folder, "*.xlsx"))

# Function to select consecutive bins within 30% difference
def select_consecutive_bins_within_30_percent(counts):
    max_val = counts.max()
    max_idx = counts.idxmax()
    threshold = 0.7 * max_val  # 70% of the highest value

    bins_within_threshold = counts[counts >= threshold].index.tolist()
    selected_bins = [max_idx]

    max_pos = counts.index.get_loc(max_idx)

    # Check if previous bin is also within threshold
    if max_pos > 0:
        prev_bin = counts.index[max_pos - 1]
        if prev_bin in bins_within_threshold:
            selected_bins.insert(0, prev_bin)

    # Check if next bin is also within threshold
    if max_pos < len(counts) - 1:
        next_bin = counts.index[max_pos + 1]
        if next_bin in bins_within_threshold:
            selected_bins.append(next_bin)

    return selected_bins

# Loop through each filtered file
for file_path in filtered_files:
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Check if 'QUANTITY' column exists and is numeric
    if 'QUANTITY' in df.columns:
        # Convert the column to numeric (if it's not already), ignore non-numeric safely
        df['QUANTITY'] = pd.to_numeric(df['QUANTITY'], errors='coerce')

        # Drop rows with missing (NaN) values in QUANTITY
        df_clean = df.dropna(subset=['QUANTITY'])

        # Get min and max values
        min_qty = df_clean['QUANTITY'].min()
        max_qty = df_clean['QUANTITY'].max()
        
        # Start with the minimum value
        min_bin = 0.000001

        # Round the maximum value up to the nearest higher multiple of 10
        max_bin = 10 ** np.ceil(np.log10(max_qty))
        
        # Create bins by multiplying by 10 until it reaches or exceeds max_bin
        bins = []
        current_bin = min_bin
        
        while current_bin < max_bin:
            bins.append(current_bin)
            current_bin *= 10        
        
        # Append the rounded maximum value as the last bin
        bins.append(max_bin)
        
        # Create labels for the bins with increased precision
        labels = [f'{bins[i]:.6f} - {bins[i+1]:.6f}' for i in range(len(bins) - 1)]
        
        # Assign Quantity_Category to df
        df['QUANTITY_RANGE'] = pd.cut(df['QUANTITY'], bins=bins, labels=labels, right=False)

        # Count how many entries fall into each bin
        range_counts = df['QUANTITY_RANGE'].value_counts().sort_index()        
        
        # ✅ Call your function to select important bins
        selected_bins = select_consecutive_bins_within_30_percent(range_counts)
        
        # Get number of rows before filtering
        rows_before_binning = len(df)
        
         # Get the minimum bin threshold from selected bins
        if selected_bins:
            min_selected_bin_value = min([float(str(interval).split(" - ")[0]) for interval in selected_bins])
        
            # Filter the DataFrame to keep only rows with QUANTITY >= that threshold
            df = df[df['QUANTITY'] >= min_selected_bin_value]
        
            print(f"Filtered rows below quantity: {min_selected_bin_value}")
        else:
            print("No selected bins found. Skipping filtering based on selected bins.")
        
        # Get number of rows after filtering
        rows_after_binning = len(df)
        
        # Create a new folder named 'Stage-2_volume_cleaned' inside the source directory
        volume_cleaned_folder = os.path.join(source_folder, "Stage-2_volume_cleaned")
        os.makedirs(volume_cleaned_folder, exist_ok=True)  # Create the folder if it doesn't exist

        # Create output file name by adding 'volume_selected_' prefix
        output_path = os.path.join(volume_cleaned_folder, "volume_selected_" + os.path.basename(file_path))
        df.to_excel(output_path, index=False)
        print(f"File: {os.path.basename(file_path)}")
        print(f"Min Quantity: {min_qty}")
        print(f"Max Quantity: {max_qty}")
        print(f"Selected Bins: {selected_bins}")
        print(f"Line items before volume binning: {rows_before_binning}")
        print(f"Line items after volume binning:  {rows_after_binning}")
        print(f"Saved filtered file to: {output_path}\n")        
    else:
        print(f"File: {os.path.basename(file_path)} - 'QUANTITY' column not found!\n")
# %%
    
"""
Stage-3: Value Cleaning
"""

# Get a list of all Excel files in volume_cleaned_folder
volume_cleaned_files = glob.glob(os.path.join(volume_cleaned_folder, "*.xlsx"))

# Loop through each filtered file
for file_path in volume_cleaned_files:
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Check if 'UNIT_RATE' column exists and is numeric
    if 'UNIT_RATE' in df.columns:
        # Convert the column to numeric (if it's not already), ignore non-numeric safely
        df['UNIT_RATE'] = pd.to_numeric(df['UNIT_RATE'], errors='coerce')

        # Drop rows with missing (NaN) values in UNIT_RATE
        df_clean = df.dropna(subset=['UNIT_RATE'])

        # Get min and max values
        min_ur = df_clean['UNIT_RATE'].min()
        max_ur = df_clean['UNIT_RATE'].max()
        
        # Start with the minimum value
        min_bin = 0.000001

        # Round the maximum value up to the nearest higher multiple of 10
        max_bin = 10 ** np.ceil(np.log10(max_ur))
        
        # Create bins by multiplying by 10 until it reaches or exceeds max_bin
        bins = []
        current_bin = min_bin
        
        while current_bin < max_bin:
            bins.append(current_bin)
            current_bin *= 10        
        
        # Append the rounded maximum value as the last bin
        bins.append(max_bin)
        
        # Create labels for the bins with increased precision
        labels = [f'{bins[i]:.6f} - {bins[i+1]:.6f}' for i in range(len(bins) - 1)]
        
        # Assign UNIT_RATE_Category to df
        df['UNIT_RATE_RANGE'] = pd.cut(df['UNIT_RATE'], bins=bins, labels=labels, right=False)

        # Count how many entries fall into each bin
        range_counts = df['UNIT_RATE_RANGE'].value_counts().sort_index()        
        
        # ✅ Call your function to select important bins
        selected_bins = select_consecutive_bins_within_30_percent(range_counts)
        
        # Get number of rows before filtering
        rows_before_binning = len(df)
        
         # Get the minimum bin threshold from selected bins
        if selected_bins:
            min_selected_bin_value = min([float(str(interval).split(" - ")[0]) for interval in selected_bins]) * 0.80
        
            # Filter the DataFrame to keep only rows with UNIT_RATE >= that threshold
            df = df[df['UNIT_RATE'] >= min_selected_bin_value]
            
            # remove outliers using iqr
            Q1 = df['UNIT_RATE'].quantile(0.25)
            Q3 = df['UNIT_RATE'].quantile(0.75)
            IQR = Q3 - Q1

            # Define bounds for outliers
            if ((Q3+IQR) >= 1.5*Q3):
                upper_bound = Q3*1.25
            else:
                upper_bound = Q3 + IQR

            # Apply IQR
            df = df[(df['UNIT_RATE'] <= upper_bound)]
            
            # Get min and max values
            min_ur_f = df['UNIT_RATE'].min()
            max_ur_f = df['UNIT_RATE'].max()
    
            print(f"Filtered rows below UNIT_RATE: {min_selected_bin_value}")
        else:
            print("No selected bins found. Skipping filtering based on selected bins.")
        
        # Get number of rows after filtering
        rows_after_binning = len(df)
        
        # Create a new folder named 'Stage-3_value_cleaned' inside the source directory
        value_cleaned_folder = os.path.join(source_folder, "Stage-3_value_cleaned")
        os.makedirs(value_cleaned_folder, exist_ok=True)  # Create the folder if it doesn't exist

        # Create output file name by adding 'value_selected_' prefix
        output_path = os.path.join(value_cleaned_folder, "value_selected_" + os.path.basename(file_path))
        df.to_excel(output_path, index=False)
        print(f"File: {os.path.basename(file_path)}")
        print(f"Min UNIT_RATE: {min_ur_f}")
        print(f"Max UNIT_RATE: {max_ur_f}")
        print(f"Selected Bins: {selected_bins}")
        print(f"Line items before value binning: {rows_before_binning}")
        print(f"Line items after value binning:  {rows_after_binning}")
        print(f"Saved filtered file to: {output_path}\n")        
    else:
        print(f"File: {os.path.basename(file_path)} - 'UNIT_RATE' column not found!\n")
# %%
    
"""
Stage-4: Data Presentation
"""

# Get a list of all Excel files in value_cleaned_folder
value_cleaned_files = glob.glob(os.path.join(value_cleaned_folder, "*.xlsx"))

# Loop through each filtered file
for file_path in value_cleaned_files:
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Make sure DATE is in datetime format
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    
    # Get index of the lowest UNIT_RATE for each DATE
    lowest_rate_indices = df.groupby('DATE')['UNIT_RATE'].idxmin()
    
    # Filter the dataframe to keep only those rows
    df = df.loc[lowest_rate_indices]
    
    # Sort by DATE and reset index
    df = df.sort_values(by='DATE', ascending=True)
    df = df.reset_index(drop=True)
       
    fraction = 0
    if len(df) > 500:
        fraction = 0.10
    elif len(df) > 250:
        fraction = 0.15
    elif len(df) > 100:
        fraction = 0.25
    else:
        fraction = 0.50
        
    # Convert dates to numeric format (like Unix timestamp) for LOWESS
    x_values = df.index
    # x_values = pd.to_datetime(df['DATE']).astype('int64') // 10**9  # Convert to seconds
    y_values = df['UNIT_RATE']
    
    # Apply LOWESS
    lowess_result = sm.nonparametric.lowess(y_values, x_values, frac=fraction)
    df_lowess = pd.DataFrame(lowess_result)
    df['LOWESS_VALUE'] = df_lowess[1]
    
    # Extract year and month from the DATE column
    df['YEAR'] = df['DATE'].dt.year
    df['MONTH'] = df['DATE'].dt.month
    
    # Group by YEAR and MONTH, then calculate the mean of monthly average
    monthly_avg = df.groupby(['YEAR', 'MONTH'])['UNIT_RATE'].mean().reset_index()
    monthly_avg_mean = monthly_avg['UNIT_RATE'].mean()
    
    lowess_mean = df_lowess[1].mean()
    
    ratio = monthly_avg_mean/lowess_mean
    
    df['MODIFIED_LOWESS'] = df_lowess[1]*ratio
    
    # Calculate the modified mean from the modified new results
    modified_mean = monthly_avg_mean
    # Calculate the squared differences from the modified mean
    squared_differences = (df['UNIT_RATE'] - modified_mean) ** 2
    # Calculate the variance using the squared differences
    variance = squared_differences.sum() / len(df)
    # Calculate the standard deviation as the square root of the variance
    modified_std_dev = variance ** 0.5
    
    # Calculate the upper and lower bounds for outlier detection
    df['UPPER_BOUND'] = df['MODIFIED_LOWESS'].apply(lambda x: 1.25*x if x + modified_std_dev >= 1.25*x else x + modified_std_dev)
    df['LOWER_BOUND'] = df['MODIFIED_LOWESS'].apply(lambda x: 0.75*x if x - modified_std_dev <= 0.75*x else x - modified_std_dev)
    
    df = df[(df['UNIT_RATE'] <= df['UPPER_BOUND']) & (df['UNIT_RATE'] >= df['LOWER_BOUND'])]
    
    # Set the figure size
    plt.figure(figsize=(14, 6))
    
    # Scatter plot for UNIT_RATE
    plt.scatter(df['DATE'], df['UNIT_RATE'], color='#0639EF', alpha=0.5, label='UNIT_RATE')
    
    # Line plot for MODIFIED_LOWESS
    plt.plot(df['DATE'], df['MODIFIED_LOWESS'], color='#DE9954', linewidth=2, label='MODIFIED_LOWESS')
    
    # Line plot for UPPER_BOUND
    plt.plot(df['DATE'], df['UPPER_BOUND'], color='#0F7173', linewidth=2, label='UPPER_BOUND')
    
    # Line plot for LOWER_BOUND
    plt.plot(df['DATE'], df['LOWER_BOUND'], color='#0F7173', linewidth=2, label='LOWER_BOUND')
    
    # Titles and labels
    plt.title('UNIT_RATE with LOWESS & BOUNDS')
    plt.xlabel('DATE')
    plt.ylabel('UNIT_RATE')
    plt.legend()
    plt.grid(True)
    
    # Rotate x-axis labels if needed
    plt.xticks(rotation=45)
    
    # Show the plot
    plt.tight_layout()
    plt.show()
    
    # Automatically determine the start and end dates
    start_date = df['DATE'].min()
    end_date = df['DATE'].max()
    
    # Create full date range (excluding Sundays)
    full_date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    full_date_range = full_date_range[full_date_range.weekday != 6]
    
    # Create a DataFrame with full date range
    df_full = pd.DataFrame(full_date_range, columns=['DATE'])
    
    # Merge with the original data (left join to preserve all dates)
    df = pd.merge(df_full, df, on='DATE', how='left')
    
    # Forward fill, then backward fill missing UNIT_RATE
    
    # Polynomial interpolation example
    df['UNIT_RATE'] = df['UNIT_RATE'].interpolate(method='polynomial', order=1)
    
    # Polynomial interpolation example
    df['MODIFIED_LOWESS'] = df['MODIFIED_LOWESS'].interpolate(method='polynomial', order=1)
    
    # Polynomial interpolation example
    df['UPPER_BOUND'] = df['UPPER_BOUND'].interpolate(method='polynomial', order=1)

    # Polynomial interpolation example
    df['LOWER_BOUND'] = df['LOWER_BOUND'].interpolate(method='polynomial', order=1)    
    
    # Set the figure size
    plt.figure(figsize=(14, 6))
    
    # Line plot for UNIT_RATE
    plt.plot(df['DATE'], df['UNIT_RATE'], color='#0639EF', linewidth=2, label='UNIT_RATE')
    
    # Line plot for MODIFIED_LOWESS
    plt.plot(df['DATE'], df['MODIFIED_LOWESS'], color='#DE9954', linewidth=2, label='MODIFIED_LOWESS')
    
    # Line plot for UPPER_BOUND
    plt.plot(df['DATE'], df['UPPER_BOUND'], color='#0F7173', linewidth=2, label='UPPER_BOUND')
    
    # Line plot for LOWER_BOUND
    plt.plot(df['DATE'], df['LOWER_BOUND'], color='#0F7173', linewidth=2, label='LOWER_BOUND')
       
    # Titles and labels
    plt.title('UNIT_RATE with LOWESS & BOUNDS')
    plt.xlabel('DATE')
    plt.ylabel('UNIT_RATE')
    plt.legend()
    plt.grid(True)
    
    # Rotate x-axis labels if needed
    plt.xticks(rotation=45)
    
    # Show the plot
    plt.tight_layout()
    plt.show()
    
    # Create a new folder named 'Stage-4_lowess' inside the source directory
    lowess_folder = os.path.join(source_folder, "Stage-4_lowess")
    os.makedirs(lowess_folder, exist_ok=True)  # Create the folder if it doesn't exist

    # Create output file name by adding 'value_selected_' prefix
    output_path = os.path.join(lowess_folder, "lowess_" + os.path.basename(file_path))
    df.to_excel(output_path, index=False)
    print(f"File: {os.path.basename(file_path)}")
    print(f"Saved filtered file to: {output_path}\n")        
# %%

"""
Stage-5: Final Usable Value for KYP
"""

# Get a list of all Excel files in value_cleaned_folder
value_cleaned_files = glob.glob(os.path.join(lowess_folder, "*.xlsx"))

# List to hold the output rows
output_data = []

# Loop through each file
for file_path in value_cleaned_files:
    # Load Excel file
    df = pd.read_excel(file_path)

    # Extract product name from filename
    file_name = os.path.basename(file_path)
    if "PRODUCT_" in file_name:
        product_name = file_name.split("PRODUCT_")[-1].replace(".xlsx", "").strip()
    else:
        product_name = "UNKNOWN"

    # Ensure DATE column is datetime
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')

    # Drop rows with missing dates
    df = df.dropna(subset=['DATE', 'UNIT_RATE'])

    if df.empty:
        latest_date = None
        latest_unit_rate = None
    else:
        # Get row with latest date
        latest_row = df.loc[df['DATE'].idxmax()]
        latest_date = latest_row['DATE'].date()
        latest_unit_rate = latest_row['UNIT_RATE']

    # Add to output list
    output_data.append({
        'Product Name': product_name,
        'Latest Date': latest_date,
        'Latest UNIT_RATE': latest_unit_rate
    })
    
    # Print output
    print(f"Product: {product_name}")
    print(f"Latest Date: {latest_date}")
    print(f"Latest UNIT_RATE: {latest_unit_rate}\n")

# Convert to DataFrame and save to Excel
output_df = pd.DataFrame(output_data)
output_path = os.path.join(source_folder, "Final Output.xlsx")
output_df.to_excel(output_path, index=False)