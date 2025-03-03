# Checks for differneces in soil data after new SustainCert data file expeort produced 
# Written by Austin Arrington (PLANT Group LLC) November 18 2024

import pandas as pd

# Load the datasets
file_new = 'S3S_FlatFile_Verified_RanBackend_20241117.xlsx'  # Replace with actual file path
file_adjusted = 'Data Results.csv'  # Replace with actual file path

data_new = pd.read_excel(file_new)
data_adjusted = pd.read_csv(file_adjusted)

# Extract and rename relevant columns for comparison
soil_data_new = data_new[['field_id', 'avg_ph', 'avg_soc', 'avg_bulk']].rename(
    columns={'avg_ph': 'soil_avg_ph', 'avg_soc': 'soil_avg_soc', 'avg_bulk': 'soil_avg_bulkdensity'}
)
soil_data_adjusted = data_adjusted[['Field ID', 'soil_avg_ph', 'soil_avg_soc', 'soil_avg_bulkdensity']].rename(
    columns={'Field ID': 'field_id'}
)

# Merge datasets on field_id to compare soil data side by side
merged_soil_data = soil_data_new.merge(soil_data_adjusted, on='field_id', suffixes=('_new', '_adjusted'))

# Create a new column to indicate differences
merged_soil_data['different'] = merged_soil_data.apply(
    lambda row: ', '.join([
        'SOC' if row['soil_avg_soc_new'] != row['soil_avg_soc_adjusted'] else '',
        'pH' if row['soil_avg_ph_new'] != row['soil_avg_ph_adjusted'] else '',
        'BD' if row['soil_avg_bulkdensity_new'] != row['soil_avg_bulkdensity_adjusted'] else ''
    ]).strip(', '),
    axis=1
)

# Filter only rows with differences
differences_summary = merged_soil_data[merged_soil_data['different'] != ''][['field_id', 'different']]

# Check for outliers in the new file
def check_outliers(row):
    outliers = []
    if row['soil_avg_ph'] < 4 or row['soil_avg_ph'] > 9:
        outliers.append('pH')
    if row['soil_avg_soc'] < 0.003 or row['soil_avg_soc'] > 0.02:
        outliers.append('SOC')
    if row['soil_avg_bulkdensity'] < 1.1 or row['soil_avg_bulkdensity'] > 1.9:
        outliers.append('BD')
    return ', '.join(outliers)

# Apply outlier check to the new dataset
soil_data_new['outliers'] = soil_data_new.apply(check_outliers, axis=1)

# Filter rows with outliers
outliers_summary = soil_data_new[soil_data_new['outliers'] != ''][['field_id', 'outliers']]

# Output the results
print("Differences Summary:")
print(differences_summary)
print("\nOutliers Summary:")
print(outliers_summary)

# Optionally, save the outputs to CSV files
differences_summary.to_csv('Field_Differences_Summary.csv', index=False)
outliers_summary.to_csv('Field_Outliers_Summary.csv', index=False)
