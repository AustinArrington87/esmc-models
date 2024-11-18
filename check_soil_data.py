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

# Output the results
print(differences_summary)

# Optionally, save the output to a CSV file
differences_summary.to_csv('Field_Differences_Summary.csv', index=False)
