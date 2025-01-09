import pandas as pd

# Load the Excel file
data_file = "SCT S3S Flat File (Unfiltered) - OLD VERSION.xlsx"
sheet_name = "Data Results"  # Adjust the sheet name if necessary
df = pd.read_excel(data_file, sheet_name=sheet_name)

# Define projects to exclude
excluded_projects = [
    "Kimberly-Clark",
    "Northern Plains Cropping",
    "Northern Plains Cropping Reductions Accounting",
    "Shenandoah Valley Pilot"
]

# Filter out the specified projects
filtered_data = df[~df['Project'].isin(excluded_projects)]

# Include only rows where both 'field_state' and 'Commodity' are not missing
filtered_data_cleaned = filtered_data.dropna(subset=['field_state', 'Commodity'])

# Summing Acres by State
state_acres_sum = filtered_data_cleaned.groupby('field_state', as_index=False)['Acres'].sum()
state_acres_sum.loc[len(state_acres_sum)] = ['Total', state_acres_sum['Acres'].sum()]

# Summing Acres by Commodity
commodity_acres_sum = filtered_data_cleaned.groupby('Commodity', as_index=False)['Acres'].sum()
commodity_acres_sum.loc[len(commodity_acres_sum)] = ['Total', commodity_acres_sum['Acres'].sum()]

# Save results to new Excel file
output_file = "summed_acres_output.xlsx"
with pd.ExcelWriter(output_file) as writer:
    state_acres_sum.to_excel(writer, sheet_name="State Summary", index=False)
    commodity_acres_sum.to_excel(writer, sheet_name="Commodity Summary", index=False)

print(f"Summary tables saved to {output_file}")
