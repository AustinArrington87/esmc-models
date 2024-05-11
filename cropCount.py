import pandas as pd

def summarize_acreage(csv_file, output_csv_file):
    # Load the data from the CSV file
    data = pd.read_csv(csv_file)
    
    # Fill missing or empty values in 'crop_name' with 'Null'
    data['crop_name'] = data['crop_name'].fillna('Null')
    
    # Group the data by 'crop_name' and sum the 'acres' for each crop
    acreage_summary = data.groupby('crop_name')['acres'].sum().reset_index()
    
    # Rename the columns for clarity
    acreage_summary.columns = ['Crop Name', 'Total Acres']
    
    # Output the summary table to a CSV file
    acreage_summary.to_csv(output_csv_file, index=False)
    print(f"Summary has been saved to {output_csv_file}")

# Replace 'yourfile.csv' with the path to your CSV file and specify an output file name
summarize_acreage('baseline_check.csv', '2024_planned_crops.csv')
