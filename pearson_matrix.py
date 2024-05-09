import pandas as pd
import numpy as np
from scipy.stats import pearsonr

def encode_dates(date_series):
    return pd.to_datetime(date_series).map(pd.Timestamp.toordinal)

def encode_categorical(data):
    return data.astype('category').cat.codes

def compute_correlations(data, main_col, other_cols):
    correlations = {}
    for col in other_cols:
        if data[col].dtype == 'object':
            try:
                # Try to convert it to datetime if it looks like a date
                temp_data = encode_dates(data[col])
            except:
                # Otherwise treat as categorical
                temp_data = encode_categorical(data[col])
        else:
            temp_data = data[col]
        
        # Drop NaN values and ensure equal length
        aligned_data = pd.DataFrame({
            'main': data[main_col],
            'secondary': temp_data
        }).dropna()

        # Calculate Pearson correlation if there are at least two unique values
        if len(aligned_data['secondary'].unique()) > 1:
            corr, _ = pearsonr(aligned_data['main'], aligned_data['secondary'])
            correlations[col] = corr
        else:
            correlations[col] = 'Undefined (constant or insufficient unique data)'
    return correlations

def main():
    # Load the data
    df = pd.read_csv('2022_Verified.csv')
    
    # Prepare the Excel writer
    writer = pd.ExcelWriter('Project_Correlations.xlsx', engine='xlsxwriter')
    
    # Columns to correlate with 'removed'
    columns_to_correlate = ['Commodity', 'Date (Cover Crop Planting)', 'soil_avg_soc', 
                            'soil_avg_bulkdensity', 'soil_clay_fraction', 
                            'optis_tillage_fall__value', 'optis_tillage_spring__value']
    
    # Iterate over each project
    for project_name, group in df.groupby('Project'):
        correlations = compute_correlations(group, 'removed', columns_to_correlate)
        pd.DataFrame.from_dict(correlations, orient='index', columns=['Correlation']).to_excel(writer, sheet_name=project_name)
    
    # Save the Excel file
    writer.save()

if __name__ == "__main__":
    main()
