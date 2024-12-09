import json
import pandas as pd
import os

# Load the GeoJSON data
with open('data/fields/this_thing_works.geojson') as f:
    geojson_data = json.load(f)

# Initialize list to hold the output data
output_data = []

# Iterate through each feature in the GeoJSON
for feature in geojson_data['features']:
    field_id = feature['properties']['field_id']
    properties = feature['properties']
    
    # Add all properties to output data
    output_data.append({
        'field_id': field_id,
        'id': properties['id'],
        **{k: properties[k] for k in properties if k not in ['field_id']}
    })

# Create a DataFrame for the output
df_output = pd.DataFrame(output_data)

# Ensure the output directory exists
output_dir = 'data/fields/'
os.makedirs(output_dir, exist_ok=True)

# Save to Excel with single sheet
df_output.to_excel(os.path.join(output_dir, 'plet_output.xlsx'), 
                  sheet_name='Output', 
                  index=False)

print("plet_output.xlsx has been created successfully in data/fields/")
