import json
import pandas as pd
import os

# Load the GeoJSON data
with open('data/fields/this_thing_works.geojson') as f:
    geojson_data = json.load(f)

# Initialize a dictionary to hold the aggregated data
aggregated_data = {}
output_data = []

# Iterate through each feature in the GeoJSON
for feature in geojson_data['features']:
    field_id = feature['properties']['field_id']
    properties = feature['properties']
    
    # Add to output data for the first tab
    output_data.append({
        'field_id': field_id,
        'id': properties['id'],
        **{k: properties[k] for k in properties if k not in ['field_id']}
    })
    
    # Create a unique key for aggregation based on field_id
    if field_id not in aggregated_data:
        aggregated_data[field_id] = {**properties}  # Copy properties
        aggregated_data[field_id]['count'] = 1  # Initialize count for summation
    else:
        # Sum the specified properties
        for prop in ['b_run_v', 'b_run_n', 'b_run_p', 'b_run_s', 'p_run_n', 'p_run_p', 'p_run_s']:
            if prop in properties:
                aggregated_data[field_id][prop] += properties[prop]
        
        # Average the specified properties
        for prop in ['pc_v', 'pc_n', 'pc_p', 'pc_s']:
            if prop in properties:
                aggregated_data[field_id][prop] += properties[prop]
        
        aggregated_data[field_id]['count'] += 1  # Increment count

# Create a DataFrame for the first tab (output)
df_output = pd.DataFrame(output_data)

# Prepare data for the second DataFrame (aggregated_results)
aggregated_results = []
for field_id, props in aggregated_data.items():
    # Create a new row for aggregated results
    row = {
        'field_id': field_id,
        'b_run_v': props['b_run_v'],
        'b_run_n': props['b_run_n'],
        'b_run_p': props['b_run_p'],
        'b_run_s': props['b_run_s'],
        'p_run_n': props['p_run_n'],
        'p_run_p': props['p_run_p'],
        'p_run_s': props['p_run_s'],
        'pc_v': round(props['pc_v'] / props['count'], 2),
        'pc_n': round(props['pc_n'] / props['count'], 2),
        'pc_p': round(props['pc_p'] / props['count'], 2),
        'pc_s': round(props['pc_s'] / props['count'], 2),
    }
    aggregated_results.append(row)

# Create a DataFrame for the aggregated results
df_aggregated = pd.DataFrame(aggregated_results)

# Ensure the output directory exists
output_dir = 'data/fields/'
os.makedirs(output_dir, exist_ok=True)

# Save to Excel with two sheets
with pd.ExcelWriter(os.path.join(output_dir, 'plet_output.xlsx')) as writer:
    df_output.to_excel(writer, sheet_name='Output', index=False)
    df_aggregated.to_excel(writer, sheet_name='aggregated_results', index=False)

print("plet_output.xlsx has been created successfully in data/fields/")
