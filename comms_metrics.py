import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows

# Load data
# verified_data = pd.read_csv('2023_Verified.csv')
# quantified_data = pd.read_csv('2023_Quantified.csv')

# # Combine the data
# combined_data = pd.concat([verified_data, quantified_data])

# List of files to combine
files = [
    '2023_Verified.csv',
    '2023_Quantified.csv',
    '2022_Verified.csv',
    '2022_Quantified.csv',
    '2021_Verified.csv'
]

# Combine the data from all files
combined_data = pd.concat([pd.read_csv(f) for f in files])

# Filter and aggregate the data
filtered_data = combined_data[
    combined_data['Commodity'].str.contains('Corn|Cotton|Wheat|Soybeans', case=False, na=False)
]

# Replace multiple crops including wheat into "Wheat"
filtered_data.loc[filtered_data['Commodity'].str.contains('Wheat', case=False, na=False), 'Commodity'] = 'Wheat'

# Sum acres by state and commodity
state_commodity_acres = filtered_data.groupby(['field_state', 'Commodity'])['Acres'].sum().unstack(fill_value=0)

# Filter for practice change containing "grazing"
grazing_data = combined_data[combined_data['Practice Change'].str.contains('grazing', case=False, na=False)]
state_grazing_acres = grazing_data.groupby('field_state')['Acres'].sum()

# Ensure there are no NaN values in the aggregated data
state_commodity_acres = state_commodity_acres.fillna(0)
state_grazing_acres = state_grazing_acres.fillna(0)

# Save the data to an Excel file
output_file = 'aggregated_data.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    state_commodity_acres.to_excel(writer, sheet_name='State_Commodity_Acres')
    state_grazing_acres.to_excel(writer, sheet_name='State_Grazing_Acres')

def plot_pie_chart(data, title, filename):
    data = data.sort_values(ascending=False)
    threshold = 0.01 * data.sum()
    small_data = data[data < threshold].sum()
    large_data = data[data >= threshold]
    if small_data > 0:
        large_data["Other"] = small_data
    fig, ax = plt.subplots(figsize=(8, 8))
    large_data.plot.pie(ax=ax, autopct='%1.1f%%', startangle=90)
    ax.set_ylabel('')
    ax.set_title(title)
    plt.savefig(filename)
    plt.close()

# Create and save pie charts as individual images
commodities = ['Corn', 'Wheat', 'Soybeans', 'Cotton']
for commodity in commodities:
    if commodity in state_commodity_acres.columns:
        plot_pie_chart(state_commodity_acres[commodity], f'Acres by State for {commodity}', f'pie_chart_{commodity}.png')

if not state_grazing_acres.empty:
    plot_pie_chart(state_grazing_acres, 'Acres by State (Grazing)', 'pie_chart_grazing.png')

# Create a US map with the data
us_map = gpd.read_file(gpd.datasets.get_path('naturalearth_lowres'))
us_map = us_map[us_map['continent'] == 'North America']  # Filter for North America first

# List of Canadian municipalities to exclude
canadian_municipalities = ["Manitoba", "Saskatchewan"]

# Filter out Canadian municipalities from the state_commodity_acres
state_acres = state_commodity_acres.sum(axis=1).reset_index()
state_acres.columns = ['field_state', 'Acres']
state_acres = state_acres[~state_acres['field_state'].isin(canadian_municipalities)]

# Remove rows with NaN values in 'field_state'
state_acres = state_acres.dropna(subset=['field_state'])

# Merge with US states GeoDataFrame
state_acres = us_map.merge(state_acres, left_on='name', right_on='field_state', how='left')

# Ensure there are no NaN values in the Acres column
state_acres['Acres'] = state_acres['Acres'].fillna(0)

# Create the map
fig, ax = plt.subplots(1, 1, figsize=(15, 10))
us_map.boundary.plot(ax=ax)
state_acres.plot(column='Acres', ax=ax, legend=True,
                 legend_kwds={'label': "Total Acres by State"},
                 cmap='OrRd')

plt.title('Total Acres by State for Selected Commodities')
plt.savefig('us_map.png')
plt.close()

# Save the data and charts to an Excel file
wb = Workbook()
ws1 = wb.active
ws1.title = 'State_Commodity_Acres'
for r in dataframe_to_rows(state_commodity_acres, index=True, header=True):
    ws1.append(r)

ws2 = wb.create_sheet(title='State_Grazing_Acres')
for r in dataframe_to_rows(state_grazing_acres.reset_index(), index=False, header=True):
    ws2.append(r)

# Add pie charts to Excel file
img_paths = [f'pie_chart_{commodity}.png' for commodity in commodities if commodity in state_commodity_acres.columns]
img_paths.append('pie_chart_grazing.png')

# Insert pie chart images into the Excel file
for idx, img_path in enumerate(img_paths):
    img = Image(img_path)
    ws1.add_image(img, f'D{idx * 20 + 2}')

# Add US map to Excel file
img = Image('us_map.png')
ws1.add_image(img, 'D100')

# Save the Excel file
wb.save(output_file)

print(f"Data and charts saved to {output_file}")
