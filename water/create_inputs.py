import csv
import json
import datetime
import os
from dateutil import parser  # Import parser for flexible date parsing

# Input CSV file path
csv_file_path = '2023_Verified.csv'

# Output directory
output_dir = 'water_inputs'
os.makedirs(output_dir, exist_ok=True)  # Ensure the directory exists

# Extract the year from the CSV file name
year = int(csv_file_path.split('/')[-1].split('_')[0])

# Define function to determine BMP name based on conditions
def get_bmp_name(practice_change, planting_date, commodity):
    if practice_change == "Cover cropping":
        if not planting_date:
            return "cov_crop_1"
        else:
            try:
                # Attempt to parse the date in any format
                planting_date = parser.parse(planting_date).replace(tzinfo=None)  # Remove timezone info
            except ValueError:
                # Skip if date format is incorrect
                print(f"Skipping due to date format error: {planting_date}")
                return None
            
            # Compare as offset-naive datetime
            if datetime.datetime(year, 5, 1) <= planting_date <= datetime.datetime(year, 12, 31):
                return "cov_crop_2"
            elif datetime.datetime(year - 1, 1, 1) <= planting_date <= datetime.datetime(year, 4, 30):
                return "cov_crop_3"
            else:
                return "cov_crop_1"
    elif practice_change == "Tillage reduction":
        return "cons_till_1"
    elif practice_change == "Nutrient management":
        return "nutrient_mgmt_2"
    else:
        return None


# Read the CSV file, normalize headers, and process each row
project_geojsons = {}

with open(csv_file_path, 'r') as csvfile:
    csvreader = csv.reader(csvfile)
    headers = next(csvreader)
    headers = [header.strip() for header in headers]  # Normalize headers by stripping whitespace
    normalized_headers = {header.lower(): index for index, header in enumerate(headers)}  # Map lowercase header to index

    # Locate column names dynamically
    project_col = next((key for key in normalized_headers if "project" in key), None)
    practice_change_col = next((key for key in normalized_headers if "practice change" in key), None)
    planting_date_col = next((key for key in normalized_headers if "cover crop planting" in key), None)
    commodity_col = next((key for key in normalized_headers if "commodity" in key), None)
    field_id_col = next((key for key in normalized_headers if "field id" in key), None)
    field_name_col = next((key for key in normalized_headers if "field name (project)" in key), None)
    acres_col = next((key for key in normalized_headers if "acres" in key), None)
    longitude_col = next((key for key in normalized_headers if "longitude" in key), None)
    latitude_col = next((key for key in normalized_headers if "latitude" in key), None)

    # Check if required columns exist
    required_columns = [project_col, practice_change_col, field_id_col, field_name_col, acres_col, longitude_col, latitude_col]
    if not all(required_columns):
        print("Error: One or more required columns not found in the CSV.")
        exit()

    # Process each row
    for row in csvreader:
        project_value = row[normalized_headers[project_col]]
        
        # Initialize a GeoJSON for each unique project
        if project_value not in project_geojsons:
            project_geojsons[project_value] = {
                "type": "FeatureCollection",
                "features": []
            }
        
        # Split practices and process each individually
        practice_changes = row[normalized_headers[practice_change_col]].split(", ")
        planting_date = row[normalized_headers[planting_date_col]] if planting_date_col else ""
        commodity = row[normalized_headers[commodity_col]] if commodity_col else ""
        field_id = row[normalized_headers[field_id_col]]
        field_name = row[normalized_headers[field_name_col]]
        acres = float(row[normalized_headers[acres_col]])
        longitude = float(row[normalized_headers[longitude_col]])
        latitude = float(row[normalized_headers[latitude_col]])

        for index, practice in enumerate(practice_changes, start=1):
            bmp_name = get_bmp_name(practice, planting_date, commodity)
            if not bmp_name:
                continue  # Skip if BMP name is not defined

            # Prepare feature for GeoJSON
            feature = {
                "type": "Feature",
                "properties": {
                    "id": f"{field_id}_{index}",
                    "field_id": field_name,
                    "user_lu": "cropland",
                    "bmp_name": bmp_name,
                    "bmp_acre": acres,
                    "n_months": 0,
                    "m_area_ac": 0,
                },
                "geometry": {
                    "type": "Point",
                    "coordinates": [
                        longitude,
                        latitude
                    ]
                }
            }
            project_geojsons[project_value]["features"].append(feature)

# Write each project GeoJSON to a separate file in the output directory
for project, geojson_data in project_geojsons.items():
    output_geojson_path = f'{output_dir}/{project.replace(" ", "_")}_{year}.geojson'
    with open(output_geojson_path, 'w') as geojsonfile:
        json.dump(geojson_data, geojsonfile, indent=2)
    print(f"GeoJSON file for project '{project}' has been created at {output_geojson_path}")
