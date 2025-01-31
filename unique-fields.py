import json

def read_geojson(file_path):
    """Read and parse a GeoJSON file"""
    with open(file_path, 'r') as f:
        return json.load(f)

def get_field_ids(geojson_data):
    """Extract all fieldIds from a GeoJSON FeatureCollection"""
    return {
        feature['properties']['fieldId'] 
        for feature in geojson_data['features'] 
        if 'fieldId' in feature['properties']
    }

def create_filtered_geojson(geojson_2025, field_ids_2024):
    """Create new GeoJSON with features only present in 2025"""
    new_features = [
        feature for feature in geojson_2025['features']
        if 'fieldId' in feature['properties'] and 
        feature['properties']['fieldId'] not in field_ids_2024
    ]
    
    return {
        "type": "FeatureCollection",
        "features": new_features
    }

def main():
    # Read both GeoJSON files
    geojson_2024 = read_geojson('CIF-2024.geojson')
    geojson_2025 = read_geojson('CIF-2025.geojson')
    
    # Get sets of fieldIds from both years
    field_ids_2024 = get_field_ids(geojson_2024)
    field_ids_2025 = get_field_ids(geojson_2025)
    
    # Create new GeoJSON with unique 2025 features
    new_geojson = create_filtered_geojson(geojson_2025, field_ids_2024)
    
    # Write the new GeoJSON to a file
    with open('CIF-2025-unique.geojson', 'w') as f:
        json.dump(new_geojson, f, indent=2)
    
    # Print some statistics
    print(f"Fields in 2024: {len(field_ids_2024)}")
    print(f"Fields in 2025: {len(field_ids_2025)}")
    print(f"Unique fields in 2025: {len(new_geojson['features'])}")
    print(f"New GeoJSON file created: CIF-2025-unique.geojson")

if __name__ == "__main__":
    main()
