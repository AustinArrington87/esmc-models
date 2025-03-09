import time
import requests
import json
import csv
from datetime import datetime

# Token management
token = None
token_expiry = 0

def get_token():
    global token, token_expiry
    url = "https://api.meti.millpont.com/token"
    payload = {
        "email": "email",
        "password": "password"
    }
    
    # Create a session for consistent headers
    session = requests.Session()
    session.headers.update({
        "Content-Type": "application/json",
        "Accept": "application/json"
    })
    
    response = session.post(url, json=payload)
    print(f"Status code: {response.status_code}")
    
    if response.status_code == 200:
        try:
            # Store raw token without any modifications
            token = response.text.strip()
            token_expiry = time.time() + 3600
            
            # Set the raw token as the Authorization header
            session.headers["Authorization"] = token
            
            # Test the token
            test_response = session.get("https://api.meti.millpont.com/sources")
            print(f"\nToken test:")
            print(f"Status: {test_response.status_code}")
            print(f"Response: {test_response.text}")
            
            if test_response.status_code == 200:
                print("Token working correctly!")
                return session
            else:
                raise ValueError("Token test failed")
            
        except Exception as e:
            print(f"Error processing token: {str(e)}")
            raise
    else:
        raise ValueError(f"Authentication failed: {response.text}")

def standardize_geojson(field):
    """
    Standardize a GeoJSON feature to match the required format:
    - Remove CRS
    - Fix coordinate structure
    - Ensure proper nesting
    - Ensure it's a Polygon type
    """
    try:
        # Get the geometry
        geometry = field["geometry"]
        
        # Extract coordinates, handling both direct and nested formats
        if isinstance(geometry["coordinates"], dict):
            # Handle nested format with type and crs
            coordinates = geometry["coordinates"].get("coordinates", [])
        else:
            coordinates = geometry["coordinates"]
        
        # Ensure we have the right nesting level for a Polygon
        # We want: [[[x1,y1], [x2,y2], ...]] for a polygon
        if isinstance(coordinates[0][0][0], list):  # Too many levels (MultiPolygon)
            coordinates = coordinates[0]  # Take first polygon
        
        # Ensure it's a closed polygon
        if coordinates[0][0] != coordinates[0][-1]:
            coordinates[0].append(coordinates[0][0])
        
        # Create standardized geometry
        standardized_geometry = {
            "type": "Polygon",
            "coordinates": coordinates
        }
        
        # Create standardized feature
        standardized_feature = {
            "type": "Feature",
            "id": field["properties"]["id"],
            "properties": {
                "start_at": "2024-01-01T00:00:00.000Z",
                "end_at": "2024-12-31T23:59:59.999Z"
            },
            "geometry": standardized_geometry
        }
        
        return standardized_feature
        
    except Exception as e:
        print(f"Error standardizing feature: {e}")
        print(f"Original feature: {json.dumps(field, indent=2)}")
        raise

def create_millpont_feature(field):
    """Convert a field from CIFCSC to Millpont format"""
    # First standardize the GeoJSON
    standardized = standardize_geojson(field)
    return standardized

def process_fields():
    try:
        session = get_token()
        
        # Read the GeoJSON file
        with open('CIFCSC_2024.geojson', 'r') as f:
            data = json.load(f)
        
        # Process each field
        results = []
        for field in data["features"]:
            try:
                feature_collection = {
                    "type": "FeatureCollection",
                    "features": [create_millpont_feature(field)]
                }
                
                url = "https://api.meti.millpont.com/sources"
                
                print(f"\nProcessing field: {field['properties'].get('name', 'unknown')}")
                
                # Make request using session with raw token
                response = session.post(url, json=feature_collection)
                print(f"Response status: {response.status_code}")
                
                if response.status_code == 200:
                    response_data = response.json()
                    if isinstance(response_data, list) and len(response_data) > 0:
                        source = response_data[0]  # Get the first source object
                        results.append({
                            'field_id': field["properties"]["id"],
                            'field_name': field["properties"]["name"],
                            'millpont_source_id': source.get('id'),
                            'created_at': source.get('created_at'),
                            'processed': source.get('processed', False),
                            'created_by': source.get('created_by'),
                            'status': 'success'
                        })
                        print(f"Successfully created source with ID: {source.get('id')}")
                    else:
                        raise Exception("Unexpected response format")
                else:
                    raise Exception(f"API Error: {response.text}")
                
            except Exception as e:
                print(f"Error processing field {field['properties'].get('name', 'unknown')}: {str(e)}")
                results.append({
                    'field_id': field["properties"].get("id", "unknown"),
                    'field_name': field["properties"].get("name", "unknown"),
                    'millpont_source_id': None,
                    'created_at': None,
                    'processed': None,
                    'created_by': None,
                    'status': 'error',
                    'error_message': str(e)
                })
        
        # Save results to CSV with expanded columns
        if results:
            fieldnames = [
                'field_id', 
                'field_name', 
                'millpont_source_id',
                'created_at',
                'processed',
                'created_by',
                'status',
                'error_message'
            ]
            
            with open('millpont_results.csv', 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(results)
                print("\nResults saved to millpont_results.csv")
                
            # Print summary
            success_count = sum(1 for r in results if r['status'] == 'success')
            error_count = sum(1 for r in results if r['status'] == 'error')
            print(f"\nProcessing Summary:")
            print(f"Total fields: {len(results)}")
            print(f"Successful: {success_count}")
            print(f"Failed: {error_count}")
        
    except Exception as e:
        print(f"Error in process_fields: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        process_fields()
    except Exception as e:
        print(f"Script failed: {str(e)}")
