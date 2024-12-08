import requests
import json
from pathlib import Path
import csv
import pandas as pd

# Global variables
url = "https://graphql.ecoharvest.ag/v1/graphql"
admin_secret_key = "EnterSecret"

# Set up the headers with the Hasura admin secret
headers = {
    "Content-Type": "application/json",
    "x-hasura-admin-secret": admin_secret_key
}

project_names = {
    "SNW": "Sustainable NW",
    "GC": "Great Lakes Project",
    "GM-NE": "General Mills-Native Energy",
    "BRANDY": "Brandywine-Christina Pilot",
    "LANCPA": "Lancaster Pilot",
    "SHENVA": "Shenandoah Valley Pilot",
    "MDLZ": "Mondelez Wheat",
    "TNCWF": "TNC/WWF (Northern Great Plains)",
    "TNCOH": "TNC Ohio",
    "PNWWA": "Pacific NW (WA)",
    "SN": "Syngenta-Nutrien",
    "NAL": "NASA Ag Lab",
    "SLM": "SLM Partners",
    "MOCS": "Missouri Partnership Pilot",
    "AGIP": "AGI SGP Pipeline",
    "NPC": "Northern Plains Cropping",
    "NJ_RCD": "New Jersey Resource and Conservation Department",
    "CTNUSCTP": "Cotton USCTP",
    "CIFCSC": "CIF Climate Smart",
    "LT": "Laura Test",
    "PGT": "PG Test Project",
    "KC": "Kimberly-Clark",
    "NGP-PFQF": "NGP-PFQF Research",
    "NPCRED": "Northern Plains Cropping Reductions Accounting",
    "SGPNACDP": "SGP-NACD Pilot",
    "NACDSGP": "NACD SGP Market",
    "SGP": "Southern Great Plains",
    "CN": "Corteva",
    "TNCNE": "TNC Nebraska",
    "FJFCSC": "FJF CSC Grazing",
    "SOR": "Sorghum",
    "CTNMNU": "Cotton Manulife",
    "CNFA": "CA Nut Fruit (Almond Board)",
    "MOBIO": "Missouri Biodiversity Pilot",
    "NGPNACDP": "NGP-NACD Pipeline",
    "KSBD": "Kansas PFQF Biodiversity",
    "DAIRY": "Trinkler Dairy",
    "NGPNACDM": "NGP-NACD Market",
    "TNCMN": "TNC Minnesota",
    "AGIM": "AGI SGP Market"
}

# Function to fetch field details based on abbreviation and year
import requests
import json
from pathlib import Path
import csv
import pandas as pd

# [Previous code remains the same until the fetch_field_details function]

def fetch_field_details(abbr="%", year=None):
    query = """
    query FieldDetails($abbr: String = "%", $year: smallint!) {
      farmField(where: {
        many_field_has_many_practice_changes: {practice_change: {abbreviation: {_nin: ["PE", "UE"]}}, year: {_eq: $year}}, 
        seasons: {year: {_eq: $year}},
        farmer_project_fields: {farmer_project: {project: {abbreviation: {_ilike: $abbr}}}}
      }) {
        id
        name
        acres
        boundary
        boundaryArray
        subboundaries
        many_field_has_many_practice_changes {
          practice_change {
            abbreviation
            name
            description
          }
          year
        }
        app_user {
          displayName
          id
          email
          phone
          street
          city
          state {
            stusps
          }
          zipcode
        }
        farmer_project_fields {
          farmer_project {
            project {
              name
              abbreviation
              id
            }
          }
        }
      }
    }
    """

    variables = {
        "abbr": abbr,
        "year": year
    }

    response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)

    if response.status_code == 200:
        print("Query successful!")
        response_data = response.json()
        
        project_abbr = abbr
        file_name_base = f"{project_abbr}_{year}"

        # Save response as JSON
        json_path = Path(f"{file_name_base}.json")
        with open(json_path, "w") as json_file:
            json.dump(response_data, json_file, indent=2)
        print(f"Saved JSON file: {json_path}")

        # Create and save GeoJSON with practice change information
        features = []
        for field in response_data["data"]["farmField"]:
            # Process practice changes
            practice_changes = []
            if field.get("many_field_has_many_practice_changes"):
                for pc in field["many_field_has_many_practice_changes"]:
                    if pc["practice_change"]:
                        practice_changes.append({
                            "abbreviation": pc["practice_change"]["abbreviation"],
                            "name": pc["practice_change"]["name"],
                            "description": pc["practice_change"]["description"]
                        })

            features.append({
                "type": "Feature",
                "properties": {
                    "id": field["id"],
                    "name": field["name"],
                    "email": field["app_user"]["email"] if field["app_user"] else None,
                    "displayName": field["app_user"]["displayName"] if field["app_user"] else None,
                    "practice_changes": practice_changes
                },
                "geometry": {
                    "type": "Polygon",
                    "coordinates": field["boundaryArray"] if field["boundaryArray"] else field["boundary"]
                }
            })

        geojson_data = {
            "type": "FeatureCollection",
            "features": features
        }

        geojson_path = Path(f"{file_name_base}.geojson")
        with open(geojson_path, "w") as geojson_file:
            json.dump(geojson_data, geojson_file, indent=2)
        print(f"Saved GeoJSON file: {geojson_path}")

        return response_data
    else:
        print(f"Query failed with status code {response.status_code}")
        print(response.text)
        return None

# Function to fetch and sum acres for each project abbreviation
def fetch_acres_summary(year):
    # Define the GraphQL query to fetch acres and project name
    query = """
    query FieldAcres($abbr: String, $year: smallint!) {
      farmField(where: {
        many_field_has_many_practice_changes: {practice_change: {abbreviation: {_nin: ["PE", "UE"]}}, year: {_eq: $year}},
        seasons: {year: {_eq: $year}},
        farmer_project_fields: {farmer_project: {project: {abbreviation: {_ilike: $abbr}}}}
      }) {
        id
        name
        acres
        many_field_has_many_practice_changes {
          practice_change {
            abbreviation
          }
        }
        farmer_project_fields {
          farmer_project {
            project {
              name
              abbreviation
              id
            }
          }
        }
      }
    }
    """

    # List to store summary data and field-level data
    summary_data = []
    field_level_data = []

    # Loop through each project abbreviation and name
    for abbr, full_name in project_names.items():
        # Define variables for each project
        variables = {
            "abbr": abbr,
            "year": year
        }

        # Send the request with the query and variables
        response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)

        # Process the response
        if response.status_code == 200:
            response_data = response.json()
            
            # Extract acres and count fields
            total_acres = 0
            field_count = 0
            
            for field in response_data["data"]["farmField"]:
                acres = field.get("acres")
                practice_changes = [pc["practice_change"]["abbreviation"] for pc in field["many_field_has_many_practice_changes"]]
                practice_change = ", ".join(practice_changes) if practice_changes else "None"
                
                if acres is not None:
                    total_acres += acres
                    field_count += 1

                # Append field-level data
                field_level_data.append({
                    "project_abbreviation": abbr,
                    "project_name": full_name,
                    "field_id": field["id"],
                    "field_name": field["name"],
                    "acres": acres,
                    "practice_change": practice_change,
                    "year": year
                })

            # Only add project-level data if field_count is greater than 0
            if field_count > 0:
                summary_data.append({
                    "project_abbreviation": abbr,
                    "project_name": full_name,
                    "field_count": field_count,
                    "total_acres": total_acres,
                    "year": year
                })
        else:
            print(f"Query failed for project {abbr} with status code {response.status_code}")
            print(response.text)

    # Sort summary data alphabetically by project_abbreviation
    summary_data = sorted(summary_data, key=lambda x: x["project_abbreviation"])

    # Create DataFrames for both summary and field-level data
    summary_df = pd.DataFrame(summary_data)
    field_level_df = pd.DataFrame(field_level_data)

    # Write data to an Excel file with two sheets
    excel_file_path = Path(f"project_acres_summary_{year}.xlsx")
    with pd.ExcelWriter(excel_file_path) as writer:
        summary_df.to_excel(writer, sheet_name="Project Summary", index=False)
        field_level_df.to_excel(writer, sheet_name="Field Level Details", index=False)

    print(f"Excel file saved: {excel_file_path}")

# Example usage
#abbr_value = "CIFCSC"  # Replace with the desired abbreviation
#year_value = 2024  # Replace with the desired year

# EXPORT FIELDS 
#field_details = fetch_field_details(abbr=abbr_value, year=year_value)
#print(field_details)

# EXPORT PROJECT ACRES
#fetch_acres_summary(year=year_value)

# get project ID 

#mrv.projectFieldBoundaries('CIFCSC', 2024)

#producerFieldBoundaries(producerId,year)

####################### MAP for PLET ##########


# {
# "type": "FeatureCollection",
# "name": "test_field_file_output5",
# "crs": { "type": "name", "properties": { "name": "urn:ogc:def:crs:OGC:1.3:CRS84" } },
# "features": [
# { "type": "Feature", "properties": { "id": "0", "field_id": "field_01", "user_lu": "cropland", "n_months": 2, "m_area_ac": 350.0, "bmp_name": "forest_buffer_100ft", "bmp_ac": 175.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.059852842732084, 40.502358400081064 ], [ -89.059494973079566, 40.495066805910746 ], [ -89.069157453697912, 40.495111539617355 ], [ -89.069157453697912, 40.496319349694659 ], [ -89.070096861535802, 40.496364083401168 ], [ -89.070096861535802, 40.495022072204222 ], [ -89.073630824354552, 40.49506680591076 ], [ -89.083114370146646, 40.498824437262336 ], [ -89.083069636440072, 40.502268932667945 ], [ -89.059852842732084, 40.502358400081064 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "1", "field_id": "field_02", "user_lu": "cropland", "n_months": 8, "m_area_ac": 30.0, "bmp_name": "cons_till_2", "bmp_ac": 69.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.059539706786126, 40.494798403671382 ], [ -89.0593160382533, 40.490056630775314 ], [ -89.068665382925687, 40.493143256528377 ], [ -89.068620649219113, 40.494843137377927 ], [ -89.059539706786126, 40.494798403671382 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "2", "field_id": "field_03", "user_lu": "cropland", "n_months": 6, "m_area_ac": 176.0, "bmp_name": "cov_crop_2", "bmp_ac": 176.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.059987043851791, 40.510410467263021 ], [ -89.059852842732084, 40.502492601200814 ], [ -89.069381122230737, 40.50244786749419 ], [ -89.069604790763577, 40.510320999849888 ], [ -89.059987043851791, 40.510410467263021 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "3", "field_id": "field_04", "user_lu": "pastureland", "n_months": 0, "m_area_ac": 0.0, "bmp_name": "forest_buffer_35ft", "bmp_ac": 245.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.06973899188327, 40.510320999849931 ], [ -89.069515323350444, 40.502492601200842 ], [ -89.082980169026939, 40.502358400081071 ], [ -89.083203837559779, 40.510142065023672 ], [ -89.06973899188327, 40.510320999849931 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "4", "field_id": "field_05", "user_lu": "cropland", "n_months": 3, "m_area_ac": 47.0, "bmp_name": "bioreactor", "bmp_ac": 15.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.083453308788961, 40.515187827469674 ], [ -89.08265225168536, 40.515210593324284 ], [ -89.082634163348487, 40.515853155228406 ], [ -89.083458221025452, 40.515898221093416 ], [ -89.083525677061118, 40.516878332235244 ], [ -89.081939313041147, 40.516913627692851 ], [ -89.081940214136708, 40.518578036779772 ], [ -89.083584405789239, 40.518556997087423 ], [ -89.08359376995989, 40.519678177651627 ], [ -89.080732100377261, 40.519743582226333 ], [ -89.080821639957492, 40.519127118089536 ], [ -89.08006656004018, 40.518220341804266 ], [ -89.079425788529434, 40.515378059918156 ], [ -89.079071291180611, 40.514459907836468 ], [ -89.083282990996821, 40.512526643054521 ], [ -89.083340776659512, 40.512540898924648 ], [ -89.083453308788961, 40.515187827469674 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "5", "field_id": "field_06", "user_lu": "cropland", "n_months": 4, "m_area_ac": 72.0, "bmp_name": "cov_crop_3", "bmp_ac": 72.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.079037405090602, 40.51448509942297 ], [ -89.079401855348436, 40.515403395309335 ], [ -89.080060119391391, 40.518286563948621 ], [ -89.080791134692689, 40.519136545107351 ], [ -89.080681511905709, 40.51980738940528 ], [ -89.080610668819375, 40.520535834232604 ], [ -89.078985441449703, 40.520513492461944 ], [ -89.078913377777511, 40.520798039615336 ], [ -89.07434639812972, 40.520903131473993 ], [ -89.074164095966708, 40.516088666744459 ], [ -89.074557435437669, 40.51599924681301 ], [ -89.07624117733809, 40.515857086214069 ], [ -89.07650154498117, 40.515793613097635 ], [ -89.079037405090602, 40.51448509942297 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "6", "field_id": "field_07", "user_lu": "pastureland", "n_months": 0, "m_area_ac": 0.0, "bmp_name": "graze_mgmt", "bmp_ac": 28.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.065618750972618, 40.516967930260783 ], [ -89.060186940244449, 40.517120062498677 ], [ -89.060125151546814, 40.5142650251437 ], [ -89.06297190529051, 40.514199279119708 ], [ -89.063116611310377, 40.514210784341024 ], [ -89.065618750972618, 40.516967930260783 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "7", "field_id": "field_08", "user_lu": "pastureland", "n_months": 6, "m_area_ac": 46.0, "bmp_name": "buffer_opt_graze", "bmp_ac": 23.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.300755997006377, 40.57504585067921 ], [ -89.300146514352122, 40.57525736794841 ], [ -89.300078226961205, 40.574903331923522 ], [ -89.297008013671302, 40.574908399147624 ], [ -89.297820193398763, 40.572330851280647 ], [ -89.297930218532116, 40.571163525237644 ], [ -89.297968128168051, 40.57011426780479 ], [ -89.299387856081765, 40.569717119838337 ], [ -89.300214076541693, 40.569578023583219 ], [ -89.300455371850504, 40.569606030052263 ], [ -89.301662319862643, 40.570172679040134 ], [ -89.30209565805886, 40.570104194405396 ], [ -89.302110866281751, 40.5721008725749 ], [ -89.301063543427318, 40.572213046904501 ], [ -89.301030394227013, 40.574938485323635 ], [ -89.300755997006377, 40.57504585067921 ] ] ] ] } },
# { "type": "Feature", "properties": { "id": "8", "field_id": "field_09", "user_lu": "cropland", "n_months": 5, "m_area_ac": 40.0, "bmp_name": "ditch", "bmp_ac": 30.0 }, "geometry": { "type": "MultiPolygon", "coordinates": [ [ [ [ -89.310319276719611, 40.569804154066972 ], [ -89.309120048888488, 40.56903005031463 ], [ -89.307865368243952, 40.568599696773354 ], [ -89.305320885322743, 40.568603829783299 ], [ -89.304303692489242, 40.568844720095093 ], [ -89.303460856141243, 40.569198835134742 ], [ -89.302880910554848, 40.569462845116341 ], [ -89.302137970665569, 40.569492876201203 ], [ -89.302072228974396, 40.565034362128713 ], [ -89.305929634997753, 40.564839721216806 ], [ -89.310068331489873, 40.564850512255667 ], [ -89.310245940504345, 40.564929237696482 ], [ -89.310319276719611, 40.569804154066972 ] ] ] ] } }
# ]
# }


def extract_coordinates(field_data):
    """
    Helper function to extract coordinates from boundary or boundaryArray
    
    Args:
        field_data (dict): Field data containing boundary or boundaryArray
    
    Returns:
        list: Coordinates in MultiPolygon format
    """
    # First try boundaryArray, then fall back to boundary
    boundary_data = field_data.get("boundaryArray") or field_data.get("boundary")
    
    if isinstance(boundary_data, dict) and "coordinates" in boundary_data:
        return boundary_data["coordinates"]
    return boundary_data

def transform_to_plet_format(input_data, year):
    """
    Transform field data to PLET format with practice change prioritization.
    
    Args:
        input_data (dict): The response data from fetch_field_details
        year (int): Year to process
    """
    # Practice change mapping
    bmp_mapping = {
        "CC": "cov_crop_2",
        "TR": "cons_till_1",
        "NM": "nutrient_mgmt_1"
    }
    
    # Practice change priority (highest to lowest)
    priority_order = ["CC", "TR", "NM"]
    
    features = []
    
    for field in input_data["data"]["farmField"]:
        # Get practice changes for the specified year
        practice_changes = []
        for pc in field.get("many_field_has_many_practice_changes", []):
            if pc["year"] == year and pc["practice_change"]["abbreviation"] in priority_order:
                practice_changes.append(pc["practice_change"]["abbreviation"])
        
        # Skip if no relevant practice changes
        if not practice_changes:
            continue
            
        # Select highest priority practice change
        selected_practice = None
        for practice in priority_order:
            if practice in practice_changes:
                selected_practice = practice
                break
        
        # Skip if no matching practice change found
        if not selected_practice:
            continue
        
        # Extract coordinates
        coordinates = extract_coordinates(field)
        if not coordinates:
            print(f"Warning: No valid coordinates found for field {field['id']}, skipping...")
            continue
            
        # Create feature
        feature = {
            "type": "Feature",
            "properties": {
                "id": str(field["id"]),
                "field_id": field["name"],
                "user_lu": "cropland",
                "n_months": 12,
                "m_area_ac": float(field["acres"]) if field["acres"] is not None else 0.0,
                "bmp_name": bmp_mapping[selected_practice],
                "bmp_ac": float(field["acres"]) if field["acres"] is not None else 0.0
            },
            "geometry": {
                "type": "MultiPolygon",
                "coordinates": coordinates
            }
        }
        
        features.append(feature)
    
    # Create final GeoJSON structure
    plet_geojson = {
        "type": "FeatureCollection",
        "name": f"field_output_{year}",
        "crs": {
            "type": "name",
            "properties": {
                "name": "urn:ogc:def:crs:OGC:1.3:CRS84"
            }
        },
        "features": features
    }
    
    # Save to file
    output_path = Path(f"plet_{year}.geojson")
    with open(output_path, "w") as f:
        json.dump(plet_geojson, f, indent=2)
    
    print(f"Saved PLET GeoJSON file: {output_path}")
    return plet_geojson

def get_plet_data(abbr, year):
    """
    Fetch field details and transform to PLET format.
    
    Args:
        abbr (str): Project abbreviation
        year (int): Year to process
    """
    # Get field details
    field_data = fetch_field_details(abbr=abbr, year=year)
    
    # Transform to PLET format
    if field_data:
        plet_data = transform_to_plet_format(field_data, year)
        return plet_data
    return None

# Usage example
abbr_value = "CIFCSC"  # Replace with the desired abbreviation
year_value = 2024  # Replace with the desired year

# Call the function to get PLET formatted data
plet_data = get_plet_data(abbr_value, year_value)

if plet_data:
    print(f"Successfully transformed data for project {abbr_value} in year {year_value}")
    print(f"Number of fields processed: {len(plet_data['features'])}")
else:
    print("Failed to process data")
