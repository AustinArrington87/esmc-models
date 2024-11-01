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
def fetch_field_details(abbr="%", year=None):
    # Define the GraphQL query with variables for `abbr` and `year`
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

    # Define the variables, including year if provided
    variables = {
        "abbr": abbr,
        "year": year
    }

    # Send the request with the query and variables
    response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)

    # Process the response
    if response.status_code == 200:
        print("Query successful!")
        response_data = response.json()
        
        # Extract project abbreviation for file naming
        project_abbr = abbr
        file_name_base = f"{project_abbr}_{year}"

        # Save response as JSON
        json_path = Path(f"{file_name_base}.json")
        with open(json_path, "w") as json_file:
            json.dump(response_data, json_file, indent=2)
        print(f"Saved JSON file: {json_path}")

        # Create and save GeoJSON based on boundary information
        features = []
        for field in response_data["data"]["farmField"]:
            features.append({
                "type": "Feature",
                "properties": {
                    "id": field["id"],
                    "name": field["name"],
                    "email": field["app_user"]["email"] if field["app_user"] else None,
                    "displayName": field["app_user"]["displayName"] if field["app_user"] else None,
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
abbr_value = "CIFCSC"  # Replace with the desired abbreviation
year_value = 2024  # Replace with the desired year

# EXPORT FIELDS 
#field_details = fetch_field_details(abbr=abbr_value, year=year_value)
#print(field_details)

# EXPORT PROJECT ACRES
fetch_acres_summary(year=year_value)

# get project ID 

#mrv.projectFieldBoundaries('CIFCSC', 2024)

#producerFieldBoundaries(producerId,year)
