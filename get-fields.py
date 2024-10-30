import requests
import json
from pathlib import Path

# Global variables
url = "https://graphql.ecoharvest.ag/v1/graphql"
admin_secret_key = "EnterSecret"

# Set up the headers with the Hasura admin secret
headers = {
    "Content-Type": "application/json",
    "x-hasura-admin-secret": admin_secret_key
}

# Function to fetch field details based on abbreviation and year
def fetch_field_details(abbr="%", year=None):
    # Define the GraphQL query with variables for `abbr` and `year`
    query = """
    query FieldDetails($abbr: String = "%", $year: smallint!) {
      farmField(where: {
        many_field_has_many_practice_changes: {practice_change: {abbreviation: {_neq: "PE"}}, year: {_eq: $year}}, 
        seasons: {year: {_eq: $year}},
        farmer_project_fields: {farmer_project: {project: {abbreviation: {_ilike: $abbr}}}}
      }) {
        id
        name
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

# Example usage
abbr_value = "KC"  # Replace with the desired abbreviation
year_value = 2024  # Replace with the desired year

field_details = fetch_field_details(abbr=abbr_value, year=year_value)
print(field_details)
# get project ID 

#mrv.projectFieldBoundaries('CIFCSC', 2024)

#producerFieldBoundaries(producerId,year)


# Project abbreviations 

# {
#   "esmcProject": [
#     {
#       "name": "Simplot",
#       "abbreviation": "SMP"
#     },
#     {
#       "name": "Austin-Test-2",
#       "abbreviation": "AT2"
#     },
#     {
#       "name": "Chris Demo",
#       "abbreviation": "CD"
#     },
#     {
#       "name": "CIF Test Project",
#       "abbreviation": "CIF_TEST"
#     },
#     {
#       "name": "Jake-Test-1",
#       "abbreviation": "JMDTEST"
#     },
#     {
#       "name": "SustainCERT Project",
#       "abbreviation": "SCT"
#     },
#     {
#       "name": "Texan by Nature",
#       "abbreviation": "TBN"
#     },
#     {
#       "name": "Sustainable NW",
#       "abbreviation": "SNW"
#     },
#     {
#       "name": "GLRI CTIC",
#       "abbreviation": "GC"
#     },
#     {
#       "name": "General Mills-Native Energy",
#       "abbreviation": "GM-NE"
#     },
#     {
#       "name": "Brandywine-Christina Pilot",
#       "abbreviation": "BRANDY"
#     },
#     {
#       "name": "Lancaster Pilot",
#       "abbreviation": "LANCPA"
#     },
#     {
#       "name": "DemoProject4",
#       "abbreviation": "RX57"
#     },
#     {
#       "name": "Regrow test project",
#       "abbreviation": "REGROW"
#     },
#     {
#       "name": "Pacific NW (OR)",
#       "abbreviation": "PNWOR"
#     },
#     {
#       "name": "SLM Demo",
#       "abbreviation": "SLM-D"
#     },
#     {
#       "name": "The Fertilizer Institute (TFI)",
#       "abbreviation": "TFI"
#     },
#     {
#       "name": "Corteva--Archived 2",
#       "abbreviation": "CNA2"
#     },
#     {
#       "name": "Corteva--Archived",
#       "abbreviation": "CNA"
#     },
#     {
#       "name": "NCBA",
#       "abbreviation": "NCBA"
#     },
#     {
#       "name": "Farmobile",
#       "abbreviation": "FM"
#     },
#     {
#       "name": "Illinois Corn Growers",
#       "abbreviation": "ICG"
#     },
#     {
#       "name": "Shenandoah Valley Pilot",
#       "abbreviation": "SHENVA"
#     },
#     {
#       "name": "Mondelez Wheat",
#       "abbreviation": "MDLZ"
#     },
#     {
#       "name": "Producer Circle",
#       "abbreviation": "PC"
#     },
#     {
#       "name": "Regrow",
#       "abbreviation": "RGW"
#     },
#     {
#       "name": "Michelle-test",
#       "abbreviation": "MLS-TEST"
#     },
#     {
#       "name": "TNC/WWF (Northern Great Plains)",
#       "abbreviation": "TNCWF"
#     },
#     {
#       "name": "TNC Ohio",
#       "abbreviation": "TNCOH"
#     },
#     {
#       "name": "Pacific NW (WA)",
#       "abbreviation": "PNWWA"
#     },
#     {
#       "name": "Silicon Ranch",
#       "abbreviation": "SR"
#     },
#     {
#       "name": "Syngenta-Nutrien",
#       "abbreviation": "SN"
#     },
#     {
#       "name": "NASA Ag Lab",
#       "abbreviation": "NAL"
#     },
#     {
#       "name": "SLM Partners",
#       "abbreviation": "SLM"
#     },
#     {
#       "name": "JLO Test Project GHG",
#       "abbreviation": "JLO"
#     },
#     {
#       "name": "Missouri Partnership Pilot",
#       "abbreviation": "MOCS"
#     },
#     {
#       "name": "AGI SGP Pipeline",
#       "abbreviation": "AGIP"
#     },
#     {
#       "name": "Northern Plains Cropping",
#       "abbreviation": "NPC"
#     },
#     {
#       "name": "z_test_Lisa",
#       "abbreviation": "Z_TEST"
#     },
#     {
#       "name": "Margarita-Test-Project",
#       "abbreviation": "MTP"
#     },
#     {
#       "name": "New Jersey Resource and Conservation Department",
#       "abbreviation": "NJ_RCD"
#     },
#     {
#       "name": "Cabbage Patch- Will Test",
#       "abbreviation": "PATCH"
#     },
#     {
#       "name": "John Deere",
#       "abbreviation": "JD"
#     },
#     {
#       "name": "Corteva-Test",
#       "abbreviation": "CTT"
#     },
#     {
#       "name": "ESMC-PM-Test",
#       "abbreviation": "ESMCPM"
#     },
#     {
#       "name": "ESMC-Test",
#       "abbreviation": "ESMC"
#     },
#     {
#       "name": "Test Project 10",
#       "abbreviation": "TEST10"
#     },
#     {
#       "name": "Cotton USCTP",
#       "abbreviation": "CTNUSCTP"
#     },
#     {
#       "name": "CIF Climate Smart",
#       "abbreviation": "CIFCSC"
#     },
#     {
#       "name": "Travis-Test1",
#       "abbreviation": "TRAVTEST"
#     },
#     {
#       "name": "Laura Test",
#       "abbreviation": "LT"
#     },
#     {
#       "name": "PG-Test-Project",
#       "abbreviation": "PGT"
#     },
#     {
#       "name": "Kimberly-Clark",
#       "abbreviation": "KC"
#     },
#     {
#       "name": "Coop Elev Demo",
#       "abbreviation": "COOPDEMO"
#     },
#     {
#       "name": "NGP-PFQF Research",
#       "abbreviation": "NGP-PFQF"
#     },
#     {
#       "name": "Northern Plains Cropping Reductions Accounting",
#       "abbreviation": "NPCRED"
#     },
#     {
#       "name": "Z Test Lisa",
#       "abbreviation": "ZTESTLIS"
#     },
#     {
#       "name": "G's EoY test project",
#       "abbreviation": "EOYTEST"
#     },
#     {
#       "name": "All Scopes and Assets/Credits",
#       "abbreviation": "All"
#     },
#     {
#       "name": "SGP-NACD Pilot",
#       "abbreviation": "SGPNACDP"
#     },
#     {
#       "name": "Data Modeling",
#       "abbreviation": "DATAAPI"
#     },
#     {
#       "name": "Test Default Project",
#       "abbreviation": "TDP"
#     },
#     {
#       "name": "Austin-Test-1",
#       "abbreviation": "AT1"
#     },
#     {
#       "name": "Austin-Test-3",
#       "abbreviation": "AT3"
#     },
#     {
#       "name": "Austin-Test-5",
#       "abbreviation": "AT5"
#     },
#     {
#       "name": "Historical Fields",
#       "abbreviation": "HSF"
#     },
#     {
#       "name": "Jake-Test 2",
#       "abbreviation": "JDTEST2"
#     },
#     {
#       "name": "NACD SGP Market",
#       "abbreviation": "NACDSGP"
#     },
#     {
#       "name": "Southern Great Plains",
#       "abbreviation": "SGP"
#     },
#     {
#       "name": "Corteva",
#       "abbreviation": "CN"
#     },
#     {
#       "name": "TNC Nebraska",
#       "abbreviation": "TNCNE"
#     },
#     {
#       "name": "FJF CSC Grazing",
#       "abbreviation": "FJFCSC"
#     },
#     {
#       "name": "Sorghum",
#       "abbreviation": "SOR"
#     },
#     {
#       "name": "Cotton Manulife",
#       "abbreviation": "CTNMNU"
#     },
#     {
#       "name": "CA Nut Fruit (Almond Board)",
#       "abbreviation": "CNFA"
#     },
#     {
#       "name": "Missouri Biodiversity Pilot",
#       "abbreviation": "MOBIO"
#     },
#     {
#       "name": "NGP-NACD Pipeline",
#       "abbreviation": "NGPNACDP"
#     },
#     {
#       "name": "Benson Hill",
#       "abbreviation": "BH"
#     },
#     {
#       "name": "AGI SGP Research",
#       "abbreviation": "AGIR"
#     },
#     {
#       "name": "Kansas PFQF Biodiversity",
#       "abbreviation": "KSBD"
#     },
#     {
#       "name": "Trinkler Dairy",
#       "abbreviation": "DAIRY"
#     },
#     {
#       "name": "NGP-NACD Market",
#       "abbreviation": "NGPNACDM"
#     },
#     {
#       "name": "TNC Minnesota",
#       "abbreviation": "TNCMN"
#     },
#     {
#       "name": "Corteva-Nutrien",
#       "abbreviation": "C-NUT"
#     },
#     {
#       "name": "AGI SGP Market",
#       "abbreviation": "AGIM"
#     }
#   ]
# }
