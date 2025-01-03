import mrvApi as mrv
import requests
import json
import uuid
import pandas as pd
import datetime as datetime

c = mrv.configure(".env.production")

# Your Hasura admin secret key
admin_secret_key = "EnterSecret"

# Set up the headers with the Hasura admin secret
headers = {
    "Content-Type": "application/json",
    "x-hasura-admin-secret": admin_secret_key
}

# Define the GraphQL endpoint
url = "https://graphql.ecoharvest.ag/v1/graphql"

# Mapping of crop types to commodity IDs - used for mapping from AGI CPFR 
crop_type_to_id = {
    "Soybeans": 2,
    "Winter Wheat": 3,
    "Sorghum": 10,
    "Corn": 1,
    "Wheat": 25,
    "Oats": 8,
    "Sunflower": 22,
    "Cow Peas": 24
}

#application method Dictionary
application_id = {
    "Anhydrous Ammonia Applicator": 1,
    "Banded & Incorporated": 2,
    "Broadcast": 3,
    "Fertigation, Drip": 4,
    "Fertigation, Furrow": 10,
    "Fertigation, Sprinkler": 11,
    "Fertigation, Subsurface Drip": 12,
    "Foliar": 13,
    "Injected": 5,
    "Sidedress": 7,
    "Surface Banded": 8,
    "Topdress": 9,
    "Other": 6
}

# tillage - eventID 3
tillage_type = {
  "tillageType": [
    {
      "id": 1,
      "name": "Chisel",
      "defaultDepthInInch": 7
    },
    {
      "id": 16,
      "name": "Crimp",
      "defaultDepthInInch": 0
    },
    {
      "id": 2,
      "name": "Cultivator",
      "defaultDepthInInch": 2.5
    },
    {
      "id": 3,
      "name": "Deep Ripping",
      "defaultDepthInInch": 14
    },
    {
      "id": 4,
      "name": "Disc",
      "defaultDepthInInch": 6.5
    },
    {
      "id": 17,
      "name": "Flail",
      "defaultDepthInInch": 0
    },
    {
      "id": 27,
      "name": "Harrow",
      "defaultDepthInInch": 1.5
    },
    {
      "id": 28,
      "name": "Harrow (heavy)",
      "defaultDepthInInch": 2
    },
    {
      "id": 29,
      "name": "Harrow (light)",
      "defaultDepthInInch": 1
    },
    {
      "id": 18,
      "name": "Leveler",
      "defaultDepthInInch": 2
    },
    {
      "id": 5,
      "name": "Mulch Tiller",
      "defaultDepthInInch": 0
    },
    {
      "id": 6,
      "name": "Offset (heavy) Disc",
      "defaultDepthInInch": 6.5
    },
    {
      "id": 15,
      "name": "Other",
      "defaultDepthInInch": 6.5
    },
    {
      "id": 7,
      "name": "Plow",
      "defaultDepthInInch": 10
    },
    {
      "id": 19,
      "name": "Rake",
      "defaultDepthInInch": 2
    },
    {
      "id": 8,
      "name": "Ridge Till",
      "defaultDepthInInch": 7.5
    },
    {
      "id": 24,
      "name": "Rod Weed",
      "defaultDepthInInch": 2.5
    },
    {
      "id": 20,
      "name": "Roller",
      "defaultDepthInInch": 0
    },
    {
      "id": 9,
      "name": "Rotary Hoe",
      "defaultDepthInInch": 6
    },
    {
      "id": 10,
      "name": "Row Cultivator",
      "defaultDepthInInch": 2.5
    },
    {
      "id": 11,
      "name": "Strip Till",
      "defaultDepthInInch": 7.5
    },
    {
      "id": 12,
      "name": "Strip Till Freshener",
      "defaultDepthInInch": 7.5
    },
    {
      "id": 25,
      "name": "Subsoil",
      "defaultDepthInInch": 17.5
    },
    {
      "id": 26,
      "name": "Sweep",
      "defaultDepthInInch": 7
    },
    {
      "id": 13,
      "name": "Tandem (light) Disc",
      "defaultDepthInInch": 6.5
    },
    {
      "id": 14,
      "name": "Vertical Tillage",
      "defaultDepthInInch": 2.5
    }
  ]
}

# tillage residue ID
tillage_residue = [
  {
    "id": 1,
    "name": "0-15%"
  },
  {
    "id": 2,
    "name": "15-30%"
  },
  {
    "id": 3,
    "name": "30-50%"
  },
  {
    "id": 4,
    "name": ">50%"
  }
]


# list projects 
#projects = mrv.projects()
#print(projects)
# get project ID 
#AGIprojId = mrv.projectId('AGI SGP Market')
#print(AGIprojId)

#AGIProducers = mrv.enrolledProducers(AGIprojId, 2024)
# "partner_grower_id" in AGI
# "partner_field_id" in AGI
#print(AGIProducers)

#AustinProducers = mrv.enrolledProducers('18a4db70-209d-4a93-87a9-ab3ff315fd14', 2024)
#print(AustinProducers)

#BarringtonFields = mrv.fieldSummary('1a8b2fd2-5a46-4d70-8a54-952871f43fca', 2024)
#print(BarringtonFields) 

########################

# Pass in the fieldID and get the seasonID back 

# Function to retrieve season_id based on fieldId and year
def get_season_id(fieldId, year):
    # Define the query with a variable for `fieldId` and `year`
    query = """
    query MyQuery($fieldId: uuid!, $year: smallint!) {
      farmSeason(where: {fieldId: {_eq: $fieldId}, year: {_eq: $year}}) {
        id
      }
    }
    """
    
    # Define the variables
    variables = {
        "fieldId": fieldId,
        "year": year
    }

    # Send the request with the query and variables
    response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)

    # Process the response
    if response.status_code == 200:
        print("Query successful!")
        response_data = response.json()
        season_id = response_data['data']['farmSeason'][0]['id'] if response_data['data']['farmSeason'] else None
        print(f"Season ID for year {year}: {season_id}")
        return season_id
    else:
        print(f"Query failed with status code {response.status_code}")
        print(response.text)
        return None

#################################
# Add Planting Event 
def insert_plant_event(commodityId, eventId, doneAt, seasonId):
    # Define the mutation
    plant_mutation = """
    mutation insertEventData($commodityId: Int, $eventId: Int, $doneAt: date, $seasonId: uuid) {
      insertFarmEventData(objects: {commodityId: $commodityId, eventId: $eventId, doneAt: $doneAt, seasonId: $seasonId}) {
        affected_rows
      }
    }
    """

    # Define the variables
    plant_variables = {
        "commodityId": commodityId,
        "eventId": eventId,
        "doneAt": doneAt,
        "seasonId": seasonId
    }

    # Send the request
    response = requests.post(url, json={'query': plant_mutation, 'variables': plant_variables}, headers=headers)
    
    if response.status_code == 200:
        response_data = response.json()
        if 'errors' not in response_data:
            affected_rows = response_data['data']['insertFarmEventData']['affected_rows']
            print(f"Plant event created successfully. Rows affected: {affected_rows}")
            return affected_rows
        else:
            print(f"Failed to create plant event with errors: {response_data['errors']}")
            return 0
    else:
        print(f"HTTP request failed with status code {response.status_code}")
        print(response.text)
        return 0

# Used for Parsing from AGI CPFR - Harvest events 
def parse_planting_events(producer_id, field_id, specific_year=None):
    with open('cpfrs_by_grower.json') as f:
        data = json.load(f)

    if producer_id not in data:
        print(f"Producer ID {producer_id} not found in JSON data.")
        return

    # Create a set to track all processed events globally
    all_processed_events = set()
    
    fields = data[producer_id]
    found = False

    for field in fields:
        if field.get("partner_field_id") == field_id:
            found = True
            print(f"\nProcessing field_id: {field_id} for producer: {producer_id}")
            field_events = field.get("field_events", [])
            
            if not field_events:
                print(f"No events found for field_id: {field_id}")
                return

            print(f"Total field events found: {len(field_events)}")

            for event in field_events:
                crop_cycle_summaries = event.get("crop_cycle_summaries", {})
                planting_list = crop_cycle_summaries.get("cash_crop_planting", [])

                print(f"\nProcessing planting list with {len(planting_list)} events")

                for planting_event in planting_list:
                    start_date = planting_event.get("start_date", "")
                    crop_type = planting_event.get("crop_type", "")
                    
                    # Create a unique event identifier
                    event_id = f"{field_id}_{start_date}_{crop_type}"

                    print(f"\nConsidering event: {event_id}")

                    if event_id in all_processed_events:
                        print(f"Skipping duplicate event: {event_id}")
                        continue

                    # Add to processed set before processing
                    all_processed_events.add(event_id)

                    # Parse the date components
                    date_parts = start_date.split("-")
                    year = int(date_parts[0])
                    month = int(date_parts[1])
                    
                    # If specific_year is provided, determine if this event should be included
                    if specific_year:
                        # Only include events from the target year OR
                        # events from Aug-Dec of the previous year
                        if year == specific_year:
                            print(f"Including event from target year {year}")
                        elif year == specific_year - 1 and month >= 8:
                            print(f"Including fall planting from {year}-{month:02d} for {specific_year} season")
                        else:
                            print(f"Skipping event from {year}-{month:02d} (outside target range for {specific_year} season)")
                            continue

                        # Use the target year for fall plantings
                        season_year = specific_year if (year == specific_year - 1 and month >= 8) else year
                    else:
                        season_year = year

                    season_id = get_season_id(field_id, season_year)
                    if not season_id:
                        print(f"No season_id found for field {field_id} in year {season_year}")
                        continue

                    commodity_id = crop_type_to_id.get(crop_type)

                    if commodity_id is not None:
                        print(f"Processing planting event: {crop_type} on {start_date} for {season_year} season")
                        affected_rows = insert_plant_event(
                            commodityId=commodity_id,
                            eventId=1,  # Event ID for planting is 1
                            doneAt=start_date.split("T")[0],
                            seasonId=season_id
                        )
                        print(f"Affected Rows: {affected_rows}")
                    else:
                        print(f"Skipping: Commodity ID not found for crop type: {crop_type}")

    if not found:
        print(f"Field ID {field_id} not found for producer {producer_id}")

    print(f"\nTotal unique events processed: {len(all_processed_events)}")

########################
# Add Harvest Event Using Mutation 

def insert_harvest_event(commodityId, eventId, doneAt, yield_value, seasonId):
    # Define the mutation
    harvest_mutation = """
    mutation insertEventData($commodityId: Int, $eventId: Int, $doneAt: date, $yield: numeric, $seasonId: uuid) {
      insertFarmEventData(objects: {commodityId: $commodityId, eventId: $eventId, doneAt: $doneAt, yield: $yield, seasonId: $seasonId}) {
        affected_rows
      }
    }
    """

    # Define the variables
    harvest_variables = {
        "commodityId": commodityId,
        "eventId": eventId,
        "doneAt": doneAt,
        "yield": yield_value,
        "seasonId": seasonId
    }

    # Send the request with the mutation and variables
    response = requests.post(url, json={'query': harvest_mutation, 'variables': harvest_variables}, headers=headers)

    # Process the response
    if response.status_code == 200:
        print("Mutation successful!")
        response_data = response.json()
        affected_rows = response_data['data']['insertFarmEventData']['affected_rows']
        #print(f"Affected Rows: {affected_rows}")
        return affected_rows
    else:
        print(f"Mutation failed with status code {response.status_code}")
        print(response.text)
        return None

################################################

# Parse Harvest - used for Parsing from AGI CPFR 
def parse_harvest_events(producer_id, field_id, specific_year=None):
    with open('cpfrs_by_grower.json') as f:
        data = json.load(f)

    if producer_id not in data:
        print(f"Producer ID {producer_id} not found in JSON data.")
        return

    # Create a set to track all processed events globally
    all_processed_events = set()
    
    fields = data[producer_id]
    found = False

    for field in fields:
        if field.get("partner_field_id") == field_id:
            found = True
            print(f"\nProcessing field_id: {field_id} for producer: {producer_id}")
            field_events = field.get("field_events", [])
            
            if not field_events:
                print(f"No events found for field_id: {field_id}")
                return

            # Debug: Print total number of field events
            print(f"Total field events found: {len(field_events)}")

            for event in field_events:
                crop_cycle_summaries = event.get("crop_cycle_summaries", {})
                harvest_list = crop_cycle_summaries.get("harvest", [])

                # Debug: Print harvest list length
                print(f"\nProcessing harvest list with {len(harvest_list)} events")

                for harvest_event in harvest_list:
                    end_date = harvest_event.get("end_date", "")
                    crop_type = harvest_event.get("crop_type", "")
                    
                    # Create a more unique event identifier
                    event_id = f"{field_id}_{end_date}_{crop_type}"

                    # Debug: Print current event being processed
                    print(f"\nConsidering event: {event_id}")

                    if event_id in all_processed_events:
                        print(f"Skipping duplicate event: {event_id}")
                        continue

                    # Add to processed set before processing
                    all_processed_events.add(event_id)

                    year = end_date.split("-")[0]
                    
                    # If specific_year is provided, skip events from other years
                    if specific_year and int(year) != specific_year:
                        print(f"Skipping event from year {year} (looking for {specific_year})")
                        continue

                    season_id = get_season_id(field_id, int(year))
                    if not season_id:
                        print(f"No season_id found for field {field_id} in year {year}")
                        continue

                    avg_dry_yield = harvest_event.get("avg_dry_yield", {}).get("value", 0)
                    commodity_id = crop_type_to_id.get(crop_type)

                    if commodity_id is not None:
                        print(f"Processing harvest event: {crop_type} on {end_date}")
                        affected_rows = insert_harvest_event(
                            commodityId=commodity_id,
                            eventId=5,
                            doneAt=end_date.split("T")[0],
                            yield_value=avg_dry_yield,
                            seasonId=season_id
                        )
                        print(f"Affected Rows: {affected_rows}")
                    else:
                        print(f"Skipping: Commodity ID not found for crop type: {crop_type}")

    if not found:
        print(f"Field ID {field_id} not found for producer {producer_id}")

    print(f"\nTotal unique events processed: {len(all_processed_events)}")


### FERTILIZER APPLICATIONS 

######################################## 
# Function to insert events across multiple seasons with fertilizer details
# Step 1: Insert core event details
def insert_event_multiple_seasons(eventId, seasonIds, doneAt):
    event_mutation = """
    mutation insertEventData($eventId: Int, $seasonId: uuid, $doneAt: date) {
      insertFarmEventData(objects: {
        eventId: $eventId,
        doneAt: $doneAt,
        seasonId: $seasonId
      }) {
        returning {
          id  # Retrieve the ID for further operations
        }
      }
    }
    """
    
    event_ids = []
    
    for seasonId in seasonIds:
        variables = {
            "eventId": eventId,
            "seasonId": seasonId,
            "doneAt": doneAt
        }
        
        response = requests.post(url, json={'query': event_mutation, 'variables': variables}, headers=headers)
        
        if response.status_code == 200:
            response_data = response.json()
            if 'errors' in response_data:
                print(f"Mutation failed for seasonId {seasonId} with errors: {response_data['errors']}")
            else:
                event_id = response_data['data']['insertFarmEventData']['returning'][0]['id']
                event_ids.append((seasonId, event_id))
                print(f"Event created for seasonId {seasonId} with event ID: {event_id}")
        else:
            print(f"HTTP request failed for seasonId {seasonId} with status code {response.status_code}")
            print(response.text)
    
    return event_ids

# Function to insert fertilizer data and then link it to an event
def insert_fertilizer_event(eventId, seasonIds, doneAt, applicationMethodId, fertilizerId, fertilizerCategoryId, rate, liquidDensity):
    total_affected_rows = 0

    # Loop over each seasonId and execute the mutations
    for seasonId in seasonIds:
        # Generate a unique fertilizer data ID
        fertilizer_data_id = str(uuid.uuid4())

        # Step 1: Insert the fertilizer data record
        fertilizer_data_mutation = """
        mutation insertFertilizerData(
            $fertilizerDataId: uuid,
            $applicationMethodId: smallint,
            $fertilizerId: smallint,
            $fertilizerCategoryId: String,
            $rate: numeric,
            $liquidDensity: numeric
        ) {
            insertFarmFertilizerData(objects: {
                id: $fertilizerDataId,
                applicationMethodId: $applicationMethodId,
                fertilizerId: $fertilizerId,
                fertilizerCategoryId: $fertilizerCategoryId,
                rate: $rate,
                liquidDensity: $liquidDensity
            }) {
                affected_rows
            }
        }
        """

        # Define the variables for inserting fertilizer data
        fertilizer_data_variables = {
            "fertilizerDataId": fertilizer_data_id,
            "applicationMethodId": applicationMethodId,
            "fertilizerId": fertilizerId,
            "fertilizerCategoryId": fertilizerCategoryId,
            "rate": rate,
            "liquidDensity": liquidDensity
        }

        # Send the request to insert fertilizer data
        response = requests.post(url, json={'query': fertilizer_data_mutation, 'variables': fertilizer_data_variables}, headers=headers)
        if response.status_code == 200 and 'errors' not in response.json():
            print(f"Fertilizer data created with ID: {fertilizer_data_id}")
        else:
            print(f"Failed to create fertilizer data for seasonId {seasonId}")
            continue  # Skip to the next seasonId if fertilizer data creation fails

        # Step 2: Insert the event data record, linking it to the fertilizer data
        event_data_mutation = """
        mutation insertEventData(
            $eventId: Int,
            $seasonId: uuid,
            $doneAt: date,
            $fertilizerDataId: uuid
        ) {
            insertFarmEventData(objects: {
                eventId: $eventId,
                doneAt: $doneAt,
                seasonId: $seasonId,
                fertilizerDataId: $fertilizerDataId
            }) {
                affected_rows
            }
        }
        """

        # Define the variables for inserting event data
        event_data_variables = {
            "eventId": eventId,
            "seasonId": seasonId,
            "doneAt": doneAt,
            "fertilizerDataId": fertilizer_data_id
        }

        # Send the request to insert event data
        response = requests.post(url, json={'query': event_data_mutation, 'variables': event_data_variables}, headers=headers)
        if response.status_code == 200:
            response_data = response.json()
            if 'errors' not in response_data:
                affected_rows = response_data['data']['insertFarmEventData']['affected_rows']
                total_affected_rows += affected_rows
                print(f"Event created and linked to fertilizer data for seasonId {seasonId}. Rows affected: {affected_rows}")
            else:
                print(f"Failed to link event data to fertilizer data for seasonId {seasonId} with errors: {response_data['errors']}")
        else:
            print(f"HTTP request failed for linking event to fertilizer data for seasonId {seasonId} with status code {response.status_code}")
            print(response.text)

    print(f"Total rows affected in linked event and fertilizer details: {total_affected_rows}")
    return total_affected_rows

# insert Tillage events 
def insert_tillage_event(eventId, seasonId, doneAt, tillage_name, tillageDepth=None, rowWidth=None, stripWidth=None, residue_name=None, tillage_type_dict=None, tillage_residue_dict=None):
    tillage_data_id = str(uuid.uuid4())
    
    # Get tillage type ID and name
    tillageTypeId = get_tillage_id(tillage_name, tillage_type_dict)
    if tillageTypeId is None:
        print(f"Error: Tillage type '{tillage_name}' not found")
        return 0
    
    if tillageDepth is None:
        for tillage in tillage_type_dict["tillageType"]:
            if tillage["id"] == tillageTypeId:
                tillageDepth = tillage["defaultDepthInInch"]
                break

    tillageResidueId = get_residue_id(residue_name, tillage_residue_dict) if residue_name else None
    
    # Step 1: Insert tillage data
    tillage_data_mutation = """
    mutation insertTillageData(
        $id: uuid,
        $striptillWidth: smallint,
        $striptillCultivated: smallint,
        $tillageResidueId: Int
    ) {
        insertFarmTillageData(objects: {
            id: $id,
            striptillWidth: $striptillWidth,
            striptillCultivated: $striptillCultivated,
            tillageResidueId: $tillageResidueId
        }) {
            affected_rows
        }
    }
    """
    
    tillage_variables = {
        "id": tillage_data_id
    }
    
    if rowWidth is not None:
        tillage_variables["striptillCultivated"] = rowWidth
    if stripWidth is not None:
        tillage_variables["striptillWidth"] = stripWidth
    if tillageResidueId is not None:
        tillage_variables["tillageResidueId"] = tillageResidueId

    response = requests.post(url, json={'query': tillage_data_mutation, 'variables': tillage_variables}, headers=headers)
    
    if response.status_code == 200:
        response_data = response.json()
        if 'errors' not in response_data:
            print(f"Tillage data created with ID: {tillage_data_id}")
        else:
            print(f"Failed to create tillage data with errors: {response_data['errors']}")
            return 0
    else:
        print(f"Failed to create tillage data with status code {response.status_code}")
        print(response.text)
        return 0

    # Step 2: Insert event data with tillageDepth and tillage type relationship
    event_data_mutation = """
    mutation insertEventData(
        $eventId: Int,
        $seasonId: uuid,
        $doneAt: date,
        $tillageDataId: uuid,
        $tillageDepth: numeric,
        $tillageTypeId: Int
    ) {
        insertFarmEventData(objects: {
            eventId: $eventId,
            seasonId: $seasonId,
            doneAt: $doneAt,
            tillageDataId: $tillageDataId,
            tillageDepth: $tillageDepth,
            many_event_data_has_many_tillage_types: {
                data: {
                    tillageTypeId: $tillageTypeId
                }
            }
        }) {
            affected_rows
        }
    }
    """
    
    event_variables = {
        "eventId": eventId,
        "seasonId": seasonId,
        "doneAt": doneAt,
        "tillageDataId": tillage_data_id,
        "tillageDepth": tillageDepth,
        "tillageTypeId": tillageTypeId
    }
    
    response = requests.post(url, json={'query': event_data_mutation, 'variables': event_variables}, headers=headers)
    
    if response.status_code == 200:
        response_data = response.json()
        if 'errors' not in response_data:
            affected_rows = response_data['data']['insertFarmEventData']['affected_rows']
            print(f"Event created and linked to tillage data. Rows affected: {affected_rows}")
            return affected_rows
        else:
            print(f"Failed to create event with errors: {response_data['errors']}")
            return 0
    else:
        print(f"HTTP request failed with status code {response.status_code}")
        print(response.text)
        return 0

# Load tillage IDs
def get_tillage_id(tillage_name, tillage_type_dict):
    """Get tillage type ID from name."""
    for tillage in tillage_type_dict["tillageType"]:
        if tillage["name"] == tillage_name:
            return tillage["id"]
    return None

# Load tillage Residue IDs
def get_residue_id(residue_name, tillage_residue_dict):
    """Get tillage residue ID from name."""
    for residue in tillage_residue_dict:
        if residue["name"] == residue_name:
            return residue["id"]
    return None

# load fert methods 
def load_fertilizer_mappings(json_path):
    """Load fertilizer mappings from JSON file."""
    with open(json_path, 'r') as f:
        fertilizers = json.load(f)
    
    # Create a mapping of fertilizer name to its details
    fert_map = {}
    for fert in fertilizers:
        fert_map[fert['name']] = {
            'id': int(fert['id']),
            'categoryId': str(fert['fertilizerCategoryId'])
        }
    return fert_map

# load commodity IDs
def load_commodity_mappings(json_path):
    """Load commodity mappings from JSON file."""
    with open(json_path, 'r') as f:
        commodities = json.load(f)
    
    # Create a mapping of commodity name to ID
    commodity_map = {}
    for commodity in commodities:
        commodity_map[commodity['name']] = int(commodity['id'])
    return commodity_map

# process Planting Events from Spreadsheet 
def process_planting_data(excel_path, commodities_json_path):
    """Process planting data from Excel sheet."""
    # Read Excel file - specifically from the Planting sheet
    df = pd.read_excel(excel_path, sheet_name='Planting')
    
    # Load commodity mappings
    commodity_map = load_commodity_mappings(commodities_json_path)
    
    # Get unique field-year combinations and get season IDs
    unique_field_years = df[['field_uuid', 'year']].drop_duplicates()
    season_ids_by_year_field = {}
    
    # Get season IDs and organize them by year and field
    for _, row in unique_field_years.iterrows():
        field_uuid = str(row['field_uuid'])
        year = int(row['year'])
        season_id = get_season_id(field_uuid, year)
        
        if year not in season_ids_by_year_field:
            season_ids_by_year_field[year] = {}
        season_ids_by_year_field[year][field_uuid] = season_id
    
    total_affected_rows = 0
    # Process each planting event
    for _, row in df.iterrows():
        # Get the season ID for this specific year and field
        year = int(row['year'])
        field_uuid = str(row['field_uuid'])
        season_id = season_ids_by_year_field[year][field_uuid]
        
        # Format planting date to YYYY-MM-DD
        if isinstance(row['planting_date'], str):
            done_at = datetime.strptime(row['planting_date'].split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        else:
            done_at = row['planting_date'].strftime('%Y-%m-%d')
        
        # Get commodity ID
        commodity_id = commodity_map[row['commodity']]
        
        # Insert planting event
        affected_rows = insert_plant_event(
            commodityId=commodity_id,
            eventId=1,  # Event ID for planting is 1
            doneAt=done_at,
            seasonId=season_id
        )
        total_affected_rows += affected_rows
    
    return total_affected_rows

# process Harvest Events from Excel Spreadsheet 
def process_harvest_data(excel_path, commodities_json_path):
    """Process harvest data from Excel sheet."""
    # Read Excel file - specifically from the Harvest sheet
    df = pd.read_excel(excel_path, sheet_name='Harvest')
    
    # Load commodity mappings
    commodity_map = load_commodity_mappings(commodities_json_path)
    
    # Get unique field-year combinations and get season IDs
    unique_field_years = df[['field_uuid', 'year']].drop_duplicates()
    season_ids_by_year_field = {}
    
    # Get season IDs and organize them by year and field
    for _, row in unique_field_years.iterrows():
        field_uuid = str(row['field_uuid'])
        year = int(row['year'])
        season_id = get_season_id(field_uuid, year)
        
        if year not in season_ids_by_year_field:
            season_ids_by_year_field[year] = {}
        season_ids_by_year_field[year][field_uuid] = season_id
    
    total_affected_rows = 0
    # Process each harvest event
    for _, row in df.iterrows():
        # Get the season ID for this specific year and field
        year = int(row['year'])
        field_uuid = str(row['field_uuid'])
        season_id = season_ids_by_year_field[year][field_uuid]
        
        # Format harvest date to YYYY-MM-DD
        if isinstance(row['harvest_date'], str):
            done_at = datetime.strptime(row['harvest_date'].split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        else:
            done_at = row['harvest_date'].strftime('%Y-%m-%d')
        
        # Get commodity ID
        commodity_id = commodity_map[row['commodity']]
        
        # Convert yield to float
        yield_value = float(row['yield'])
        
        # Insert harvest event
        affected_rows = insert_harvest_event(
            commodityId=commodity_id,
            eventId=5,  # Event ID for harvest is 5
            doneAt=done_at,
            yield_value=yield_value,
            seasonId=season_id
        )
        total_affected_rows += affected_rows if affected_rows is not None else 0
    
    return total_affected_rows

# process Fertilizer data from Excel Spreadhseet 
def process_fertilizer_data(excel_path, fertilizers_json_path):
    """Process fertilizer data from Excel and create events."""
    # Read Excel file - specifically from the Fertilizer sheet
    df = pd.read_excel(excel_path, sheet_name='Fertilizer')
    
    # Load fertilizer mappings
    fert_map = load_fertilizer_mappings(fertilizers_json_path)
    
    # Get unique field-year combinations and get season IDs
    unique_field_years = df[['field_uuid', 'year']].drop_duplicates()
    season_ids_by_year = {}  # Dictionary to store season IDs by year
    
    # Get season IDs and organize them by year and field
    season_ids_by_year_field = {}  # Dictionary to store season IDs by year and field
    
    for _, row in unique_field_years.iterrows():
        field_uuid = str(row['field_uuid'])
        year = int(row['year'])
        season_id = get_season_id(field_uuid, year)
        
        # Create nested dictionary structure
        if year not in season_ids_by_year_field:
            season_ids_by_year_field[year] = {}
        season_ids_by_year_field[year][field_uuid] = season_id
    
    # Process each fertilizer application
    for _, row in df.iterrows():
        # Get fertilizer details
        fert_details = fert_map[row['fert_type']]
        
        # Get the season ID for this specific year and field
        year = int(row['year'])
        field_uuid = str(row['field_uuid'])
        application_season = season_ids_by_year_field[year][field_uuid]
        # Put it in a list since the function expects a list of season IDs
        application_seasons = [application_season]
        
        # Format application date to YYYY-MM-DD
        if isinstance(row['application_date'], str):
            done_at = datetime.strptime(row['application_date'].split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        else:
            done_at = row['application_date'].strftime('%Y-%m-%d')
        
        # Get application method ID as integer
        application_method_id = int(application_id[row['application_method']])
        
        # Convert rate to float/decimal
        rate = float(row['rate'])
        
        # Prepare parameters for insert_fertilizer_event
        # Set liquidDensity to 8.3 for Manure Dairy Cattle Slurry, None for everything else
        liquid_density = float(8.3) if row['fert_type'] == "Manure Dairy Cattle Slurry" else None
        
        params = {
            'eventId': 11,
            'seasonIds': application_seasons,
            'doneAt': done_at,
            'applicationMethodId': application_method_id,
            'fertilizerId': fert_details['id'],
            'fertilizerCategoryId': fert_details['categoryId'],
            'rate': rate,
            'liquidDensity': liquid_density
        }
        
        # Insert fertilizer event with appropriate parameters
        insert_fertilizer_event(**params)

# function to process tillage events 
def process_tillage_data(excel_path, tillage_type_dict=None, tillage_residue_dict=None):
    """Process tillage data from Excel sheet."""
    # Read Excel file - specifically from the Tillage sheet
    df = pd.read_excel(excel_path, sheet_name='Tillage')
    
    # Get unique field-year combinations and get season IDs
    unique_field_years = df[['field_uuid', 'year']].drop_duplicates()
    season_ids_by_year_field = {}
    
    # Get season IDs and organize them by year and field
    for _, row in unique_field_years.iterrows():
        field_uuid = str(row['field_uuid'])
        year = int(row['year'])
        season_id = get_season_id(field_uuid, year)
        
        if year not in season_ids_by_year_field:
            season_ids_by_year_field[year] = {}
        season_ids_by_year_field[year][field_uuid] = season_id
    
    total_affected_rows = 0
    # Process each tillage event
    for _, row in df.iterrows():
        # Get the season ID for this specific year and field
        year = int(row['year'])
        field_uuid = str(row['field_uuid'])
        season_id = season_ids_by_year_field[year][field_uuid]
        
        # Format tillage date - handle timezone info
        if isinstance(row['tillage_date'], str):
            done_at = row['tillage_date'].split('T')[0]
        else:
            done_at = row['tillage_date'].strftime('%Y-%m-%d')

        # Set row/strip width only for Strip Till and Strip Till Freshener
        row_width = None
        strip_width = None
        if row['tillage_type'] in ["Strip Till", "Strip Till Freshener"]:
            # Use get() method to safely handle missing columns
            row_width = float(row.get('width_row')) if pd.notna(row.get('width_row')) else None
            strip_width = float(row.get('width_till')) if pd.notna(row.get('width_till')) else None

        # Get depth if provided
        tillage_depth = float(row['depth']) if pd.notna(row['depth']) else None
        
        # Insert tillage event
        affected_rows = insert_tillage_event(
            eventId=3,  # Event ID for tillage is 3
            seasonId=season_id,
            doneAt=done_at,
            tillage_name=row['tillage_type'],
            tillageDepth=tillage_depth,
            rowWidth=row_width,
            stripWidth=strip_width,
            residue_name=row['post_till_residue'] if pd.notna(row['post_till_residue']) else None,
            tillage_type_dict=tillage_type_dict,
            tillage_residue_dict=tillage_residue_dict
        )
        total_affected_rows += affected_rows if affected_rows is not None else 0
    
    print(f"\nTotal tillage events processed: {total_affected_rows}")
    return total_affected_rows

# function to run on all fields for a producer 
def process_producer_events(producer_ids, specific_year=None):
    with open('cpfrs_by_grower.json') as f:
        data = json.load(f)
    
    for producer_id in producer_ids:
        if producer_id not in data:
            print(f"Producer ID {producer_id} not found in JSON data.")
            continue
            
        print(f"\nProcessing producer: {producer_id}")
        fields = data[producer_id]
        
        for field in fields:
            field_id = field.get("partner_field_id")
            if field_id:
                print(f"\nProcessing field: {field_id}")
                parse_planting_events(producer_id, field_id, specific_year)
                parse_harvest_events(producer_id, field_id, specific_year)

################################################
# Example usage - Fertilizer application --- MANUAL / Hardcoded 

# Get season_ids - Austin Barrington
#'1be67eef-fe42-493d-b0d4-957b379d4621', 'name': 'anne-1'
# 'id': '97e81cfb-44cc-4dc2-a4fa-3ef99d4eaa3a', 'name': 'anne-2'  
#'ed4c5739-bf87-4348-86ea-c35e0e289c37', 'name': 'bad-demo-field'
# {'id': '0ddf6d8f-b960-4fc1-9da3-4ed2a5575a07', 'name': 'Custom Year_Demo'

# seasons = []
# fields = ['1be67eef-fe42-493d-b0d4-957b379d4621', '97e81cfb-44cc-4dc2-a4fa-3ef99d4eaa3a', 'ed4c5739-bf87-4348-86ea-c35e0e289c37', '0ddf6d8f-b960-4fc1-9da3-4ed2a5575a07']
# for f in fields:
#     season_id = get_season_id(f, 2024)
#     seasons.append(season_id)


# ##################
# # Usage example with necessary parameters
# seasons = seasons
# event_id = 11
# done_at = "2024-10-04"
# application_method_id = 3
# fertilizer_id = 3
# fertilizer_category_id = 'Organic'
# rate = 6000
# liquid_density = 8.3

# # Execute the function -- Insert Fert events
# insert_fertilizer_event(
#     eventId=event_id,
#     seasonIds=seasons,
#     doneAt=done_at,
#     applicationMethodId=application_method_id,
#     fertilizerId=fertilizer_id,
#     fertilizerCategoryId=fertilizer_category_id,
#     rate=rate,
#     liquidDensity=liquid_density
# )

######
# Example usage Tillage - Hardcoded 
# Usage examples
# if __name__ == "__main__":
   
#     # Get seasons for multiple fields
#     fields = [
#         '1be67eef-fe42-493d-b0d4-957b379d4621',
#         '97e81cfb-44cc-4dc2-a4fa-3ef99d4eaa3a', 
#         'ed4c5739-bf87-4348-86ea-c35e0e289c37',
#         '0ddf6d8f-b960-4fc1-9da3-4ed2a5575a07'
#     ]
    
#     seasons = []
#     for f in fields:
#         season_id = get_season_id(f, 2024)
#         seasons.append(season_id)

#     # Common parameters
#     done_at = "2024-10-04"
#     tillage_name = "Disc"
#     residue_name = "15-30%"
    
#     # For Strip Till and Strip Till Freshener, include striptill parameters
#     if tillage_name in ["Strip Till", "Strip Till Freshener"]:
#         for season_id in seasons:
#             insert_tillage_event(
#                 eventId=3,
#                 seasonId=season_id,
#                 doneAt=done_at,
#                 tillage_name=tillage_name,
#                 tillageDepth=10,  # Optional: override default depth
#                 rowWidth=30,      # Will be mapped to striptillCultivated
#                 stripWidth=8,     # Will be mapped to striptillWidth
#                 residue_name=residue_name,
#                 tillage_type_dict=tillage_type,
#                 tillage_residue_dict=tillage_residue
#             )
#     else:
#         # For other tillage types
#         for season_id in seasons:
#             insert_tillage_event(
#                 eventId=3,
#                 seasonId=season_id,
#                 doneAt=done_at,
#                 tillage_name=tillage_name,
#                 residue_name=residue_name,
#                 tillage_type_dict=tillage_type,
#                 tillage_residue_dict=tillage_residue
#             )

######################################################################
# Usage -- FROM EXCEL Sheet -- BULK UPLOAD 
if __name__ == "__main__":
    excel_path = 'mmrv_data.xlsx'
    fertilizers_json_path = 'fertilizers.json'
    commodities_json_path = 'commodities.json'
    
    # Process fertilizer data
    print("\nProcessing fertilizer data...")
    process_fertilizer_data(
        excel_path=excel_path,
        fertilizers_json_path=fertilizers_json_path
    )
    
    # Process planting data
    print("\nProcessing planting data...")
    planting_rows = process_planting_data(
        excel_path=excel_path,
        commodities_json_path=commodities_json_path
    )
    print(f"Total planting events processed: {planting_rows}")
    
    # Process harvest data
    print("\nProcessing harvest data...")
    harvest_rows = process_harvest_data(
        excel_path=excel_path,
        commodities_json_path=commodities_json_path
    )
    print(f"Total harvest events processed: {harvest_rows}")

    # Process tillage data
    tillage_rows = process_tillage_data(
        excel_path='mmrv_data.xlsx',
        tillage_type_dict=tillage_type,
        tillage_residue_dict=tillage_residue
    )

    print(f"Total tillage events processed: {tillage_rows}")

################################################
# Example Usage, Planting and Harvest - DO IT THIS WAY IF YOU'RE RUNNIGN SPECIFIC FIELDS AGI Data

# # Example dictionary of producers and their field IDs
# producers_fields = {
#     "7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c": ["3d72c0ce-c16c-48c0-be8c-384a0dcaff68", "76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84", "07206026-bd14-4732-b82a-1657dc68e3fa", '37a2f972-90c5-40c9-a5f1-56d5f8768e27'],
#     "63b36320-8e38-454f-97c9-e4dcb3510d61": ["fe325a21-4f88-42aa-8f81-122c9ca04aec", "9d8feb4b-d87b-49c3-99f1-73839623267f"],
#     "37459734-03ac-4b4f-95c0-bb4311beb121": ["ebfdbc2f-566a-4cf8-a6d6-19fece8ebe35", "6115fcad-56bf-44d9-9957-da03b6cc1a84", "ba6184ab-31ee-4c20-b87a-1682704bca7a"]
# }

# # Iterate over the producers and their field IDs
# for producer_id, field_ids in producers_fields.items():
#     for field_id in field_ids:
#         # Call to process only for the year 2024
#         parse_planting_events(producer_id, field_id, specific_year=2024)
#         parse_harvest_events(producer_id, field_id, specific_year=2024)

###########
# Example Usage, Planting and Harvest - DO IT THIS WAY IF YOU'RE RUNNING ALL FIELDS FOR A PRODUCER 

# Usage example -- All fields for a producer(s)
# producer_ids = [
#     "63b36320-8e38-454f-97c9-e4dcb3510d61",
#     "37459734-03ac-4b4f-95c0-bb4311beb121",
#     "ef920cb5-944a-4756-8b8d-6c5e20ab0f91",
#     '94477c71-1741-4e73-8ff3-b01ef68d9816'
# ]
# process_producer_events(producer_ids, specific_year=2024)

######
