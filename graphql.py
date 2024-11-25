import mrvApi as mrv
import requests
import json
import uuid

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

# Mapping of crop types to commodity IDs
crop_type_to_id = {
    "Soybeans": 2,
    "Winter Wheat": 3,
    "Sorghum": 10,
    "Corn": 1,
    "Wheat": 25,
    "Oats": 8
}

# list projects 
#projects = mrv.projects()
#print(projects)
# get project ID 
AGIprojId = mrv.projectId('AGI SGP Market')
print(AGIprojId)

AGIProducers = mrv.enrolledProducers(AGIprojId, 2024)
# "partner_grower_id" in AGI
# "partner_field_id" in AGI
print(AGIProducers)

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
# Example usage - Fertilizer application 

# Get season_ids 
#'1be67eef-fe42-493d-b0d4-957b379d4621', 'name': 'anne-1'
# 'id': '97e81cfb-44cc-4dc2-a4fa-3ef99d4eaa3a', 'name': 'anne-2'  
#'ed4c5739-bf87-4348-86ea-c35e0e289c37', 'name': 'bad-demo-field'
# {'id': '0ddf6d8f-b960-4fc1-9da3-4ed2a5575a07', 'name': 'Custom Year_Demo'

seasons = []
fields = ['1be67eef-fe42-493d-b0d4-957b379d4621', '97e81cfb-44cc-4dc2-a4fa-3ef99d4eaa3a', 'ed4c5739-bf87-4348-86ea-c35e0e289c37', '0ddf6d8f-b960-4fc1-9da3-4ed2a5575a07']
for f in fields:
    season_id = get_season_id(f, 2024)
    seasons.append(season_id)


##################
# Usage example with necessary parameters
seasons = seasons
event_id = 11
done_at = "2024-10-04"
application_method_id = 3
fertilizer_id = 3
fertilizer_category_id = 'Organic'
rate = 6000
liquid_density = 8.3

# Execute the function -- Insert Fert events
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
################################################
# Example Usage, Planting and Harvest - DO IT THIS WAY IF YOU'RE RUNNIGN SPECIFIC FIELDS 

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
producer_ids = [
    "1b8781bf-3126-474d-93ea-a6d359d60f11"
]
process_producer_events(producer_ids, specific_year=2024)

######

# '1b8781bf-3126-474d-93ea-a6d359d60f11' C. Basinger ✓ (All fields rewritten 11/25)

# '7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c' - J. WILLIAMS ✓ (All fields rewritten 11/25)
# P Rinehart - '3d72c0ce-c16c-48c0-be8c-384a0dcaff68' ✓
# K Rinehart - '76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84' ✓
# Soybeans harvested 2023-10-09, 21.04690989 bu/ac + '' harvested 2024-06-16 213.1069567 bua/acre + Winter Wheat harvested 2024-06-17 10.8198256 bu/acre (removed from JSON)
# P Rinehart West - '07206026-bd14-4732-b82a-1657dc68e3fa' ✓ -- missing the partner_field_id, also empty yield Sorghum (removed from JSON)
# Smith West - '37a2f972-90c5-40c9-a5f1-56d5f8768e27' ✓

# '63b36320-8e38-454f-97c9-e4dcb3510d61' - T. Ball ✓
# Krueger W - 'fe325a21-4f88-42aa-8f81-122c9ca04aec' ✓ # Soybean planting and harvest event, some conflict with a corn planting but no harvest (removed from JSON)
# Krueger E - '9d8feb4b-d87b-49c3-99f1-73839623267f' ✓ # No Harvest events + Corn 2024-04-09, empty planting event 2024-05-10 (removed from JSON)

# '37459734-03ac-4b4f-95c0-bb4311beb121' - T. Langford ✓
# East of Goldie - 'ebfdbc2f-566a-4cf8-a6d6-19fece8ebe35' Winter Wheat planting and harvest + empty planting event 2024-05-09 (removed from JSON)
# Kevin Thomas - '6115fcad-56bf-44d9-9957-da03b6cc1a84' empty plant event 2023-10-10, winter wheat harvest (removd from JSON)
# Singletons - 'ba6184ab-31ee-4c20-b87a-1682704bca7a' empty plant event 2023-12-07 + Corn {planting }+ Sorghum harvest event but no sorghum planting (removed from JSON)
