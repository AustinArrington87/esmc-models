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
    "Sorghum": 10
}

# list projects 
#projects = mrv.projects()
#print(projects)
# get project ID 
#AGIprojId = mrv.projectId('AGI SGP Market')
#print(AGIprojId)

#AGIProducers = mrv.enrolledProducers(AGIprojId, 2023)
# "partner_grower_id" in AGI
# "partner_field_id" in AGI
#print(AGIProducers)

#AustinProducers = mrv.enrolledProducers('18a4db70-209d-4a93-87a9-ab3ff315fd14', 2024)
#print(AustinProducers)

#BarringtonFields = mrv.fieldSummary('1a8b2fd2-5a46-4d70-8a54-952871f43fca', 2024)
#print(BarringtonFields) 

# '7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c' - J. WILLIAMS 
# '63b36320-8e38-454f-97c9-e4dcb3510d61' - T. Ball
# '37459734-03ac-4b4f-95c0-bb4311beb121' - T. Langford 

williamsFields = mrv.fieldSummary('7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c', 2023)
#print(williamsFields)
# P Rinehart - '3d72c0ce-c16c-48c0-be8c-384a0dcaff68' ✓
# K Rinehart - '76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84' ✓
# Soybeans harvested 2023-10-09, 21.04690989 bu/ac + '' harvested 2024-06-16 213.1069567 bua/acre + Winter Wheat harvested 2024-06-17 10.8198256 bu/acre
# P Rinehart West - '07206026-bd14-4732-b82a-1657dc68e3fa' ✓ -- missing the partner_field_id, also empty yield Sorghum 
# Smith West - '37a2f972-90c5-40c9-a5f1-56d5f8768e27' ✓

ballFields = mrv.fieldSummary('63b36320-8e38-454f-97c9-e4dcb3510d61', 2023)
#print(ballFields)
# Krueger W - 'fe325a21-4f88-42aa-8f81-122c9ca04aec'
# Krueger E - '9d8feb4b-d87b-49c3-99f1-73839623267f'

langfordFields = mrv.fieldSummary('37459734-03ac-4b4f-95c0-bb4311beb121', 2023)
#print(langfordFields)
# East of Goldie - 'ebfdbc2f-566a-4cf8-a6d6-19fece8ebe35'
# Kevin Thomas - '6115fcad-56bf-44d9-9957-da03b6cc1a84'
# Singletons - 'ba6184ab-31ee-4c20-b87a-1682704bca7a'

# test Payload on Williams field "P Rinehart"
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

# Function to parse harvest events from JSON
def parse_harvest_events(producer_id, field_id, specific_year=None):
    with open('cpfrs_by_grower.json') as f:
        data = json.load(f)

    # Check if the producer_id exists in the JSON data
    if producer_id not in data:
        print(f"Producer ID {producer_id} not found in JSON data.")
        return

    fields = data[producer_id]
    found = False  # Flag to check if field_id is found

    for field in fields:
        if field.get("partner_field_id") == field_id:
            found = True
            print(f"Found field_id: {field_id} for producer: {producer_id}")
            harvest_events = field.get("field_events", [])
            if not harvest_events:
                print(f"No harvest events found for field_id: {field_id}")
                return  # Exit if no harvest events

            # Loop through harvest events
            for harvest in harvest_events:
                crop_cycle_summaries = harvest.get("crop_cycle_summaries", {})
                harvest_list = crop_cycle_summaries.get("harvest", [])
                
                for harvest in harvest_list:
                    end_date = harvest.get("end_date", "")
                    year = end_date.split("-")[0]  # Extract the year from the end_date

                    # Check if we should process this harvest based on the specific_year parameter
                    if specific_year and year != str(specific_year):
                        continue  # Skip this harvest if it doesn't match the specific year

                    # Get the season_id for the corresponding year
                    season_id = get_season_id(field_id, int(year))

                    avg_dry_yield = harvest.get("avg_dry_yield", {}).get("value", 0)

                    # Call insert_harvest_event with the extracted values
                    commodity_id = crop_type_to_id.get(harvest.get("crop_type", ""))
                    if commodity_id is not None and season_id is not None:
                        affected_rows = insert_harvest_event(
                            commodityId=commodity_id,
                            eventId=5,
                            doneAt=end_date.split("T")[0],  # Strip time
                            yield_value=avg_dry_yield,
                            seasonId=season_id
                        )
                        print(f"Affected Rows for {harvest.get('crop_type', '')} on {end_date}: {affected_rows}")
                    else:
                        print(f"Commodity ID not found for crop type: {harvest.get('crop_type', '')} or season_id is None")

    if not found:
        print(f"Field ID {field_id} not found for producer {producer_id} in JSON data.")


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

# Step 2: Function to insert fertilizer details with a generated UUID
def insert_fertilizer_details(event_ids, applicationMethodId, fertilizerId, fertilizerCategoryId, rate, liquidDensity):
    # Mutation to insert fertilizer details with a client-generated ID
    fertilizer_mutation = """
    mutation insertFertilizerDetails(
        $id: uuid,
        $applicationMethodId: smallint,
        $fertilizerId: smallint,
        $fertilizerCategoryId: String,
        $rate: numeric,
        $liquidDensity: numeric
    ) {
      insertFarmFertilizerData(objects: {
        id: $id,
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
    
    total_affected_rows = 0
    
    for seasonId, eventId in event_ids:
        # Generate a unique ID for each fertilizer entry
        unique_id = str(uuid.uuid4())
        
        variables = {
            "id": unique_id,
            "applicationMethodId": applicationMethodId,
            "fertilizerId": fertilizerId,
            "fertilizerCategoryId": fertilizerCategoryId,
            "rate": rate,
            "liquidDensity": liquidDensity
        }
        
        response = requests.post(url, json={'query': fertilizer_mutation, 'variables': variables}, headers=headers)
        
        if response.status_code == 200:
            response_data = response.json()
            if 'errors' in response_data:
                print(f"Failed to insert fertilizer details for seasonId {seasonId} with errors: {response_data['errors']}")
            else:
                affected_rows = response_data['data']['insertFarmFertilizerData']['affected_rows']
                total_affected_rows += affected_rows
                print(f"Fertilizer details added for seasonId {seasonId}! Rows affected: {affected_rows}")
        else:
            print(f"HTTP request failed for fertilizer data with status code {response.status_code}")
            print(response.text)
    
    print(f"Total fertilizer rows affected: {total_affected_rows}")
    return total_affected_rows



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
# Example usage of creating events and then inserting fertilizer details
season_ids = seasons  # Example season IDs
event_id = 11  # Event ID for application
done_at = "2024-10-04"  # Application date
application_method_id = 3  # Application method ID
fertilizer_id = 3  # Fertilizer ID
fertilizer_category_id = 'Organic'  # Fertilizer category
rate = 6000  # Application rate
liquid_density = 8.3  # Liquid density

# Step 1: Create events and get their IDs
created_event_ids = insert_event_multiple_seasons(event_id, season_ids, done_at)

# Step 2: Use event IDs to add fertilizer details
total_rows_affected = insert_fertilizer_details(
    created_event_ids,
    applicationMethodId=application_method_id,
    fertilizerId=fertilizer_id,
    fertilizerCategoryId=fertilizer_category_id,
    rate=rate,
    liquidDensity=liquid_density
)

print(f"Total rows affected in fertilizer details: {total_rows_affected}")
################################################

# Example usage -- Harvest events 
producer_id = "7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c"  # Example producer ID - AGI
field_ids = ["3d72c0ce-c16c-48c0-be8c-384a0dcaff68", "76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84", "07206026-bd14-4732-b82a-1657dc68e3fa", "37a2f972-90c5-40c9-a5f1-56d5f8768e27"]

# # Call the new function with producer_id and field_id
# for field_id in field_ids:
#     #parse_harvest_events(producer_id, field_id)  
#     # Call to process only for the year 2024
#     parse_harvest_events(producer_id, field_id, specific_year=2024)

# P Rinehart - '3d72c0ce-c16c-48c0-be8c-384a0dcaff68' ✓
# K Rinehart - '76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84' ✓
# Soybeans harvested 2023-10-09, 21.04690989 bu/ac + '' harvested 2024-06-16 213.1069567 bua/acre + Winter Wheat harvested 2024-06-17 10.8198256 bu/acre
# P Rinehart West - '07206026-bd14-4732-b82a-1657dc68e3fa' ✓ -- missing the partner_field_id, also empty yield Sorghum 
# Smith West - '37a2f972-90c5-40c9-a5f1-56d5f8768e27' ✓
