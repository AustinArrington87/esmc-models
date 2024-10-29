import mrvApi as mrv
import requests

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


# list projects 
projects = mrv.projects()
#print(projects)
# get project ID 
AGIprojId = mrv.projectId('AGI SGP Market')
#print(AGIprojId)

AGIProducers = mrv.enrolledProducers(AGIprojId, 2023)
# "partner_grower_id" in AGI
# "partner_field_id" in AGI

#print(AGIProducers)

# '7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c' - J. WILLIAMS 
# '63b36320-8e38-454f-97c9-e4dcb3510d61' - T. Ball
# '37459734-03ac-4b4f-95c0-bb4311beb121' - T. Langford 

williamsFields = mrv.fieldSummary('7b227522-8f36-4ecf-8bb7-1ad98d4d6c8c', 2023)
#print(williamsFields)
# P Rinehart - '3d72c0ce-c16c-48c0-be8c-384a0dcaff68' 
# K Rinehart - '76e43cc9-2e2a-47e8-a5ae-baaefbdf0c84'
# P Rinehart West - '07206026-bd14-4732-b82a-1657dc68e3fa'
# Smith West - '37a2f972-90c5-40c9-a5f1-56d5f8768e27'

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

# Function to retrieve season_id based on fieldId
def get_season_id(fieldId):
    # Define the query with a variable for `fieldId`
    query = """
    query MyQuery($fieldId: uuid!) {
      farmSeason(where: {fieldId: {_eq: $fieldId}, year: {_eq: "2024"}}) {
        id
      }
    }
    """
    
    # Define the variables
    variables = {
        "fieldId": fieldId
    }

    # Send the request with the query and variables
    response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)

    # Process the response
    if response.status_code == 200:
        print("Query successful!")
        response_data = response.json()
        season_id = response_data['data']['farmSeason'][0]['id'] if response_data['data']['farmSeason'] else None
        print(f"Season ID: {season_id}")
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
        print(f"Affected Rows: {affected_rows}")
        return affected_rows
    else:
        print(f"Mutation failed with status code {response.status_code}")
        print(response.text)
        return None

################################################


# Example usage
field_id = "3d72c0ce-c16c-48c0-be8c-384a0dcaff68"
season_id = get_season_id(field_id)
print(season_id)
affected_rows = insert_harvest_event(commodityId=2, eventId=5, doneAt="2024-10-03", yield_value=23.4705479706396, seasonId=season_id)
print(f"Affected Rows: {affected_rows}")


