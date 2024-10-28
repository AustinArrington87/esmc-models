import requests

# STEP ONE --> pass in the fieldID and get the seasonID back 

# Define the GraphQL endpoint
url = "https://graphql.ecoharvest.ag/v1/graphql"

# Define the query with a variable for `fieldId`
query = """
query MyQuery($fieldId: uuid!) {
  farmSeason(where: {fieldId: {_eq: $fieldId}, year: {_eq: "2024"}}) {
    id
  }
}
"""

# Your Hasura admin secret key
admin_secret_key = "ENTER SECRET"

# Set up the headers with the Hasura admin secret
headers = {
    "Content-Type": "application/json",
    "x-hasura-admin-secret": admin_secret_key
}

# Define the variables
variables = {
    "fieldId": "3d72c0ce-c16c-48c0-be8c-384a0dcaff68"
}

# Send the request with the query and variables
response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)


# Print the response
if response.status_code == 200:
    print("Query successful!")
    print(response.json())
    season_id = response.json()['data']['farmSeason'][0]['id']
else:
    print(f"Query failed with status code {response.status_code}")
    print(response.text)
    seson_id = None

print(season_id)

########################
# STEP TWO -- Add Harvest Event Using Mutation 

# Define the mutation with variables
harvest_mutation = """
mutation insertEventData($commodityId: Int, $eventId: Int, $doneAt: date, $yield: numeric, $seasonId: uuid) {
  insertFarmEventData(objects: {commodityId: $commodityId, eventId: $eventId, doneAt: $doneAt, yield: $yield, seasonId: $seasonId}) {
    affected_rows
  }
}
"""


# Define the variables
harvest_variables = {
    "commodityId": 2,
    "eventId": 5,
    "doneAt": "2024-10-03",
    "yield": 23.46,
    "seasonId": season_id
}

# Send the request with the mutation and variables
response = requests.post(url, json={'query': harvest_mutation, 'variables': harvest_variables}, headers=headers)

# Print the response
if response.status_code == 200:
    print("Mutation successful!")
    print(response.json())
else:
    print(f"Mutation failed with status code {response.status_code}")
    print(response.text)



