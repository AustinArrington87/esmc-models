import requests
import json

# Step 1: Get OAuth token
auth_url = 'https://sustainability-api.farmobile.com/api/v1/oauth2/token'
auth_data = {
    'username': '',
    'password': '',
    'client_id': '',
    'client_secret': '',
    'grant_type': 'password'
}

response = requests.post(auth_url, data=auth_data)
response.raise_for_status()
token = response.json()['access_token']

# Step 2: Get list of partner growers
growers_url = 'https://sustainability-api.farmobile.com/api/v1/partners/esmc/growers'
headers = {'Authorization': f'Bearer {token}'}

response = requests.get(growers_url, headers=headers)
response.raise_for_status()
growers = response.json()

# Step 3: Get list of enrollments for each grower
enrollments_by_grower = {}

for grower in growers:
    partner_grower_id = grower.get('grower_id')
    grower_name = grower.get('name')
    if not partner_grower_id:
        print("Skipping grower with missing ID")
        continue

    enrollments_url = f'https://sustainability-api.farmobile.com/api/v1/growers/{partner_grower_id}/enrollments'

    response = requests.get(enrollments_url, headers=headers)
    if response.status_code == 404:
        print(f"Enrollments not found for grower ID {partner_grower_id}")
        continue

    response.raise_for_status()
    enrollments = response.json()

    for enrollment in enrollments:
        enrollment.update({
            "partner_grower_id": partner_grower_id,
            "grower_id": enrollment.get('grower_id'),
            "grower_name": grower_name,
            "partner_field_id": enrollment.get('field_id')
        })

    enrollments_by_grower[partner_grower_id] = enrollments

# Step 4: Get CPFRs for each enrollment
cpfrs_by_grower = {}

for partner_grower_id, enrollments in enrollments_by_grower.items():
    cpfrs_by_grower[partner_grower_id] = []

    for enrollment in enrollments:
        enrollment_id = enrollment['id']
        cpfrs_url = f'https://sustainability-api.farmobile.com/api/v1/enrollments/{enrollment_id}/cpfrs'

        response = requests.get(cpfrs_url, headers=headers)
        if response.status_code == 404:
            print(f"CPFRs not found for enrollment ID {enrollment_id}")
            continue

        response.raise_for_status()
        cpfrs = response.json()

        enrollment_data = {
            "enrollment_id": enrollment_id,
            "partner_grower_id": enrollment["partner_grower_id"],
            "grower_id": enrollment["grower_id"],
            "grower_name": enrollment["grower_name"],
            "partner_field_id": enrollment["partner_field_id"],
            "field_events_summary": enrollment.get("field_events_summary", {}),
            "field_events": cpfrs
        }

        cpfrs_by_grower[partner_grower_id].append(enrollment_data)

# Save cpfrs_by_grower as a JSON document
with open('cpfrs_by_grower.json', 'w') as json_file:
    json.dump(cpfrs_by_grower, json_file, indent=4)

print("CPFRs data has been saved to cpfrs_by_grower.json")
