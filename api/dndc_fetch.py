import requests
import json

API_KEY = "EnterKey"
URL = "https://api.regrow.ag/dndc-scenarios-service/v0/api/projects"

headers = {
    "x-apikey": API_KEY,
    "Content-Type": "application/json"
}

response = requests.get(URL, headers=headers)

if response.status_code == 200:
    projects = response.json().get("projects", [])
    for project in projects:
        print(f"Project Name: {project.get('project_name')}")
        print(f"Created At: {project.get('created_at')}")
        print(f"Protocol: {project.get('protocol')}")
        print(f"Is Finalized: {project.get('is_finalized')}")
        print(f"Results Available: {project.get('results_available')}")
        print("-" * 40)
else:
    print(f"Failed to fetch projects. Status code: {response.status_code}")
    print(response.text)

# Try different output endpoints for the TNC Minnesota project
project_name = "TNC_Minnesota-2024-Batch-1-Final-3"

print(f"\nTrying different output endpoints for: {project_name}")
print("=" * 60)

# List of different output endpoints to try
endpoints = [
    #"project_level_scope_3",
    #"project_level_scope_1_2", 
    #"project_level_scope_2",
    #"field_level_results",
    #"project_summary",
    "project_level_scope_1"
]

for endpoint in endpoints:
    URL = f"https://api.regrow.ag/dndc-scenarios-service/v0/api/{endpoint}"
    
    params = {
        "project_name": project_name,
    }
    
    print(f"\nTrying endpoint: {endpoint}")
    response = requests.get(URL, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()
        print(f"✅ SUCCESS! Status code: {response.status_code}")
        
        # Save data to JSON file
        filename = f"{project_name}_{endpoint}.json"
        with open(filename, 'w') as f:
            json.dump(data, f, indent=2)
        
        print(f"✅ Data saved to: {filename}")
        print(f"Response keys: {list(data.keys()) if isinstance(data, dict) else 'Not a dict'}")
        break
    else:
        print(f"❌ Failed. Status code: {response.status_code}")
        print(f"Error: {response.text}")
