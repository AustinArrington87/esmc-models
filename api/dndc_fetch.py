import requests
import json
import csv
import statistics

API_KEY = "Apikey"
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

endpoints = ["project_level_scope_1"]

for endpoint in endpoints:
    URL = f"https://api.regrow.ag/dndc-scenarios-service/v0/api/{endpoint}"
    params = {"project_name": project_name}
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

        # --- FLATTEN FIELD-LEVEL DATA TO CSV ---
        field_level = data.get('field_level', {})
        csv_rows = []
        dsoc_deltas = []
        direct_n2o_deltas = []
        indirect_n2o_deltas = []
        dsoc_fields = []
        direct_n2o_fields = []
        indirect_n2o_fields = []
        for field_id, field_data in field_level.items():
            row = {}
            row['field_name'] = field_data.get('field_name')
            ri = field_data.get('reporting_information', {})
            row['start_date'] = ri.get('start_date')
            row['end_date'] = ri.get('end_date')
            row['acres'] = ri.get('acres')
            outcomes = field_data.get('outcomes', {})
            # Reversible outcomes
            rev = outcomes.get('reversible_outcomes', {})
            dsoc = rev.get('dsoc', {})
            row['dsoc_baseline'] = dsoc.get('baseline')
            row['dsoc_practice'] = dsoc.get('practice_change')
            # Calculate dsoc_delta as dsoc_practice - dsoc_baseline
            dsoc_baseline = dsoc.get('baseline')
            dsoc_practice = dsoc.get('practice_change')
            row['dsoc_delta'] = dsoc_practice - dsoc_baseline if dsoc_baseline is not None and dsoc_practice is not None else None
            # Non-reversible outcomes
            nonrev = outcomes.get('non_reversible_outcomes', {})
            direct_n2o = nonrev.get('direct_n2o', {})
            row['direct_n2o_baseline'] = direct_n2o.get('baseline')
            row['direct_n2o_practice'] = direct_n2o.get('practice_change')
            row['direct_n2o_delta'] = direct_n2o.get('preliminary_credit')
            # Indirect n2o
            indirect_n2o = nonrev.get('indirect_n2o', {})
            row['indirect_n2o_baseline'] = indirect_n2o.get('baseline')
            row['indirect_n2o_practice'] = indirect_n2o.get('practice_change')
            row['indirect_n2o_delta'] = indirect_n2o.get('credit')
            csv_rows.append(row)
            # Collect for stats
            if row['dsoc_delta'] is not None:
                dsoc_deltas.append(row['dsoc_delta'])
                dsoc_fields.append((row['field_name'], row['dsoc_delta']))
            if row['direct_n2o_delta'] is not None:
                direct_n2o_deltas.append(row['direct_n2o_delta'])
                direct_n2o_fields.append((row['field_name'], row['direct_n2o_delta']))
            if row['indirect_n2o_delta'] is not None:
                indirect_n2o_deltas.append(row['indirect_n2o_delta'])
                indirect_n2o_fields.append((row['field_name'], row['indirect_n2o_delta']))
        csv_filename = f"{project_name}_{endpoint}_fields.csv"
        fieldnames = [
            'field_name', 'start_date', 'end_date', 'acres',
            'dsoc_baseline', 'dsoc_practice', 'dsoc_delta',
            'direct_n2o_baseline', 'direct_n2o_practice', 'direct_n2o_delta',
            'indirect_n2o_baseline', 'indirect_n2o_practice', 'indirect_n2o_delta'
        ]
        with open(csv_filename, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in csv_rows:
                writer.writerow(row)
        print(f"✅ Field-level data saved to: {csv_filename}")

        # --- RECALCULATE TOTALS AS REQUESTED ---
        # non-reversible outcomes = sum of all fields dsoc_delta
        # reversible outcomes = sum of all fields direct_n2o_delta + indirect_n2o_delta
        total_non_reversible = sum(dsoc_deltas)
        total_reversible = sum(direct_n2o_deltas) + sum(indirect_n2o_deltas)
        total_impacts = total_non_reversible + total_reversible

        print("\nProject Name:", project_name)
        print("Total Non-Reversible Outcomes (sum dsoc_delta):", total_non_reversible)
        print("Total Reversible Outcomes (sum direct_n2o_delta + indirect_n2o_delta):", total_reversible)
        print("Total Impacts:", total_impacts)

        # --- OUTLIER DETECTION ---
        def find_outliers(values, fields, label):
            if len(values) < 2:
                print(f"Not enough data for outlier detection for {label}.")
                return []
            mean = statistics.mean(values)
            stdev = statistics.stdev(values)
            outliers = [fname for fname, val in fields if abs(val - mean) > 2 * stdev]
            print(f"Outliers for {label} (>|2 std dev| from mean): {outliers}")
            print(f"Number of fields: {len(values)}")
            print(f"Number of fields with outliers: {len(outliers)}")
            percent = (len(outliers) / len(values)) * 100 if values else 0
            print(f"Percentage of fields with outliers: {percent:.2f}%\n")
            return outliers

        find_outliers(dsoc_deltas, dsoc_fields, 'dsoc_delta')
        find_outliers(direct_n2o_deltas, direct_n2o_fields, 'direct_n2o_delta')
        find_outliers(indirect_n2o_deltas, indirect_n2o_fields, 'indirect_n2o_delta')

        break
    else:
        print(f"❌ Failed. Status code: {response.status_code}")
        print(f"Error: {response.text}")
