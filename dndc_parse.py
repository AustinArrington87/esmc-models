import json
import csv

# Sample JSON data
file_path = 'TNCMN-final-3_verified.dndc.json'


# Load the JSON into a Python dictionary
with open(file_path, 'r') as file:
    data = json.load(file)

# Open a CSV file to write the parsed data
with open('dndc_raw_parsed.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    # Write the header of the CSV file
    writer.writerow(['Field', 'baseline_dsoc_raw', 'baseline_dsoc_sd', 'practice_dsoc_raw', 'practice_dsoc_sd'])

    # Iterate over each field in the session_uncertainties
    for field_key, field_value in data['data']['session_uncertainties'].items():
        field_name = field_key.rsplit('_', 1)[0]  # Remove the year from the field name
        # Retrieve data for baseline and practice_change scenarios
        baseline_dsoc_raw = field_value['scenarios_uncertainties']['baseline']['dsoc']['raw_output']
        baseline_dsoc_sd = field_value['scenarios_uncertainties']['baseline']['dsoc']['standard_deviation']
        practice_dsoc_raw = field_value['scenarios_uncertainties']['practice_change']['dsoc']['raw_output']
        practice_dsoc_sd = field_value['scenarios_uncertainties']['practice_change']['dsoc']['standard_deviation']

        # Write the data to the CSV file
        writer.writerow([field_name, baseline_dsoc_raw, baseline_dsoc_sd, practice_dsoc_raw, practice_dsoc_sd])
