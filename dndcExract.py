import json
import csv
import re

def process_dndc_data(dndcJson, output_csv):
    field_data = []

    for field in dndcJson['data']['session_uncertainties']:
        field_info = {}
        # Strip "_2023" from the field name
        field_info['field'] = re.sub(r'_2023$', '', field)

        field_info['baseline_dsoc_raw'] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['dsoc']['raw_output']
        field_info['practice_dsoc_raw'] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['dsoc']['raw_output']
        field_info["dsoc_raw_delta"] = field_info['practice_dsoc_raw'] - field_info['baseline_dsoc_raw']
        field_info['baseline_direct_n2o_raw'] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['direct_n2o']['raw_output']
        field_info['practice_direct_n2o_raw'] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['direct_n2o']['raw_output']
        field_info['n2o_direct_raw_delta'] = field_info['baseline_direct_n2o_raw'] - field_info['practice_direct_n2o_raw']

        field_data.append(field_info)

    with open(output_csv, 'w', newline='') as csvfile:
        fieldnames = ['field', 'baseline_dsoc_raw', 'practice_dsoc_raw', 'dsoc_raw_delta', 'baseline_direct_n2o_raw', 'practice_direct_n2o_raw', 'n2o_direct_raw_delta']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for data in field_data:
            writer.writerow(data)

if __name__ == "__main__":
    # Example usage
    dndc_filename = 'NACD SGP Market-2023-Batch1B-Final2-V2.dndc.json'
    output_csv = 'NACD_SGP_DNDC_Raw.csv'

    try:
        with open(dndc_filename) as file:
            dndcJson = json.load(file)
        
        process_dndc_data(dndcJson, output_csv)
        print(f"CSV file '{output_csv}' created successfully.")
    except Exception as e:
        print(f"Error processing DNDC JSON file: {e}")
