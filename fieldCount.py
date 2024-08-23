import os
import geopandas as gpd
import pandas as pd

# List of project folders
project_folders = [
    "AGI SGP Market (AGIM)",
    "Benson Hill (BH)",
    "Corteva-Nutrien (C-NUT)",
    "Cotton Manulife (CTNMNU)",
    "Cotton USCTP (CTNUSCTP)",
    "John Deere (JD)",
    "Missouri Partnership Pilot (MOCS)",
    "NACD SGP Market (NACDSGP)",
    "Northern Plains Cropping (NPC)",
    "Syngenta-Nutrien (SN)",
    "TNC Minnesota (TNCMN)",
    "TNC Nebraska (TNCNE)"
]

def count_unique_field_ids_in_project_folders():
    results = []
    total_monitoring_fields = 0
    total_modeling_fields = 0

    # Assume the current working directory is the base directory
    base_directory = os.getcwd()

    for project in project_folders:
        project_dir = os.path.join(base_directory, project)
        if os.path.isdir(project_dir):
            for filename in os.listdir(project_dir):
                if filename.endswith(".geojson"):
                    file_path = os.path.join(project_dir, filename)
                    geo_data = gpd.read_file(file_path)

                    # Count unique field IDs
                    unique_field_count = geo_data['fieldId'].nunique()
                    results.append({
                        'Project': project,
                        'File': filename,
                        'Unique Field Count': unique_field_count
                    })

                    # Determine if the file is for monitoring or modeling
                    if "monitoring" in filename.lower():
                        total_monitoring_fields += unique_field_count
                    elif "modeled" in filename.lower():
                        total_modeling_fields += unique_field_count

                    print(f"Project: {project}, File: {filename}, Unique Fields: {unique_field_count}")

    # Export results to CSV
    results_df = pd.DataFrame(results)
    results_df.to_csv(os.path.join(base_directory, 'field_counts.csv'), index=False)

    # Print totals
    print(f"Total Monitoring Fields: {total_monitoring_fields}")
    print(f"Total Modeling Fields: {total_modeling_fields}")

# Run the function
count_unique_field_ids_in_project_folders()
