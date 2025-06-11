import mrvApi as mrv
import requests
import json
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Any
import time
import os
import sys
from contextlib import contextmanager

# Configure API connection
c = mrv.configure(".env.production")

# Your Hasura admin secret key
admin_secret_key = "EnterKey"

# Set up the headers with the Hasura admin secret
headers = {
    "Content-Type": "application/json",
    "x-hasura-admin-secret": admin_secret_key
}

# Define the GraphQL endpoint
url = "https://graphql.ecoharvest.ag/v1/graphql"

# GMI Project IDs
GMI_PROJECTS = {
    "AGI_SGP_Market": "47e36083-ffa6-4237-8b13-0d9f50475836",
    "AGI_SGP_Pipeline": "0d6e8249-56e1-41a7-aeb4-9f562113625c",
    "NACD_SGP_Market": "9998ac10-7a6e-40f1-aa5b-cf2338b71ac6",
    "NGP_NACD_Market": "e8e6c8d3-401d-493a-ad8c-14da14b076e",
    "NGP_NACD_Pipeline": "a8c83ee3-8409-4b70-8a18-e1e1fb76ba84"
}

# Years to analyze
ANALYSIS_YEARS = list(range(2017, 2025))  # 2017 to 2024

# Global variable for log file
log_file = None

# Add this at the top of the file after the imports
TILLAGE_TYPE_MAP = {
    1: "Chisel",
    16: "Crimp",
    2: "Cultivator",
    3: "Deep Ripping",
    4: "Disc",
    17: "Flail",
    27: "Harrow",
    28: "Harrow (heavy)",
    29: "Harrow (light)",
    18: "Leveler",
    5: "Mulch Tiller",
    6: "Offset (heavy) Disc",
    15: "Other",
    7: "Plow",
    19: "Rake",
    8: "Ridge Till",
    24: "Rod Weed",
    20: "Roller",
    9: "Rotary Hoe",
    10: "Row Cultivator",
    11: "Strip Till",
    12: "Strip Till Freshener",
    25: "Subsoil",
    26: "Sweep",
    13: "Tandem (light) Disc",
    14: "Vertical Tillage"
}

def load_fertilizer_mapping():
    """Load fertilizer mapping from JSON file."""
    try:
        with open('fertilizers.json', 'r') as f:
            fertilizers = json.load(f)
        return {fert['id']: {
            'name': fert['name'],
            'type': fert['fertilizerTypeName'],
            'category': fert['fertilizerCategoryId']
        } for fert in fertilizers}
    except Exception as e:
        log_progress(f"Error loading fertilizer mapping: {str(e)}")
        return {}

# Load fertilizer mapping at module level
FERTILIZER_MAP = load_fertilizer_mapping()

@contextmanager
def setup_logging():
    """Context manager to handle logging setup and cleanup."""
    global log_file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = "gmi_analysis_results"
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Open log file
    log_file = open(f"{output_dir}/analysis_log_{timestamp}.txt", 'w')
    
    try:
        yield
    finally:
        if log_file:
            log_file.close()
            log_file = None

def log_progress(message: str, indent: int = 0):
    """Helper function to log progress with timestamp and indentation."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    indent_str = "  " * indent
    formatted_message = f"[{timestamp}] {indent_str}{message}"
    
    # Print to console
    print(formatted_message)
    
    # Write to log file if it exists
    if log_file:
        log_file.write(formatted_message + '\n')
        log_file.flush()  # Ensure immediate writing to file

def get_season_id(field_id: str, year: int) -> Optional[str]:
    """Get season ID for a field and year."""
    query = """
    query MyQuery($fieldId: uuid!, $year: smallint!) {
      farmSeason(where: {fieldId: {_eq: $fieldId}, year: {_eq: $year}}) {
        id
      }
    }
    """
    
    variables = {
        "fieldId": field_id,
        "year": year
    }

    try:
        response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)
        response.raise_for_status()
        
        response_data = response.json()
        if response_data['data']['farmSeason']:
            return response_data['data']['farmSeason'][0]['id']
        return None
    except Exception as e:
        log_progress(f"Error getting season ID for field {field_id}, year {year}: {str(e)}")
        return None

def get_field_events(season_id: str) -> List[Dict[str, Any]]:
    """Get all events for a season including cover crops, tillage, fertilizer, and pesticide data."""
    query = """
    query GetFieldEvents($seasonId: uuid!) {
      farmEventData(where: {seasonId: {_eq: $seasonId}}) {
        id
        eventId
        doneAt
        coverCropSeedingRate
        coverCropTerminationTypeId
        isCoverCropSeededAerially
        tillageDepth
        tillageDataId
        fertilizerDataId
        pesticideCount
        many_event_data_has_many_tillage_types {
          tillageTypeId
        }
        fertilizer_datum {
          applicationMethodId
          fertilizerId
          fertilizerCategoryId
          rate
          injectionDepth
        }
      }
    }
    """
    
    variables = {
        "seasonId": season_id
    }

    try:
        response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)
        response.raise_for_status()
        
        response_data = response.json()
        
        if 'errors' in response_data:
            log_progress(f"GraphQL errors: {response_data['errors']}")
            return []
            
        if 'data' in response_data and 'farmEventData' in response_data['data']:
            events = response_data['data']['farmEventData']
            log_progress(f"Retrieved {len(events)} events for season {season_id}")
            return events
        else:
            log_progress(f"Unexpected response structure: {response_data}")
            return []
            
    except Exception as e:
        log_progress(f"Error getting field events: {str(e)}")
        return []

def get_project_details(project_id: str) -> Dict[str, Any]:
    """Get detailed project information including fields and farmers."""
    query = """
    query GetProjectDetails($projectId: uuid!) {
      esmcProject(where: {id: {_eq: $projectId}}) {
        id
        name
        abbreviation
        isActive
        contractYears
        currentEnrollmentYear
        historicalYears
        farmer_projects {
          id
          farmerAppUserId
          farmer_project_fields {
            field {
              id
              acres
              boundary
            }
          }
        }
      }
    }
    """
    
    variables = {
        "projectId": project_id
    }
    
    try:
        response = requests.post(url, json={'query': query, 'variables': variables}, headers=headers)
        response.raise_for_status()
        
        response_data = response.json()
        
        if 'errors' in response_data:
            print(f"GraphQL errors: {response_data['errors']}")
            return {}
            
        if 'data' in response_data and 'esmcProject' in response_data['data']:
            return response_data['data']['esmcProject'][0] if response_data['data']['esmcProject'] else {}
        else:
            print(f"Unexpected response structure: {response_data}")
            return {}
            
    except Exception as e:
        print(f"Error getting project details: {e}")
        return {}

def get_project_producers(project_id: str, year: int) -> List[Dict[str, Any]]:
    """Get list of producers enrolled in a project for a given year."""
    log_progress(f"Getting producers for project {project_id}, year {year}")
    producers = mrv.enrolledProducers(project_id, year)
    log_progress(f"Found {len(producers)} producers")
    return producers

def get_producer_fields(producer_id: str, year: int) -> List[Dict[str, Any]]:
    """Get list of fields for a producer in a given year."""
    log_progress(f"Getting fields for producer {producer_id}, year {year}")
    field_summary = mrv.fieldSummary(producer_id, year)
    fields = field_summary.get('fields', [])
    log_progress(f"Found {len(fields)} fields")
    return fields

def analyze_project_data(project_id: str) -> Dict[str, Any]:
    """Analyze data for a specific project across all years."""
    log_progress(f"Starting analysis for project {project_id}")
    project_data = {
        'years': {},
        'total_fields': 0,
        'total_producers': 0
    }
    
    # Initialize year data
    for year in ANALYSIS_YEARS:
        project_data['years'][year] = {
            'producers': 0,
            'fields': 0,
            'total_acres': 0.0,
            'cover_crop_acres': 0.0,
            'fields_with_cover_crops': 0,
            'fields_without_cover_crops': 0,
            'tillage_events': 0,
            'fertilizer_applications': 0,
            'pesticide_applications': 0
        }
    
    # Get current producers (2024)
    log_progress("Getting current producers (2024)")
    current_producers = get_project_producers(project_id, 2024)
    if not current_producers:
        log_progress("No current producers found")
        return project_data
    
    log_progress(f"Found {len(current_producers)} current producers")
    project_data['total_producers'] = len(current_producers)
    
    # Process each producer
    for i, producer in enumerate(current_producers, 1):
        log_progress(f"Processing producer {i}/{len(current_producers)}: {producer.get('name', producer['id'])}")
        
        # Get current fields
        current_fields = get_producer_fields(producer['id'], 2024)
        if not current_fields:
            log_progress("No fields found for producer", indent=1)
            continue
            
        log_progress(f"Found {len(current_fields)} fields", indent=1)
        project_data['total_fields'] += len(current_fields)
        
        # Process each field's historical data
        for j, field in enumerate(current_fields, 1):
            log_progress(f"Processing field {j}/{len(current_fields)}: {field.get('name', field['id'])}", indent=2)
            
            # Check each year for this field
            for year in ANALYSIS_YEARS:
                log_progress(f"Checking year {year}", indent=3)
                season_id = get_season_id(field['id'], year)
                if not season_id:
                    log_progress("No season ID found, skipping", indent=4)
                    continue
                    
                events = get_field_events(season_id)
                if not events:
                    log_progress("No events found, skipping", indent=4)
                    continue
                
                # Update year data
                year_data = project_data['years'][year]
                year_data['fields'] += 1
                year_data['producers'] = max(year_data['producers'], 1)  # Count producer only once per year
                year_data['total_acres'] += float(field.get('acres', 0))
                
                has_cover_crop = False
                for event in events:
                    # Count cover crop events
                    if event.get('coverCropSeedingRate'):
                        has_cover_crop = True
                        year_data['cover_crop_acres'] += float(field.get('acres', 0))
                    
                    # Count tillage events
                    if event.get('tillageDepth') or event.get('tillageDataId'):
                        year_data['tillage_events'] += 1
                    
                    # Count fertilizer applications
                    if event.get('fertilizerDataId'):
                        year_data['fertilizer_applications'] += 1
                    
                    # Count pesticide applications
                    if event.get('pesticideCount'):
                        year_data['pesticide_applications'] += event['pesticideCount']
                
                if has_cover_crop:
                    year_data['fields_with_cover_crops'] += 1
                else:
                    year_data['fields_without_cover_crops'] += 1
                
                log_progress(f"Year {year} summary for field:", indent=4)
                log_progress(f"Cover crop acres: {year_data['cover_crop_acres']:.1f}", indent=5)
                log_progress(f"Tillage events: {year_data['tillage_events']}", indent=5)
                log_progress(f"Fertilizer applications: {year_data['fertilizer_applications']}", indent=5)
                log_progress(f"Pesticide applications: {year_data['pesticide_applications']}", indent=5)
    
    return project_data

def generate_project_report(project_data: Dict[str, Any], project_name: str) -> str:
    """Generate a report for a specific project."""
    report = f"\n{project_name} Analysis Report"
    report += "\n" + "=" * 50
    
    # Create table headers
    report += "\nYear | Producers | Fields | Total Acres | Cover Crop Acres | Cover Crop % | Fields w/ CC | Fields w/o CC | Tillage Events | Fertilizer Apps | Pesticide Apps"
    report += "\n" + "-" * 50
    
    # Data rows
    for year in sorted(project_data['years'].keys()):
        data = project_data['years'][year]
        cover_crop_percentage = (data['cover_crop_acres'] / data['total_acres'] * 100) if data['total_acres'] > 0 else 0.0
        
        report += f"\n{year} | {data['producers']} | {data['fields']} | {data['total_acres']:.1f} | {data['cover_crop_acres']:.1f} | {cover_crop_percentage:.1f}% | {data['fields_with_cover_crops']} | {data['fields_without_cover_crops']} | {data['tillage_events']} | {data['fertilizer_applications']} | {data['pesticide_applications']}"
    
    return report

def analyze_combined_data(all_projects_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    """Analyze combined data across all projects."""
    combined_data = {
        'years': {},
        'total_fields': 0,
        'total_producers': 0
    }
    
    # Initialize year data
    for year in ANALYSIS_YEARS:
        combined_data['years'][year] = {
            'producers': 0,
            'fields': 0,
            'total_acres': 0.0,
            'cover_crop_acres': 0.0,
            'fields_with_cover_crops': 0,
            'fields_without_cover_crops': 0,
            'tillage_events': 0,
            'fertilizer_applications': 0,
            'pesticide_applications': 0
        }
    
    # Combine data from all projects
    for project_id, project_data in all_projects_data.items():
        for year, year_data in project_data['years'].items():
            if year in combined_data['years']:
                # Add producers and fields
                combined_data['years'][year]['producers'] = max(
                    combined_data['years'][year]['producers'],
                    year_data.get('producers', 0)
                )
                combined_data['years'][year]['fields'] = max(
                    combined_data['years'][year]['fields'],
                    year_data.get('fields', 0)
                )
                
                # Add acres
                combined_data['years'][year]['total_acres'] += year_data.get('total_acres', 0.0)
                combined_data['years'][year]['cover_crop_acres'] += year_data.get('cover_crop_acres', 0.0)
                
                # Add fields with/without cover crops
                combined_data['years'][year]['fields_with_cover_crops'] = max(
                    combined_data['years'][year]['fields_with_cover_crops'],
                    year_data.get('fields_with_cover_crops', 0)
                )
                combined_data['years'][year]['fields_without_cover_crops'] = max(
                    combined_data['years'][year]['fields_without_cover_crops'],
                    year_data.get('fields_without_cover_crops', 0)
                )
                
                # Add events
                combined_data['years'][year]['tillage_events'] += year_data.get('tillage_events', 0)
                combined_data['years'][year]['fertilizer_applications'] += year_data.get('fertilizer_applications', 0)
                combined_data['years'][year]['pesticide_applications'] += year_data.get('pesticide_applications', 0)
        
        combined_data['total_fields'] = max(combined_data['total_fields'], project_data['total_fields'])
        combined_data['total_producers'] = max(combined_data['total_producers'], project_data['total_producers'])
    
    return combined_data

def calculate_trends(data: Dict[str, Any]) -> Dict[str, Any]:
    """Calculate year-over-year trends for key metrics."""
    years = sorted(data['years'].keys())
    if len(years) < 2:
        return {}
    
    trends = {
        'cover_crop_acres': [],
        'tillage_events': [],
        'fertilizer_applications': [],
        'pesticide_applications': []
    }
    
    # Calculate year-over-year changes
    for i in range(1, len(years)):
        prev_year = years[i-1]
        curr_year = years[i]
        
        prev_data = data['years'][prev_year]
        curr_data = data['years'][curr_year]
        
        # Calculate percentage changes
        for metric in trends.keys():
            if prev_data[metric] > 0:
                change = ((curr_data[metric] - prev_data[metric]) / prev_data[metric]) * 100
                trends[metric].append({
                    'year': curr_year,
                    'change': change,
                    'previous': prev_data[metric],
                    'current': curr_data[metric]
                })
    
    return trends

def generate_trend_report(trends: Dict[str, Any], project_name: str) -> str:
    """Generate a report showing trends for key metrics."""
    report = f"\n{project_name} Practice Trends (2017-2024)"
    report += "\n" + "=" * 50 + "\n"
    
    for metric, changes in trends.items():
        if not changes:
            continue
            
        report += f"\n{metric.replace('_', ' ').title()}:\n"
        report += "-" * 30 + "\n"
        
        for change in changes:
            report += f"Year {change['year']}: {change['change']:+.1f}% "
            report += f"({change['previous']:.1f} → {change['current']:.1f})\n"
        
        report += "\n"
    
    return report

def export_to_csv(all_projects_data: Dict[str, Dict[str, Any]], combined_data: Dict[str, Any]) -> None:
    """Export analysis results to Excel file with multiple sheets."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = "gmi_analysis_results"
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Create Excel writer
    excel_filename = f"{output_dir}/GMI_Analysis_{timestamp}.xlsx"
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Export individual project data to separate sheets
        for project_name, project_data in all_projects_data.items():
            # Create detailed data sheet
            df = pd.DataFrame([
                {
                    'Year': year,
                    'Producers': data['producers'],
                    'Fields': data['fields'],
                    'Total_Acres': data['total_acres'],
                    'Cover_Crop_Acres': data['cover_crop_acres'],
                    'Cover_Crop_Percentage': data['cover_crop_acres'] / data['total_acres'] * 100 if data['total_acres'] > 0 else 0,
                    'Fields_with_Cover_Crops': data['fields_with_cover_crops'],
                    'Fields_without_Cover_Crops': data['fields_without_cover_crops'],
                    'Tillage_Events': data['tillage_events'],
                    'Fertilizer_Applications': data['fertilizer_applications'],
                    'Pesticide_Applications': data['pesticide_applications']
                }
                for year, data in project_data['years'].items()
            ])
            df.to_excel(writer, sheet_name=f"{project_name[:30]}", index=False)
            
            # Create project summary sheet
            summary_data = []
            for year in ANALYSIS_YEARS:
                if year in project_data['years']:
                    data = project_data['years'][year]
                    if data['fields'] > 0:  # Only include years with data
                        summary_data.append({
                            'Year': year,
                            'Total_Acres': data['total_acres'],
                            'Cover_Crop_Acres': data['cover_crop_acres'],
                            'Cover_Crop_Percentage': data['cover_crop_acres'] / data['total_acres'] * 100 if data['total_acres'] > 0 else 0,
                            'Fields_with_Cover_Crops': data['fields_with_cover_crops'],
                            'Fields_without_Cover_Crops': data['fields_without_cover_crops'],
                            'Tillage_Events_per_Field': data['tillage_events'] / data['fields'] if data['fields'] > 0 else 0,
                            'Fertilizer_Apps_per_Field': data['fertilizer_applications'] / data['fields'] if data['fields'] > 0 else 0,
                            'Pesticide_Apps_per_Field': data['pesticide_applications'] / data['fields'] if data['fields'] > 0 else 0,
                            'Tillage_Events_per_Acre': data['tillage_events'] / data['total_acres'] if data['total_acres'] > 0 else 0,
                            'Fertilizer_Apps_per_Acre': data['fertilizer_applications'] / data['total_acres'] if data['total_acres'] > 0 else 0,
                            'Pesticide_Apps_per_Acre': data['pesticide_applications'] / data['total_acres'] if data['total_acres'] > 0 else 0
                        })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name=f"{project_name[:30]}_Summary", index=False)
        
        # Export combined data
        combined_df = pd.DataFrame([
            {
                'Year': year,
                'Producers': data['producers'],
                'Fields': data['fields'],
                'Total_Acres': data['total_acres'],
                'Cover_Crop_Acres': data['cover_crop_acres'],
                'Cover_Crop_Percentage': data['cover_crop_acres'] / data['total_acres'] * 100 if data['total_acres'] > 0 else 0,
                'Fields_with_Cover_Crops': data['fields_with_cover_crops'],
                'Fields_without_Cover_Crops': data['fields_without_cover_crops'],
                'Tillage_Events': data['tillage_events'],
                'Fertilizer_Applications': data['fertilizer_applications'],
                'Pesticide_Applications': data['pesticide_applications']
            }
            for year, data in combined_data['years'].items()
        ])
        combined_df.to_excel(writer, sheet_name='All_Projects_Combined', index=False)
        
        # Create comprehensive summary sheet
        summary_rows = []
        for project_name, project_data in all_projects_data.items():
            # Get the most recent year's data (2024)
            latest_data = project_data['years'][2024]
            summary_rows.append({
                'Project': project_name,
                'Total_Producers': latest_data['producers'],
                'Total_Fields': latest_data['fields'],
                'Total_Acres': latest_data['total_acres'],
                'Cover_Crop_Acres': latest_data['cover_crop_acres'],
                'Cover_Crop_Percentage': latest_data['cover_crop_acres'] / latest_data['total_acres'] * 100 if latest_data['total_acres'] > 0 else 0,
                'Fields_with_Cover_Crops': latest_data['fields_with_cover_crops'],
                'Fields_without_Cover_Crops': latest_data['fields_without_cover_crops'],
                'Tillage_Events_per_Field': latest_data['tillage_events'] / latest_data['fields'] if latest_data['fields'] > 0 else 0,
                'Fertilizer_Apps_per_Field': latest_data['fertilizer_applications'] / latest_data['fields'] if latest_data['fields'] > 0 else 0,
                'Pesticide_Apps_per_Field': latest_data['pesticide_applications'] / latest_data['fields'] if latest_data['fields'] > 0 else 0,
                'Tillage_Events_per_Acre': latest_data['tillage_events'] / latest_data['total_acres'] if latest_data['total_acres'] > 0 else 0,
                'Fertilizer_Apps_per_Acre': latest_data['fertilizer_applications'] / latest_data['total_acres'] if latest_data['total_acres'] > 0 else 0,
                'Pesticide_Apps_per_Acre': latest_data['pesticide_applications'] / latest_data['total_acres'] if latest_data['total_acres'] > 0 else 0
            })
        
        # Add combined data to summary
        latest_combined = combined_data['years'][2024]
        summary_rows.append({
            'Project': 'All Projects Combined',
            'Total_Producers': latest_combined['producers'],
            'Total_Fields': latest_combined['fields'],
            'Total_Acres': latest_combined['total_acres'],
            'Cover_Crop_Acres': latest_combined['cover_crop_acres'],
            'Cover_Crop_Percentage': latest_combined['cover_crop_acres'] / latest_combined['total_acres'] * 100 if latest_combined['total_acres'] > 0 else 0,
            'Fields_with_Cover_Crops': latest_combined['fields_with_cover_crops'],
            'Fields_without_Cover_Crops': latest_combined['fields_without_cover_crops'],
            'Tillage_Events_per_Field': latest_combined['tillage_events'] / latest_combined['fields'] if latest_combined['fields'] > 0 else 0,
            'Fertilizer_Apps_per_Field': latest_combined['fertilizer_applications'] / latest_combined['fields'] if latest_combined['fields'] > 0 else 0,
            'Pesticide_Apps_per_Field': latest_combined['pesticide_applications'] / latest_combined['fields'] if latest_combined['fields'] > 0 else 0,
            'Tillage_Events_per_Acre': latest_combined['tillage_events'] / latest_combined['total_acres'] if latest_combined['total_acres'] > 0 else 0,
            'Fertilizer_Apps_per_Acre': latest_combined['fertilizer_applications'] / latest_combined['total_acres'] if latest_combined['total_acres'] > 0 else 0,
            'Pesticide_Apps_per_Acre': latest_combined['pesticide_applications'] / latest_combined['total_acres'] if latest_combined['total_acres'] > 0 else 0
        })
        
        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name='Project_Summary', index=False)
        
        # Create trend analysis sheet
        trend_rows = []
        for project_name, project_data in all_projects_data.items():
            # Calculate year-over-year changes
            for year in range(2017, 2025):  # Start from 2017 to have previous year for comparison
                if year in project_data['years'] and year-1 in project_data['years']:
                    current = project_data['years'][year]
                    previous = project_data['years'][year-1]
                    
                    # Calculate percentage changes
                    trend_rows.append({
                        'Project': project_name,
                        'Year': year,
                        'Cover_Crop_Acres_Change': ((current['cover_crop_acres'] - previous['cover_crop_acres']) / previous['cover_crop_acres'] * 100) if previous['cover_crop_acres'] > 0 else 0,
                        'Cover_Crop_Fields_Change': ((current['fields_with_cover_crops'] - previous['fields_with_cover_crops']) / previous['fields_with_cover_crops'] * 100) if previous['fields_with_cover_crops'] > 0 else 0,
                        'Tillage_Events_Change': ((current['tillage_events'] - previous['tillage_events']) / previous['tillage_events'] * 100) if previous['tillage_events'] > 0 else 0,
                        'Fertilizer_Apps_Change': ((current['fertilizer_applications'] - previous['fertilizer_applications']) / previous['fertilizer_applications'] * 100) if previous['fertilizer_applications'] > 0 else 0,
                        'Pesticide_Apps_Change': ((current['pesticide_applications'] - previous['pesticide_applications']) / previous['pesticide_applications'] * 100) if previous['pesticide_applications'] > 0 else 0
                    })
        
        trend_df = pd.DataFrame(trend_rows)
        trend_df.to_excel(writer, sheet_name='Trend_Analysis', index=False)
    
    log_progress(f"Exported analysis to Excel file: {excel_filename}")

def main():
    """Main function to analyze all GMI projects."""
    log_progress("Starting GMI project analysis")
    all_projects_data = {}
    
    # Analyze each project
    for project_name, project_id in GMI_PROJECTS.items():
        log_progress(f"\nAnalyzing {project_name}...")
        start_time = time.time()
        
        project_data = analyze_project_data(project_id)
        all_projects_data[project_name] = project_data
        
        # Generate and print project report
        report = generate_project_report(project_data, project_name)
        print(report)
        
        # Calculate and print trends for individual project
        trends = calculate_trends(project_data)
        trend_report = generate_trend_report(trends, project_name)
        print(trend_report)
        
        end_time = time.time()
        log_progress(f"Completed {project_name} analysis in {end_time - start_time:.1f} seconds")
    
    # Analyze combined data
    log_progress("\nAnalyzing combined data across all projects...")
    combined_data = analyze_combined_data(all_projects_data)
    combined_report = generate_project_report(combined_data, "All GMI Projects Combined")
    print(combined_report)
    
    # Calculate and print combined trends
    combined_trends = calculate_trends(combined_data)
    combined_trend_report = generate_trend_report(combined_trends, "All GMI Projects Combined")
    print(combined_trend_report)
    
    # Export to Excel
    export_to_csv(all_projects_data, combined_data)
    
    log_progress("Analysis complete!")

if __name__ == "__main__":
    main() 
