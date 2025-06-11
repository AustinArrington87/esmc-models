import mrvApi as mrv
import requests
import json
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Any
import time

# Configure API connection
c = mrv.configure(".env.production")

# Your Hasura admin secret key
admin_secret_key = ""

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

def log_progress(message: str, indent: int = 0):
    """Helper function to log progress with timestamp and indentation."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    indent_str = "  " * indent
    print(f"[{timestamp}] {indent_str}{message}")

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
            'cover_crop_acres': 0,
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
                
                for event in events:
                    # Count cover crop events
                    if event.get('coverCropSeedingRate'):
                        year_data['cover_crop_acres'] += field.get('acres', 0)
                    
                    # Count tillage events
                    if event.get('tillageDepth') or event.get('tillageDataId'):
                        year_data['tillage_events'] += 1
                    
                    # Count fertilizer applications
                    if event.get('fertilizerDataId'):
                        year_data['fertilizer_applications'] += 1
                    
                    # Count pesticide applications
                    if event.get('pesticideCount'):
                        year_data['pesticide_applications'] += event['pesticideCount']
                
                log_progress(f"Year {year} summary for field:", indent=4)
                log_progress(f"Cover crop acres: {year_data['cover_crop_acres']:.1f}", indent=5)
                log_progress(f"Tillage events: {year_data['tillage_events']}", indent=5)
                log_progress(f"Fertilizer applications: {year_data['fertilizer_applications']}", indent=5)
                log_progress(f"Pesticide applications: {year_data['pesticide_applications']}", indent=5)
    
    # Print final summary
    log_progress("\nFinal project summary:")
    for year in ANALYSIS_YEARS:
        year_data = project_data['years'][year]
        if year_data['fields'] > 0:  # Only show years with data
            log_progress(f"Year {year}:", indent=1)
            log_progress(f"Fields: {year_data['fields']}", indent=2)
            log_progress(f"Producers: {year_data['producers']}", indent=2)
            log_progress(f"Cover crop acres: {year_data['cover_crop_acres']:.1f}", indent=2)
            log_progress(f"Tillage events: {year_data['tillage_events']}", indent=2)
            log_progress(f"Fertilizer applications: {year_data['fertilizer_applications']}", indent=2)
            log_progress(f"Pesticide applications: {year_data['pesticide_applications']}", indent=2)
    
    return project_data

def generate_project_report(project_data: Dict[str, Any], project_name: str) -> str:
    """Generate a report for a specific project."""
    report = f"\nProject: {project_name}\n"
    report += f"Total Producers: {project_data['total_producers']}\n"
    report += f"Total Fields: {project_data['total_fields']}\n\n"
    
    # Sort years for chronological display
    years = sorted(project_data['years'].keys())
    
    # Create table headers
    report += "Year | Producers | Fields | Cover Crop Acres | Tillage Events | Fertilizer Apps | Pesticide Apps\n"
    report += "-" * 100 + "\n"
    
    # Add data rows
    for year in years:
        data = project_data['years'][year]
        report += f"{year} | {data['producers']} | {data['fields']} | {data['cover_crop_acres']:.1f} | "
        report += f"{data['tillage_events']} | {data['fertilizer_applications']} | {data['pesticide_applications']}\n"
    
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
            'cover_crop_acres': 0,
            'tillage_events': 0,
            'fertilizer_applications': 0,
            'pesticide_applications': 0
        }
    
    # Combine data from all projects
    for project_data in all_projects_data.values():
        for year, year_data in project_data['years'].items():
            combined_data['years'][year]['producers'] += year_data['producers']
            combined_data['years'][year]['fields'] += year_data['fields']
            combined_data['years'][year]['cover_crop_acres'] += year_data['cover_crop_acres']
            combined_data['years'][year]['tillage_events'] += year_data['tillage_events']
            combined_data['years'][year]['fertilizer_applications'] += year_data['fertilizer_applications']
            combined_data['years'][year]['pesticide_applications'] += year_data['pesticide_applications']
        
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
    report = f"\nTrend Analysis for {project_name}\n"
    report += "=" * 50 + "\n\n"
    
    for metric, changes in trends.items():
        if not changes:
            continue
            
        report += f"{metric.replace('_', ' ').title()}:\n"
        report += "-" * 30 + "\n"
        
        for change in changes:
            report += f"Year {change['year']}: {change['change']:+.1f}% "
            report += f"({change['previous']:.1f} → {change['current']:.1f})\n"
        
        report += "\n"
    
    return report

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
    
    log_progress("Analysis complete!")

if __name__ == "__main__":
    main() 
