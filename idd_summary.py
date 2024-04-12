import pandas as pd
from docx import Document

# Load the dataset
file_path = 'IDD_Analytics_2023.csv'  # Update the path if necessary
data = pd.read_csv(file_path)

# Define the practice change categories
categories = ['Nutrient management', 'Normal operation', 'Tillage reduction', 'Cover cropping']

# Create a new Word document
doc = Document()

# Function to process and add a category to the document
def process_category(category):
    # Filter the data for the current category
    filtered_data = data[data['practice_change'] == category]

    # Group the data by program_region_crop and calculate both the number of fields and the summed acres
    grouped_data = filtered_data.groupby('program_region_crop').agg(
        Number_of_Fields=('id', 'nunique'),
        Summed_Acres=('acres', 'sum')
    ).reset_index()

    # Round the acres to two decimal places
    grouped_data['Summed_Acres'] = grouped_data['Summed_Acres'].round(2)

    # Add category as a heading in the Word document
    doc.add_heading(category, level=1)

    # Add the table to the Word document
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'

    # Add header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Program Region Crop'
    hdr_cells[1].text = 'Number of Fields'
    hdr_cells[2].text = 'Summed Acres'

    # Add data rows
    for index, row in grouped_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['program_region_crop'])
        row_cells[1].text = str(row['Number_of_Fields'])
        row_cells[2].text = f"{row['Summed_Acres']:.2f}"  # Format the acres value to 2 decimal places

    # Add a space after the table
    doc.add_paragraph()

# Process each category and add it to the document
for category in categories:
    process_category(category)

# Save the document
doc.save('Agricultural_Practices_Summary.docx')
