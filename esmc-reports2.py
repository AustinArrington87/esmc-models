import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_UNDERLINE
from docx.shared import RGBColor
from docx.oxml.ns import qn
import os

# Load the data
year = "2021"
file_path = year+'_Verified.csv'
data = pd.read_csv(file_path)

# Create the 'Producers' directory if it doesn't exist
output_dir = "Producers"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def format_cell_data(data):
    # Format data to 3 decimal places if it is a number
    try:
        # Check if the data is convertible to a float
        float_data = float(data)
        # Return formatted data
        return f"{float_data:.3f}"
    except ValueError:
        # Return original data if it's not a number
        return data


# Function to create and format a table in the Word document
def create_table(document, data, column_names, title=None, note=None):
    if title:
        document.add_heading(title, level=1)
    table = document.add_table(rows=1, cols=len(column_names))
    for i, column in enumerate(column_names):
        table.cell(0, i).text = column
        table.cell(0, i).paragraphs[0].runs[0].font.bold = True  # Make header bold
    for row_data in data:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            row_cells[i].text = format_cell_data(cell_data)
    if note:
        para = document.add_paragraph()
        run = para.add_run(note)
        run.italic = True

# Calculate Delta for N2O, Methane, Fossil Fuel, Pesticides, and Upstream Emissions
def calculate_delta(group, baseline_column, practice_column):
    return group[baseline_column] - group[practice_column]

# custom font for bulleted sections 
def add_custom_bullet(doc, label, value):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(label)
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(12)  # Setting the font size to 11, modify as needed

    value_run = p.add_run(value)
    value_run.font.name = 'Calibri'
    value_run.font.size = Pt(12)

# Hyperlink 

def add_hyperlink(paragraph, text, url):
    """
    A function that places a hyperlink within a paragraph object.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param text: The text to be used for the link.
    :param url: The URL the link points to.
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    # Create the w:hyperlink tag and set needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id, )
    hyperlink.set(qn('w:history'), '1')  # optional

    # Create a w:r element
    new_run = OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = OxmlElement('w:rPr')

    # Add color if needed
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0000FF')  # sets the color to blue
    rPr.append(c)

    # Add underline if needed
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink

######### MAIN FOR LOOP to Generate Reports ##########
# Generate a report for each producer
for producer, group in data.groupby('Producer (Project)'):
    doc = Document()

    # Access the 'Project' column of the first row in the group
    project = group['Project'].iloc[0]

    # Access the header of the document
    section = doc.sections[0]

    header = section.header
    # Add a paragraph in the header and then add a run
    header_para = header.paragraphs[0]
    run = header_para.add_run()
    # Add your PNG image to the run (specify the path to your image)
    run.add_picture('header.png', width=Inches(11))
    doc.add_paragraph() # empty space 

    # producer name 
    #doc.add_heading(producer, 0)

    para = doc.add_heading()
    run = para.add_run(f"Eco-Harvest Carbon Impact Report")
    run_font = run.font
    run_font.name = 'Calibri'
    run_font.size = Pt(24)
    run_font.underline = WD_UNDERLINE.SINGLE
    run_font.color.rgb = RGBColor(0, 0, 0)

    # summary info at top 
    summary_para1 = doc.add_paragraph()
    # Add a run for "Producer:" and make it bold
    run_bold = summary_para1.add_run('Producer: ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Producer-1" without bold
    run_regular = summary_para1.add_run(producer)
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    # Add a run for "Project:" and make it bold
    summary_para2 = doc.add_paragraph()
    run_bold = summary_para2.add_run('Project: ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Project" without bold
    run_regular = summary_para2.add_run(project)
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    # Add a run for "Year" and make it bold
    summary_para3 = doc.add_paragraph()
    run_bold = summary_para3.add_run('Year: ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Project" without bold
    run_regular = summary_para3.add_run(year)
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    # YOUR RESULTS 
    para = doc.add_heading()
    run = para.add_run(f"Your Results")
    run_font = run.font
    run_font.name = 'Calibri'
    run_font.size = Pt(16)
    run_font.underline = WD_UNDERLINE.SINGLE
    run_font.color.rgb = RGBColor(0, 0, 0)

    # Fields Summary section load 
    fields_data = group[['Field Name (MRV)', 'Acres', 'Commodity']].copy()  # Replace 'Commodity'
    fields_data.rename(columns={'Field Name (MRV)': 'Field'}, inplace=True)
    # Reductions section load 
    reductions_data = group[['Field Name (MRV)', 'reduced']].copy()
    reductions_data.rename(columns={'Field Name (MRV)': 'Field'}, inplace=True)
    total_reductions = reductions_data['reduced'].sum()
    # Removals section load 
    removals_data = group[['Field Name (MRV)', 'baseline_dsoc', 'dsoc', 'removed']].copy()
    removals_data.rename(columns={'Field Name (MRV)': 'Field', 'removed': 'Removed'}, inplace=True)
    total_removals = removals_data['Removed'].sum()


    # Create a bullet point for Total Reductions and apply styles
    add_custom_bullet(doc, "Total Reductions: ", f"{total_reductions:.3f} tonnes CO2e")

    # Create a bullet point for Total Removals and apply styles
    add_custom_bullet(doc, "Total Removals: ", f"{total_removals:.3f} tonnes CO2e")

    # Create a bullet point for Total Acres and apply styles
    add_custom_bullet(doc, "Total Acres: ", f"{fields_data['Acres'].sum():.3f}")


    # Eco-Harvest Program Info 
    para = doc.add_heading()
    run = para.add_run(f"Eco-Harvest: rewarding producers for regenerative ag ")
    run_font = run.font
    run_font.name = 'Calibri'
    run_font.size = Pt(16)
    run_font.underline = WD_UNDERLINE.SINGLE
    run_font.color.rgb = RGBColor(0, 0, 0)

    add_custom_bullet(doc, "Customizable practices: ", f"Choose how to implement regenerative practices.")
    add_custom_bullet(doc, "Fair Payments: ", f"As a non-profit, ESMC maximizes producer benefits.")
    add_custom_bullet(doc, "Return on Investment: ", f"Many farmers experience a 15-25% return on investment after 3–5-years of regenerative farming.")
    add_custom_bullet(doc, "Weather-Ready Farming: ", f"Thank you for participating in the Eco-Harvest program. Continue working with your advisor to improve your soil health and boost your farm’s productivity and earnings.")

    # Thank You 
    thanks_para1 = doc.add_paragraph()
    # Add another run for "Producer-1" without bold
    run_regular = thanks_para1.add_run("Thank you for participating in the Eco-Harvest program. Continue working with your advisor to improve your soil health and boost your farm’s productivity and earnings.")
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    # Contact Us 
     # summary info at top 
    summary_para1 = doc.add_paragraph()
    # Add a run for "Producer:" and make it bold
    run_bold = summary_para1.add_run('Need Assistance? ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Producer-1" without bold
    run_regular = summary_para1.add_run("Contact your Project Representative or ")
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)
    add_hyperlink(summary_para1, 'support@ecoharvest.ag', 'http://www.ecoharvest.ag')

    ############## SECTION 2 ##################################
     # Appendix 
    para = doc.add_heading()
    run = para.add_run(f"Appendix")
    run_font = run.font
    run_font.name = 'Calibri'
    run_font.size = Pt(16)
    run_font.underline = WD_UNDERLINE.SINGLE
    run_font.color.rgb = RGBColor(0, 0, 0)

    para = doc.add_heading()
    run = para.add_run(f"Carbon Impact Explained")
    run_font = run.font
    run_font.name = 'Calibri'
    run_font.size = Pt(12)
    run_font.underline = WD_UNDERLINE.SINGLE
    run_font.color.rgb = RGBColor(0, 0, 0)

    summary_para1 = doc.add_paragraph()
    run_regular = summary_para1.add_run("ESMC uses historical data, a scientific model, soil samples and weather data to calculate emissions reductions and soil carbon removals. Any negative emissions are increased emissions, which is not unusual in the early years of adopting a practice change. If concerned, please speak with your technical advisor about further modifications to your practice change(s). ")

    summary_para2 = doc.add_paragraph()
    # Add a run for "Producer:" and make it bold
    run_bold = summary_para2.add_run('Greenhouse Gas Emissions (GHG) Reductions (mtCO2e) ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Producer-1" without bold
    run_regular = summary_para2.add_run("are the sum of the reductions of methane (CH4), direct nitrous oxide (N2O), indirect nitrous oxide (N2O) and supply chain and operational emissions between the baseline year and the project year.")
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    summary_para3 = doc.add_paragraph()
    # Add a run for "Producer:" and make it bold
    run_bold = summary_para3.add_run('Soil Carbon Removals (mtCO2e): ')
    run_bold.font.bold = True
    run_bold.font.name = 'Calibri'
    run_bold.font.size = Pt(12)
    run_bold.font.color.rgb = RGBColor(0, 0, 0)

    # Add another run for "Producer-1" without bold
    run_regular = summary_para3.add_run("The difference in the amount of carbon in the soil between the baseline year and the project year.")
    run_regular.font.bold = False  # Not necessary as False is the default
    run_regular.font.name = 'Calibri'
    run_regular.font.size = Pt(12)
    run_regular.font.color.rgb = RGBColor(0, 0, 0)

    # IMPACT Breakdown Table 
    

    ############## SECTION 3 ##################################

    # # Fields section
    # create_table(doc, fields_data.values, ['Field', 'Acres', 'Commodity'], title='Fields')
    # doc.add_paragraph()
    # para = doc.add_paragraph()
    # run = para.add_run(f"Total Acres: {fields_data['Acres'].sum():.2f}")
    # run.italic = True

    # # Practices section
    # practices_data = group[['Field Name (MRV)', 'Practice Change']].copy()
    # practices_data.rename(columns={'Field Name (MRV)': 'Field'}, inplace=True)
    # create_table(doc, practices_data.values, ['Field', 'Practice Change'], title='Practices')

    # # Soil Data section
    # soil_data = group[['Field Name (MRV)', 'soil_avg_soc', 'soil_avg_ph', 'soil_avg_bulkdensity', 'soil_clay_fraction']].copy()
    # soil_data.rename(columns={'Field Name (MRV)': 'Field', 'soil_avg_soc': 'SOC', 'soil_avg_ph': 'pH', 'soil_avg_bulkdensity': 'BD', 'soil_clay_fraction': 'Clay'}, inplace=True)
    # create_table(doc, soil_data.values, ['Field', 'SOC', 'pH', 'BD', 'Clay'], title='Soil Data')

    # Reductions section
    # doc.add_heading('Reductions', level=1)
    # doc.add_paragraph("This section shows the GHG emissions reductions from the intervention, and includes DNDC modeled reductions as well as reductions from field activity changes (calculated through ecoinvent3.8 emission factors).")
    
    # # Example for Total Reductions
    # create_table(doc, reductions_data.values, ['Field', 'Reduced'], title='Total Reductions')
    # doc.add_paragraph()
    # para = doc.add_paragraph()
    # run = para.add_run(f"Total Reductions: {total_reductions:.3f} tonnes CO2e")
    # run.italic = True

    # Similar approach for N2O, Methane, etc.
    
    # Direct N2O
    group['Delta Direct N2O'] = calculate_delta(group, 'baseline_n20_direct', 'n2o_direct')
    # N2O Direct section
    n2o_data = group[['Field Name (MRV)', 'baseline_n20_direct', 'n2o_direct', 'Delta Direct N2O']].copy()
    n2o_data.rename(columns={
        'Field Name (MRV)': 'Field',
        'baseline_n20_direct': 'Direct N2O Baseline',
        'n2o_direct': 'Direct N2O Practice',
        'Delta Direct N2O': 'Delta'
    }, inplace=True)
    # create_table(doc, n2o_data.values, ['Field', 'Direct N2O Baseline', 'Direct N2O Practice', 'Delta'], title='N2O Direct Emissions')
    # doc.add_paragraph()
    # Summarize the total Delta Direct N2O
    total_delta_direct_n2o = n2o_data['Delta'].sum()
    # para = doc.add_paragraph()
    # run = para.add_run(f"Total Direct N2O Reductions: {total_delta_direct_n2o:.3f} tonnes CO2e")
    # run.italic = True

    #Indirect N2O
    group['Delta Indirect N2O'] = calculate_delta(group, 'baseline_n2o_indirect', 'n2o_indirect')
    # N2O Direct section
    n2oindirect_data = group[['Field Name (MRV)', 'baseline_n2o_indirect', 'n2o_indirect', 'Delta Indirect N2O']].copy()
    n2oindirect_data.rename(columns={
        'Field Name (MRV)': 'Field',
        'baseline_n2o_indirect': 'Indirect N2O Baseline',
        'n2o_indirect': 'Indirect N2O Practice',
        'Delta Indirect N2O': 'Delta'
    }, inplace=True)
    # create_table(doc, n2oindirect_data.values, ['Field', 'Indirect N2O Baseline', 'Indirect N2O Practice', 'Delta'], title='N2O Indirect Emissions')
    # doc.add_paragraph()
    # Summarize the total Delta Indirect N2O
    total_delta_indirect_n2o = n2oindirect_data['Delta'].sum()
    # para = doc.add_paragraph()
    # run = para.add_run(f"Total Indirect N2O Reductions: {total_delta_indirect_n2o:.3f} tonnes CO2e")
    # run.italic = True

    # Methane 
    group['Delta Methane'] = calculate_delta(group, 'baseline_methane', 'methane')
    # Methane section
    methane_data = group[['Field Name (MRV)', 'baseline_methane', 'methane', 'Delta Methane']].copy()
    methane_data.rename(columns={
        'Field Name (MRV)': 'Field',
        'baseline_methane': 'CH4 Baseline',
        'methane': 'CH4 Practice',
        'Delta Methane': 'Delta'
    }, inplace=True)
    # create_table(doc, methane_data.values, ['Field', 'CH4 Baseline', 'CH4 Practice', 'Delta'], title='Methane Emissions')
    # doc.add_paragraph()
    # Summarize the total Methane
    total_delta_methane = methane_data['Delta'].sum()
    # para = doc.add_paragraph()
    # run = para.add_run(f"Total Methane Reductions: {total_delta_methane:.3f} tonnes CO2e")
    # run.italic = True

    #Emission Factors 
    group['Delta Field Emissions'] = calculate_delta(group, 'field_baseline_emissions_demands_fossil', 'field_practice_emissions_demands_fossil')
    # Emissions section
    emissions_data = group[['Field Name (MRV)', 'field_baseline_emissions_demands_fossil', 'field_practice_emissions_demands_fossil', 'Delta Field Emissions']].copy()
    emissions_data.rename(columns={
        'Field Name (MRV)': 'Field',
        'field_baseline_emissions_demands_fossil': 'Emissions Demands Baseline',
        'field_practice_emissions_demands_fossil': 'Emissions Demands Practice',
        'Delta Field Emissions': 'Delta'
    }, inplace=True)
    # create_table(doc, emissions_data.values, ['Field', 'Emissions Demands Baseline', 'Emissions Demands Practice', 'Delta'], title='Emissions Demands')
    # doc.add_paragraph()
    # Summarize the total Emissions
    total_delta_emissions = emissions_data['Delta'].sum()
    # convert kg to CO23
    total_delta_emissions = total_delta_emissions * 0.001
    # para = doc.add_paragraph()
    # doc.add_paragraph("Derived from ecoinvent3.8 cutoff emission factors for upstream and on-field process emissions.")
    # run = para.add_run(f"Total Emissions Demands Reductions: {total_delta_emissions:.3f} kg CO2e")
    # run.italic = True

    # Removals section
    #doc.add_heading('Removals', level=1)
    #doc.add_paragraph("This section shows the carbon removal from the intervention as modeled by the DNDC.")
    #create_table(doc, removals_data.values, ['Field', 'Baseline DSOC', 'DSOC', 'Removed'], title='Total Removals')
    #doc.add_paragraph()
    #para = doc.add_paragraph()
    #run = para.add_run(f"Total Removals: {total_removals:.3f} tonnes CO2e")
    #run.italic = True

    ### IMPACT TABLE 
    table = doc.add_table(rows=8, cols=2)
    table.style = 'Table Grid'
    headers = ['Impact Breakdown', 'DELTA mtCO2e']
    # Insert column headers and bold them
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Define the rows under "Impact Breakdown"
    impact_breakdown = [
        "Greenhouse Gas Emissions Reductions",
        "   N2O Indirect",
        "   N2O Direct",
        "   Methane",
        "   Supply Chain and Operational Emissions",
        "Total Soil Carbon Removals",
        "Total Carbon Impact (Reductions + Removals)"
        ]
    # Values to be bolded in the Impact Breakdown column
    bold_values = [
        "Greenhouse Gas Emissions Reductions",
        "Total Soil Carbon Removals",
        "Total Carbon Impact (Reductions + Removals)"
        ]

    total_carbon_impact = total_reductions + total_removals

    # List of corresponding values for "DELTA mtCO2e"

    delta_values = [
    round(total_reductions, 3),
    round(total_delta_indirect_n2o, 3),
    round(total_delta_direct_n2o, 3),
    round(total_delta_methane, 3),
    round(total_delta_emissions, 3),
    round(total_removals, 3),
    round(total_carbon_impact, 3)
    ]

    # Populate the table rows with data
    for row in range(1, 8):  # Start from 1 to leave the header row as is
        impact_cell = table.cell(row, 0)
        delta_cell = table.cell(row, 1)
        
        impact_cell.text = impact_breakdown[row - 1]
        delta_cell.text = str(delta_values[row - 1])  # Convert numerical values to string

        # Check if the impact breakdown should be bold
        if impact_breakdown[row - 1] in bold_values:
            impact_cell.paragraphs[0].runs[0].font.bold = True

    # Save the document
    safe_producer_name = producer.replace('/', '_').replace('\\', '_')
    doc_file_name = f"{safe_producer_name}.docx"
    doc.save(os.path.join(output_dir, doc_file_name))

print("Reports generated successfully.")
