import pandas as pd
import re

# Load the CSV file
file_path = '2023_Verified.csv'
df = pd.read_csv(file_path)

# Define the crop_yields dictionary with conversion factors (kg per unit)
crop_yields = {
    "Wheat, Winter": {"unit": "bushels", "kg_per_unit": 27.21},
    "Soybeans": {"unit": "bushels", "kg_per_unit": 27.21},
    "Corn, Grain": {"unit": "bushels", "kg_per_unit": 25.4},
    "Sunflowers": {"unit": "pounds", "kg_per_unit": 0.453592},
    "Oats": {"unit": "bushels", "kg_per_unit": 14.51},
    "Sorghum, Grain": {"unit": "bushels", "kg_per_unit": 25.4},
    "Cotton": {"unit": "lbs", "kg_per_unit": 0.453592},
    "Pea, Chinese Red/Cowpea": {"unit": "bushels", "kg_per_unit": 27.21},
    "Wheat, Spring": {"unit": "bushels", "kg_per_unit": 27.21},
    "Barley": {"unit": "bushels", "kg_per_unit": 21.77},
    "Rye": {"unit": "bushels", "kg_per_unit": 25.4},
    "Beans, Dry": {"unit": "bushels", "kg_per_unit": 27.21}
}

# Function to extract the crop name from the crop_yield column
def extract_crop_name(crop_yield_str):
    try:
        # Extract the crop name (e.g., "Soybeans" from "Crop- (Soybeans: Yield- 36.90 bushels/acre Type-normal)")
        crop_name = re.search(r'Crop-\s*\(([^:]+):', crop_yield_str).group(1)
        return crop_name.strip()
    except (AttributeError, ValueError):
        return None

# Extract numeric yield value and the unit (bushels or pounds per acre)
def extract_yield_and_unit(crop_yield_str):
    try:
        # Extract the yield number
        yield_value = float(re.search(r'Yield-\s*([\d.]+)', crop_yield_str).group(1))
        
        # Extract the unit (bushels or pounds)
        unit_match = re.search(r'(bushels|pounds)/acre', crop_yield_str)
        if unit_match:
            unit = unit_match.group(1)
        else:
            unit = None
            
        return yield_value, unit
    except (AttributeError, ValueError):
        return None, None

# Extract yield and unit from the crop_yield column
df['crop_yield_extracted'] = df['crop_yield'].apply(lambda x: extract_yield_and_unit(str(x)))

# Calculate crop weight based on extracted yield and unit
def calculate_crop_weight(row, crop_yield_data):
    crop_name = extract_crop_name(row["crop_yield"])
    yield_value, unit = row["crop_yield_extracted"]
    
    if crop_name in crop_yield_data and yield_value is not None and unit == crop_yield_data[crop_name]["unit"]:
        # Get the kg per unit based on the crop
        kg_per_unit = crop_yield_data[crop_name]["kg_per_unit"]
        total_crop_weight = yield_value * kg_per_unit  # Yield * kg per unit
        return total_crop_weight
    return None

# Calculate crop_weight
df["crop_weight"] = df.apply(calculate_crop_weight, axis=1, crop_yield_data=crop_yields)

# Define function to calculate total emissions and ci_score
def calculate_ci_score(row):
    # Convert relevant emissions columns to kg (multiply by 0.001 for tonne-to-kg conversion)
    total_emissions = (
        (row["n2o_direct"] * 0.001) +
        (row["n2o_indirect_adjusted"] * 0.001) +
        (row["methane"] * 0.001) +
        row["field_practice_emissions"]
    )
    
    if row["crop_weight"] and row["crop_weight"] != 0:
        return total_emissions / row["crop_weight"]
    return None

# Calculate ci_score based on emissions and crop weight
df["ci_score"] = df.apply(calculate_ci_score, axis=1)

# Save the updated CSV file with the new columns
output_file_path = "2023_Verified_with_All_Crops_Corrected.csv"
df.to_csv(output_file_path, index=False)

print(f"File saved at: {output_file_path}")
