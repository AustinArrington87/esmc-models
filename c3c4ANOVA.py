# 1 WAY ANOVA 
import pandas as pd
from scipy import stats

# Load the correct CSV file
file_path_correct = 'SMC_crop_c3_c4.csv'
data_correct = pd.read_csv(file_path_correct)

# Filter out the rows where "Commodity" is sunflowers, peas, ryes, beans, cotton, barley, or oats
excluded_commodities = ["sunflowers", "peas", "ryes", "beans", "cotton", "barley", "oats"]
filtered_data = data_correct[~data_correct['Commodity'].isin(excluded_commodities)]

# Rename columns to remove spaces and special characters
filtered_data = filtered_data.rename(columns={
    'field_state': 'field_state',
    'crop type': 'crop_type',
    'soc stock (tonnes/ha)': 'soc_stock_tonnes_ha'
})

# Convert 'soc_stock_tonnes_ha' to numeric, setting errors as NaN
filtered_data['soc_stock_tonnes_ha'] = pd.to_numeric(filtered_data['soc_stock_tonnes_ha'], errors='coerce')

# Drop rows with NaN values in 'soc_stock_tonnes_ha'
filtered_data = filtered_data.dropna(subset=['soc_stock_tonnes_ha'])

# Perform a one-way ANOVA for 'crop_type'
anova_crop_type = stats.f_oneway(filtered_data[filtered_data['crop_type'] == 'c3']['soc_stock_tonnes_ha'],
                                 filtered_data[filtered_data['crop_type'] == 'c4']['soc_stock_tonnes_ha'])

# Display the ANOVA result for 'crop_type'
print(anova_crop_type)

# Save the ANOVA result to a CSV file
anova_result_df = pd.DataFrame({'Source': ['crop_type'], 
                                'F-Statistic': [anova_crop_type.statistic], 
                                'p-Value': [anova_crop_type.pvalue]})
anova_result_df.to_csv('/results/1WAYANOVA_Crop.csv', index=True)

# 2 WAY ANOVA 

import pandas as pd
import statsmodels.api as sm
from statsmodels.formula.api import ols

# Load the correct CSV file
file_path_correct = 'ESMC_crop_c3_c4.csv'
data_correct = pd.read_csv(file_path_correct)

# Filter out the rows where "Commodity" is sunflowers, peas, ryes, beans, cotton, barley, or oats
excluded_commodities = ["sunflowers", "peas", "ryes", "beans", "cotton", "barley", "oats"]
filtered_data = data_correct[~data_correct['Commodity'].isin(excluded_commodities)]

# Rename columns to remove spaces and special characters
filtered_data = filtered_data.rename(columns={
    'field_state': 'field_state',
    'crop type': 'crop_type',
    'soc stock (tonnes/ha)': 'soc_stock_tonnes_ha'
})

# Convert 'soc_stock_tonnes_ha' to numeric, setting errors as NaN
filtered_data['soc_stock_tonnes_ha'] = pd.to_numeric(filtered_data['soc_stock_tonnes_ha'], errors='coerce')

# Drop rows with NaN values in 'soc_stock_tonnes_ha'
filtered_data = filtered_data.dropna(subset=['soc_stock_tonnes_ha'])

# Perform the 2-way ANOVA with the cleaned data
model = ols('soc_stock_tonnes_ha ~ C(field_state) + C(crop_type) + C(field_state):C(crop_type)', data=filtered_data).fit()
anova_table = sm.stats.anova_lm(model, typ=2)

# Display the ANOVA table
print(anova_table)

anova_table.to_csv('results/2WAYANOVA_C3_C4.csv', index=True)
