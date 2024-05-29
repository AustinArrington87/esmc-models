### 1 - Way TUKEY'S HSD - Austin Arrington (c)

# import pandas as pd
# from scipy import stats
# from statsmodels.stats.multicomp import pairwise_tukeyhsd
# import itertools

# # Load the dataset
# file_path_correct = 'ESMC_crop_c3_c4.csv'  # Replace this with the path to your dataset
# data = pd.read_csv(file_path_correct)

# # Rename columns to remove spaces and special characters
# filtered_data = data.rename(columns={
#     'field_state': 'field_state',
#     'crop type': 'crop_type',
#     #'soc stock (tonnes/ha)': 'crop_yield_tonnes_ha'
#     'yield (bushels/acre)': 'crop_yield'
# })

# # Drop rows with missing values in the relevant columns
# #data_correct = data.dropna(subset=['soc stock (tonnes/ha)', 'crop type', 'field_state'])
# data_correct = filtered_data.dropna(subset=['crop_yield', 'crop_type', 'field_state'])

# # Filter out the rows where "Commodity" is sunflowers, peas, ryes, beans, cotton, barley, or oats
# excluded_commodities = ["sunflowers", "peas", "ryes", "beans", "cotton", "barley", "oats"]
# filtered_data = data_correct[~data_correct['Commodity'].isin(excluded_commodities)]

# # Ensure 'crop type' is categorical
# filtered_data['crop_type'] = filtered_data['crop_type'].astype('category')

# # Convert 'soc stock (tonnes/ha)' to numeric, setting errors as NaN
# #filtered_data['soc stock (tonnes/ha)'] = pd.to_numeric(filtered_data['soc stock (tonnes/ha)'], errors='coerce')
# filtered_data['crop_yield'] = pd.to_numeric(filtered_data['crop_yield'], errors='coerce')

# # Drop rows with NaN values in 'soc stock (tonnes/ha)'
# #filtered_data = filtered_data.dropna(subset=['soc stock (tonnes/ha)'])
# filtered_data = filtered_data.dropna(subset=['crop_yield'])

# # One-way ANOVA
# #anova_results = stats.f_oneway(*[filtered_data[filtered_data['crop_type'] == crop]['soc stock (tonnes/ha)'] for crop in filtered_data['crop_type'].unique()])
# anova_results = stats.f_oneway(*[filtered_data[filtered_data['crop_type'] == crop]['crop_yield'] for crop in filtered_data['crop_type'].unique()])

# # Display the ANOVA test results
# print("ANOVA results:")
# print(f"F-statistic: {anova_results.statistic}")
# print(f"p-value: {anova_results.pvalue}")

# # Tukey's HSD for Post Hoc analysis
# tukey_results = pairwise_tukeyhsd(endog=filtered_data['crop_yield'],
#                                   groups=filtered_data['crop_type'],
#                                   alpha=0.05)

# # Display the Tukey's HSD test results
# print(tukey_results)

# # Convert the Tukey's HSD results to a DataFrame
# tukey_results_df = pd.DataFrame(data=tukey_results._results_table.data[1:], columns=tukey_results._results_table.data[0])

# # Add Tukey's letters
# def get_tukey_letters(groups, pvals, alpha=0.05):
#     group_dict = {group: set() for group in groups}
#     for (i, group1), (j, group2) in itertools.combinations(enumerate(groups), 2):
#         if pvals[i, j] > alpha:
#             group_dict[group1].add(group2)
#             group_dict[group2].add(group1)
#     group_list = sorted(group_dict.items(), key=lambda x: x[0])
#     letters = []
#     current_letter = 'A'
#     used_letters = {}
#     for group, group_set in group_list:
#         for letter, letter_set in used_letters.items():
#             if not letter_set.intersection(group_set):
#                 used_letters[letter].add(group)
#                 letters.append((group, letter))
#                 break
#         else:
#             used_letters[current_letter] = {group}
#             letters.append((group, current_letter))
#             current_letter = chr(ord(current_letter) + 1)
#     return letters

# groups = tukey_results.groupsunique
# pvals = tukey_results._results_table.data[1:]

# # Extract group names and p-values
# group_names = [row[0] for row in pvals]
# group_pvals = {(row[0], row[1]): row[4] for row in pvals}

# # Create a square matrix of p-values for get_tukey_letters function
# pval_matrix = pd.DataFrame(index=groups, columns=groups, data=1.0)
# for (group1, group2), pval in group_pvals.items():
#     pval_matrix.at[group1, group2] = pval
#     pval_matrix.at[group2, group1] = pval

# letters = get_tukey_letters(groups, pval_matrix.values)

# # Create a DataFrame for letters
# letters_df = pd.DataFrame(letters, columns=['group', 'tukey_letter'])

# # Merge Tukey's HSD results with letters
# tukey_results_with_letters_df = tukey_results_df.merge(letters_df, left_on='group1', right_on='group', how='left').drop(columns='group')

# # Export the results to a CSV file
# output_file_path = 'results/ANOVAs/Tukeys_letters/tukey_hsd_results_1way_crop_Yield_with_letters.csv'  # Corrected file path for export
# tukey_results_with_letters_df.to_csv(output_file_path, index=False)


### 2 - Way TUKEY'S HSD - Austin Arrington (c)

import pandas as pd
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import itertools

# Load the dataset
file_path_correct = 'ESMC_crop_c3_c4.csv'  # Replace this with the path to your dataset
data = pd.read_csv(file_path_correct)

# Rename columns to remove spaces and special characters
filtered_data = data.rename(columns={
    'field_state': 'field_state',
    'crop type': 'crop_type',
    #'soc stock (tonnes/ha)': 'soc_stock_tonnes_ha'
    'yield (bushels/acre)': 'crop_yield'
})

# Drop rows with missing values in the relevant columns
data_correct = filtered_data.dropna(subset=['crop_yield', 'crop_type', 'field_state'])

# Filter out the rows where "Commodity" is sunflowers, peas, ryes, beans, cotton, barley, or oats
excluded_commodities = ["sunflowers", "peas", "ryes", "beans", "cotton", "barley", "oats"]
filtered_data = data_correct[~data_correct['Commodity'].isin(excluded_commodities)]

# Ensure 'field_state' and 'crop type' are categorical
filtered_data['field_state'] = filtered_data['field_state'].astype('category')
filtered_data['crop_type'] = filtered_data['crop_type'].astype('category')

# Add a new column for interaction groups
filtered_data['interaction_group'] = filtered_data['field_state'].astype(str) + "_" + filtered_data['crop_type'].astype(str)

# Convert 'soc stock (tonnes/ha)' to numeric, setting errors as NaN
#filtered_data['soc stock (tonnes/ha)'] = pd.to_numeric(filtered_data['soc stock (tonnes/ha)'], errors='coerce')
filtered_data['crop_yield'] = pd.to_numeric(filtered_data['crop_yield'], errors='coerce')

# Drop rows with NaN values in 'soc stock (tonnes/ha)'
#filtered_data = filtered_data.dropna(subset=['soc stock (tonnes/ha)'])
#filtered_data = filtered_data.dropna(subset=['soil_clay_fraction'])

# Tukey's HSD for Post Hoc analysis
tukey_results = pairwise_tukeyhsd(endog=filtered_data['crop_yield'],
                                  groups=filtered_data['field_state'],
                                  alpha=0.05)

# Display the Tukey's HSD test results
print(tukey_results)

# Convert the Tukey's HSD results to a DataFrame
tukey_results_df = pd.DataFrame(data=tukey_results._results_table.data[1:], columns=tukey_results._results_table.data[0])

# Add Tukey's letters
def get_tukey_letters(groups, pvals, alpha=0.05):
    group_dict = {group: set() for group in groups}
    for (i, group1), (j, group2) in itertools.combinations(enumerate(groups), 2):
        if pvals[i, j] > alpha:
            group_dict[group1].add(group2)
            group_dict[group2].add(group1)
    group_list = sorted(group_dict.items(), key=lambda x: x[0])
    letters = []
    current_letter = 'A'
    used_letters = {}
    for group, group_set in group_list:
        for letter, letter_set in used_letters.items():
            if not letter_set.intersection(group_set):
                used_letters[letter].add(group)
                letters.append((group, letter))
                break
        else:
            used_letters[current_letter] = {group}
            letters.append((group, current_letter))
            current_letter = chr(ord(current_letter) + 1)
    return letters

groups = tukey_results.groupsunique
pvals = tukey_results._results_table.data[1:]

# Extract group names and p-values
group_names = [row[0] for row in pvals]
group_pvals = {(row[0], row[1]): row[4] for row in pvals}

# Create a square matrix of p-values for get_tukey_letters function
pval_matrix = pd.DataFrame(index=groups, columns=groups, data=1.0)
for (group1, group2), pval in group_pvals.items():
    pval_matrix.at[group1, group2] = pval
    pval_matrix.at[group2, group1] = pval

letters = get_tukey_letters(groups, pval_matrix.values)

# Create a DataFrame for letters
letters_df = pd.DataFrame(letters, columns=['group', 'tukey_letter'])

# Merge Tukey's HSD results with letters
tukey_results_with_letters_df = tukey_results_df.merge(letters_df, left_on='group1', right_on='group', how='left').drop(columns='group')

# Export the results to a CSV file
output_file_path = 'results/ANOVAs/Tukeys_letters/2Way/tukey_hsd_results_2way_Yield_State_with_letters.csv'  # Corrected file path for export
tukey_results_with_letters_df.to_csv(output_file_path, index=False)
