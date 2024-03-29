import pandas as pd
from scipy import stats

# Load the dataset
df = pd.read_csv("merged_habiterre_dndc_2022.csv")

# Paired T-Tests
# Replace 'total_reductions_DNDC' and 'total_reductions_ecosys' with the actual column names
t_statistic_1, p_value_1 = stats.ttest_rel(df['total_reductions_DNDC'], df['total_reductions_ecosys'])
print(f"Paired T-Test between 'total_reductions_DNDC' and 'total_reductions_ecosys': t={t_statistic_1}, p={p_value_1}")

# Replace 'reduced' and 'h_delta_dSOC' with the actual column names
t_statistic_2, p_value_2 = stats.ttest_rel(df['reduced'], df['h_delta_dSOC'])
print(f"Paired T-Test between 'reduced' and 'h_delta_dSOC': t={t_statistic_2}, p={p_value_2}")

# Max and Min Values
max_total_reductions_DNDC = df['total_reductions_DNDC'].max()
min_total_reductions_DNDC = df['total_reductions_DNDC'].min()
max_total_reductions_ecosys = df['total_reductions_ecosys'].max()
min_total_reductions_ecosys = df['total_reductions_ecosys'].min()

max_reduced = df['reduced'].max()
min_reduced = df['reduced'].min()
max_h_delta_dSOC = df['h_delta_dSOC'].max()
min_h_delta_dSOC = df['h_delta_dSOC'].min()

print(f"Max and Min for 'total_reductions_DNDC': {max_total_reductions_DNDC}, {min_total_reductions_DNDC}")
print(f"Max and Min for 'total_reductions_ecosys': {max_total_reductions_ecosys}, {min_total_reductions_ecosys}")
print(f"Max and Min for 'reduced': {max_reduced}, {min_reduced}")
print(f"Max and Min for 'h_delta_dSOC': {max_h_delta_dSOC}, {min_h_delta_dSOC}")

# Mean and Standard Deviation
mean_total_reductions_DNDC = df['total_reductions_DNDC'].mean()
std_total_reductions_DNDC = df['total_reductions_DNDC'].std()
mean_total_reductions_ecosys = df['total_reductions_ecosys'].mean()
std_total_reductions_ecosys = df['total_reductions_ecosys'].std()

mean_reduced = df['reduced'].mean()
std_reduced = df['reduced'].std()
mean_h_delta_dSOC = df['h_delta_dSOC'].mean()
std_h_delta_dSOC = df['h_delta_dSOC'].std()

print(f"Mean and Std for 'total_reductions_DNDC': {mean_total_reductions_DNDC}, {std_total_reductions_DNDC}")
print(f"Mean and Std for 'total_reductions_ecosys': {mean_total_reductions_ecosys}, {std_total_reductions_ecosys}")
print(f"Mean and Std for 'reduced': {mean_reduced}, {std_reduced}")
print(f"Mean and Std for 'h_delta_dSOC': {mean_h_delta_dSOC}, {std_h_delta_dSOC}")

# Summary Info
# Adjust 'Program', 'field_state', 'county_name', 'Field ID', and 'Acres' to match your column names
summary_info = df.groupby(['program_region', 'field_state', 'county_name']).agg({
    'Field ID': 'count',
    'Acres': 'sum'
}).rename(columns={'Field ID': 'Number of Fields'})
print(summary_info)
