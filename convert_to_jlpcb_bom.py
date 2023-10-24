import pandas as pd

df = pd.read_csv("DaisySeedBreakout.csv")
df_part_numbers = pd.read_csv("JLPCB_Part_Numbers.csv")


# Merge the two DataFrames on columns A and B
merged = pd.merge(df, df_part_numbers, on=['Comment', 'Footprint'], how='left')

# Fill the 'COPY' column in file1 with values from file2
#merged['COPY_x'] = merged['COPY_y'].fillna(merged['COPY_x'])

# Drop the 'COPY_y' column, as it's no longer needed
#merged = merged.drop(columns=['COPY_y'])

# Save the updated DataFrame back to file1
merged['LCSC Part #'] = merged['LCSC Part #_y']
merged = merged.drop(columns=['LCSC Part #_x', 'LCSC Part #_y'])
merged = merged[merged['LCSC Part #'] != "HAND PLACE"]
merged = merged[merged['LCSC Part #'] != "DNP"]
merged.to_csv('JLPCB_BOM.csv', index=False)
