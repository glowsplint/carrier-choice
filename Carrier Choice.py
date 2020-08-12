import pkg_resources.py2_warn
import pandas as pd
import numpy as np
import os

from pathlib import Path
from datetime import datetime

'''
This script takes in a Freight Matrix file in the same directory as the executable,
    and creates another file f'Carrier Choice - {date_string}.xlsx' which yields the
    top 5 choices by cost (i.e. 1st choice = lowest cost)
'''

# Get BAF Month from user
print('Please specify a 3-letter BAF month and press Enter to submit your selection. (e.g. jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec)')
baf_month = input().lower().strip()

allowed_months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                  'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
if baf_month not in allowed_months:
    error = 'The month you have indicated is not permissible, please check your spelling.'
    print(error)
    input()
    raise Exception(error)

# Read the Freight Matrix
print('Reading Freight Matrix file...this action can take up to 2 minutes.')
files = list(Path(os.getcwd()).glob('Freight Matrix *.xlsx'))
if len(files) == 1:
    df = pd.read_excel(files[0].name, header=1)
else:
    error = 'There are multiple Freight Matrix files that fit the pattern, or no suitable Freight Matrix files! Please include only one in the same directory as this script with the naming pattern of "Freight Matrix *.xlsx".'
    print(error)
    input()
    raise FileExistsError(error)

print('Building carrier choice file...')
# Create additional identifier key columns
df['plant_pod_ct'] = df['Plant'] + \
    df['Port of Discharge'] + df['Container Type']
df['plant_pod_carrier_ct'] = df['Plant'] + \
    df['Port of Discharge'] + df['Carrier'] + df['Container Type']

# Filter by provided BAF Month, and drop
new_df = (df.loc[df['BAF Month'].str.lower() == baf_month]
            .sort_values('Total Logistics Cost USD')
            .drop_duplicates(['plant_pod_carrier_ct']))

# Group by port pairs and sort by total logs cost
grouped_df = (new_df
              .groupby('plant_pod_ct')
              .apply(pd.DataFrame.sort_values, 'Total Logistics Cost USD'))

# Drop the first multiIndex level to prevent index and column having same name
grouped_df.index = grouped_df.index.droplevel()

# Convert all values the Carrier column, within a group, to a series of lists
grouped_df = grouped_df.groupby('plant_pod_ct').agg({'Carrier': list})

# Take only the first three choices
choices = ['1st choice', '2nd choice',
           '3rd choice', '4th choice', '5th choice']
grouped_df[choices] = pd.DataFrame(grouped_df['Carrier'].to_list(),
                                   index=grouped_df.index)[[0, 1, 2, 3, 4]]
grouped_df.drop('Carrier', inplace=True, axis=1)
grouped_df.reset_index(inplace=True)

# Assembling the final output
final_columns = ['Plant', 'Port of Loading',
                 'Port of Discharge', 'Port of Discharge Name', 'Container Type', 'BAF Month']

final_df = grouped_df.merge(new_df.drop_duplicates('plant_pod_ct')[
                            final_columns + ['plant_pod_ct']], on='plant_pod_ct')


final_df = final_df[final_columns + choices]

date_string = datetime.now().strftime("%d.%m.%Y")
final_df.to_excel(f'Carrier Choice - {date_string}.xlsx', index=False)
print(
    f'Carrier choice file created at {date_string}. You may now close this window.')

input()
