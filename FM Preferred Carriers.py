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
print('Reading Freight Matrix file...this action can take up to 2 minutes.')

files = list(Path(os.getcwd()).glob('Freight Matrix *.xlsx'))
if len(files) == 1:
    df = pd.read_excel(files[0].name, header=1)
else:
    raise FileExistsError(
        'There are multiple Freight Matrix files that fit the pattern! Please include only one in the same directory as this script.')

print('Building carrier choice file...')
# Create additional identifier key columns
df['port_pair'] = df['Port of Loading'] + df['Port of Discharge']
df['port_pair_carrier'] = df['Port of Loading'] + \
    df['Port of Discharge'] + df['Carrier']
new_df = df.drop_duplicates(['port_pair_carrier'])

# Group by port pairs and sort by total logs cost
grouped_df = (new_df
              .groupby('port_pair')
              .apply(pd.DataFrame.sort_values, 'Total Logistics Cost USD'))

# Drop the first multiIndex level to prevent index and column having same name
grouped_df.index = grouped_df.index.droplevel()
grouped_df[['port_pair', 'Total Logistics Cost USD', 'Carrier']]

# Convert all values the Carrier column, within a group, to a series of lists
grouped_df = grouped_df.groupby('port_pair').agg({'Carrier': list})

# Split the series into multiple columns
pd.DataFrame(grouped_df['Carrier'].to_list(), index=grouped_df.index)

# Take only the first five choices
choices = ['1st choice', '2nd choice',
           '3rd choice', '4th choice', '5th choice']
grouped_df[choices] = pd.DataFrame(grouped_df['Carrier'].to_list(),
                                   index=grouped_df.index)[[0, 1, 2, 3, 4]]
grouped_df.drop('Carrier', inplace=True, axis=1)
grouped_df.reset_index(inplace=True)

# Assembling the final output
final_columns = ['Plant', 'Port of Loading',
                 'Port of Discharge', 'Port of Discharge Name', 'Container Type']

final_df = grouped_df.merge(new_df.drop_duplicates('port_pair')[
                            final_columns + ['port_pair']], on='port_pair')


final_df = final_df[final_columns + choices]

date_string = datetime.now().strftime("%d.%m.%Y")
final_df.to_excel(f'Carrier Choice - {date_string}.xlsx')
print(
    f'Carrier choice file created at {date_string}. This window will automatically close in a few seconds.')

input()
