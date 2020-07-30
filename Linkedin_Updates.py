#!/usr/bin/env python
# coding: utf-8

# Import libraries.
import numpy as np
import pandas as pd

print('UPDATES DATA')
print()
print('Be sure the file you downloaded is on the same directory as this executable file')
print('Remember to backup your Linkedin_updates.xlsx file before you continue')
print()

# Input the downloaded file name from Linkedin containing the new Updates info
while True:
    print('Enter file name of downloaded Linkedin UPDATES data .xls or .xlsx')

    file_name_updates = input()

    if (file_name_updates.find('updates') != -1):
        print ("File name correct")
        print('File name to be updated:', file_name_updates)
        break
    else:
        print ("File name incorrect. Please check if the file contains Updates data")
    continue

# Just for checking I list the names of the sheets inside the file.
updates = pd.ExcelFile(file_name_updates)
shs = updates.sheet_names
print('The following sheets will be updated:', shs)

# Load each of the sheets in the udate file to a different Pandas dataframe.
new_metrics = pd.read_excel(file_name_updates,
                               sheet_name='Update metrics (aggregated)',
                               header=1
                              )

new_engagement = pd.read_excel(file_name_updates,
                               sheet_name='Update engagement',
                               header=1
                              )

# Convert update_metrics Date field to datetime.
new_metrics['Date'] = pd.to_datetime(new_metrics['Date'])
# Convert update_engagement dates fields to datetime.
new_engagement['Created date'] = pd.to_datetime(new_engagement['Created date'])
new_engagement['Campaign start date'] = pd.to_datetime(new_engagement['Campaign start date'])
new_engagement['Campaign end date'] = pd.to_datetime(new_engagement['Campaign end date'])

# Define download_date as the last date with data in the downloaded file from Linkedin
first_date = new_metrics['Date'].min()
download_date = new_metrics['Date'].max()

# Load master file Linkedin_updates.xls on Linkedin_updates_metrics dataframe
old_metrics = pd.read_excel('Linkedin_updates.xlsx',sheet_name='Update metrics (aggregated)')
old_engagement = pd.read_excel('Linkedin_updates.xlsx', sheet_name='Update engagement')

# Convert update_metrics Date field to datetime.
old_metrics['Date'] = pd.to_datetime(old_metrics['Date'])

# Convert update_engagement dates fields to datetime.
old_engagement['Created date'] = pd.to_datetime(old_engagement['Created date'])
old_engagement['Campaign start date'] = pd.to_datetime(old_engagement['Campaign start date'])
old_engagement['Campaign end date'] = pd.to_datetime(old_engagement['Campaign end date'])

# Define last_date as the last date available in master update file
last_date = old_metrics['Date'].max()

old_engagement.set_index('Update link', inplace=True)

new_engagement['Type of content'] = np.nan
new_engagement['Content Format'] = np.nan
new_engagement['Language'] = np.nan
new_engagement.set_index('Update link', inplace=True)

df_engagement = old_engagement.append(new_engagement)
df_engagement['Type of content'].update(old_engagement['Type of content'])
df_engagement['Content Format'].update(old_engagement['Content Format'])
df_engagement['Language'].update(old_engagement['Language'])

df_engagement=df_engagement.drop_duplicates(subset=['Update title'], keep='last')

#  Trim the data from the downloaded file to update the info from the last date available on the master file.
df_metrics = old_metrics.loc[(old_metrics['Date'] < first_date)]

df_metrics = df_metrics.append(new_metrics, ignore_index=True)

from openpyxl import load_workbook

with pd.ExcelWriter('Linkedin_updates.xlsx',
#                     mode='a',
                    datetime_format='yyyy-mm-dd',
                    engine="openpyxl") as writer:
    writer.book = load_workbook('Linkedin_updates.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

    df_metrics.to_excel(writer, sheet_name='Update metrics (aggregated)', index=False, header=True)

    df_engagement.to_excel(writer, sheet_name='Update engagement', index=True, header=True)


writer.save()

print()
print('FINISHED.')
