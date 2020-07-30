#!/usr/bin/env python
# coding: utf-8

# Import libraries.
import numpy as np
import pandas as pd

print('VISITORS DATA')
print()
print('Be sure the file you downloaded is on the same directory as this executable file')
print('Remember to backup your Linkedin_visitors.xlsx file before you continue')
print()

# Input the downloaded file name from Linkedin containing the new Visitors info
while True:
    print('Enter file name of downloaded Linkedin visitors data .xls or .xlsx')

    file_name_visitors = input()

    if (file_name_visitors.find('visitors') != -1):
        print ("File name correct")
        print('File name to be updated:', file_name_visitors)
        break
    else:
        print ("File name incorrect. Please check if the file contains Visitors data")
    continue

# Just for checking I list the names of the sheets inside the file.
visitors = pd.ExcelFile(file_name_visitors)
shs = visitors.sheet_names
print('The following sheets will be updated:', shs)

# Load each of the sheets in the update file to a different Pandas dataframe.
# New visitors sheet is different than the rest in the data structure.
visitor_metrics = pd.read_excel(file_name_visitors,
                              sheet_name='Visitor metrics')

# Location, Job function, Seniority, Industry and Company size sheets have equal structure.
visitors_Location = pd.read_excel(file_name_visitors,
                                   sheet_name='Location')

visitors_Job_function = pd.read_excel(file_name_visitors,
                                       sheet_name='Job function')

visitors_Seniority = pd.read_excel(file_name_visitors,
                                    sheet_name='Seniority')

visitors_Industry = pd.read_excel(file_name_visitors,
                                   sheet_name='Industry')

visitors_Company_size = pd.read_excel(file_name_visitors,
                                       sheet_name='Company size')

# Convert new_visitors Date field to datetime.
visitor_metrics['Date'] = pd.to_datetime(visitor_metrics['Date'])

# I define the date of the files download. This is the last day in which the files have been downloaded from Linkedin.
download_date = visitor_metrics['Date'].max()
download_date = pd.to_datetime(download_date, format='%Y-%m-%d')
print('Last day downloaded from Linkedin:', download_date)

# Add download_date to each of the detailed visitors dataframes. This will give us a date for the amount and distribution of each of the sheets in the original downloaded file.
visitors_Location['Date'] = download_date
visitors_Job_function['Date'] = download_date
visitors_Seniority['Date'] = download_date
visitors_Industry['Date'] = download_date
visitors_Company_size['Date'] = download_date

# Read the Linkedin_visitors_New_visitors.csv file in which the new data will be stored
Linkedin_visitors = pd.read_excel('Linkedin_visitors.xlsx')

# Define last_date as the last available date in Linkedin_visitors.xls
last_date = Linkedin_visitors['Date'].max()
print('Previous last date available in database was:',  last_date)
print('New last date will be:', download_date)

# visitors_metrics dataframe is trimmed to include only the data that is not on Linkedin_followers_New_followers.csv. Data from last_date to download_date (included).
visitor_metrics = visitor_metrics.loc[(visitor_metrics['Date'] > last_date) & (visitor_metrics['Date'] <= download_date)]

# Write new data on master excel file
from openpyxl import load_workbook

with pd.ExcelWriter('Linkedin_visitors.xlsx',
                    mode='a',
                    datetime_format='yyyy-mm-dd',
                    engine="openpyxl") as writer:
    writer.book = load_workbook('Linkedin_visitors.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

    visitor_metrics.to_excel(writer, sheet_name='Visitor metrics', index=False, header=False,
                           startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Visitor metrics'))+1)
    visitors_Location.to_excel(writer, sheet_name='Location', index=False, header=False,
                                startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Location'))+1)
    visitors_Job_function.to_excel(writer, sheet_name='Job function', index=False, header=False,
                                    startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Job function'))+1)
    visitors_Seniority.to_excel(writer, sheet_name='Seniority', index=False, header=False,
                                 startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Seniority'))+1)
    visitors_Industry.to_excel(writer, sheet_name='Industry', index=False, header=False,
                                startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Industry'))+1)
    visitors_Company_size.to_excel(writer, sheet_name='Company size', index=False, header=False,
                                    startrow=len(pd.read_excel('Linkedin_visitors.xlsx', sheet_name='Company size'))+1)

writer.save()

print()
print('FINISHED.')
