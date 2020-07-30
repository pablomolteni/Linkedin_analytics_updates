#!/usr/bin/env python
# coding: utf-8

# Import libraries.
import numpy as np
import pandas as pd

print('FOLLOWERS DATA')
print()
print('Be sure the file you downloaded is on the same directory as this executable file')
print('Remember to backup your Linkedin_followers.xlsx file before you continue')
print()

# Input the downloaded file name from Linkedin containing the new Followers info
while True:
    print('Enter file name of downloaded Linkedin followers data .xls or .xlsx')

    file_name_followers = input()

    if (file_name_followers.find('visitors') != -1):
        print ("File name correct")
        print('File name to be updated:', file_name_followers)
        break
    else:
        print ("File name incorrect. Please check if the file contains Visitors data")
    continue

# Just for checking I list the names of the sheets inside the file.
followers = pd.ExcelFile(file_name_followers)
shs = followers.sheet_names
print('The following sheets will be updated:', shs)

# Load each of the sheets in the update file to a different Pandas dataframe.
# New followers sheet is different than the rest in the data structure.
new_followers = pd.read_excel(file_name_followers,
                              sheet_name='New followers')

# Location, Job function, Seniority, Industry and Company size sheets have equal structure.
followers_Location = pd.read_excel(file_name_followers,
                                   sheet_name='Location')

followers_Job_function = pd.read_excel(file_name_followers,
                                       sheet_name='Job function')

followers_Seniority = pd.read_excel(file_name_followers,
                                    sheet_name='Seniority')

followers_Industry = pd.read_excel(file_name_followers,
                                   sheet_name='Industry')

followers_Company_size = pd.read_excel(file_name_followers,
                                       sheet_name='Company size')

# Convert new_followers Date field to datetime.
new_followers['Date'] = pd.to_datetime(new_followers['Date'])

# I define the date of the files download. This is the last day in which the files have been downloaded from Linkedin.
download_date = new_followers['Date'].max()
download_date = pd.to_datetime(download_date, format='%Y-%m-%d')
print('Last day downloaded from Linkedin:', download_date)

# Add download_date to each of the detailed followers dataframes. This will give us a date for the amount and distribution of each of the sheets in the original downloaded file.
followers_Location['Date'] = download_date
followers_Job_function['Date'] = download_date
followers_Seniority['Date'] = download_date
followers_Industry['Date'] = download_date
followers_Company_size['Date'] = download_date

# Read the Linkedin_followers_New_followers.csv file in which the new data will be stored
Linkedin_followers = pd.read_excel('Linkedin_followers.xlsx')

# Define last_date as the last available date in Linkedin_followers_New_followers.csv
last_date = Linkedin_followers['Date'].max()
print('Previous last date available in database was:',  last_date)
print('New last date will be:', download_date)

# new_followers dataframe is trimmed to include only the data that is not on Linkedin_followers_New_followers.csv. Data from last_date to download_date (included).
new_followers = new_followers.loc[(new_followers['Date'] > last_date) & (new_followers['Date'] <= download_date)]

# Write new info on the excel master data file
from openpyxl import load_workbook

with pd.ExcelWriter('Linkedin_followers.xlsx', mode='a',datetime_format='yyyy-mm-dd', engine="openpyxl") as writer:
    writer.book = load_workbook('Linkedin_followers.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

    new_followers.to_excel(writer, sheet_name='New followers', index=False, header=False,
                           startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='New followers'))+1)
    followers_Location.to_excel(writer, sheet_name='Location', index=False, header=False,
                                startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='Location'))+1)
    followers_Job_function.to_excel(writer, sheet_name='Job function', index=False, header=False,
                                    startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='Job function'))+1)
    followers_Seniority.to_excel(writer, sheet_name='Seniority', index=False, header=False,
                                 startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='Seniority'))+1)
    followers_Industry.to_excel(writer, sheet_name='Industry', index=False, header=False,
                                startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='Industry'))+1)
    followers_Company_size.to_excel(writer, sheet_name='Company size', index=False, header=False,
                                    startrow=len(pd.read_excel('Linkedin_followers.xlsx', sheet_name='Company size'))+1)

writer.save()

print()
print('FINISHED.')
