# Linkedin_analytics_data
Simple Python code for updating data downloaded from Linkedin.
Please feel free to send feedback and suggest any changes.

## Problem to solve
> **Friend:** It took me about a whole day to update Linkedin data from my customers.

> **Me:** Why?

> **Friend:** I need to download datasets from 3 different Excel files to feed my Data Studio dashboards.

> **Me:** Doesn't Linkedin have tools for marketing people like you?

> **Friend:** They are expensive and they don't always work as I'd like. I just need all the data updated without all the copy-pasting.

> **Me:** Ok. So you want to download those 3 files (followers, visitors and updates) from Linkedin and *magically* update the datasource file (.xlsx) for your dashboards?

> **Friend:** Yes

> **Me:** You know that is not the best solution?

> **Friend:** Yes, but it is all I need for now.

> **Me:** Ok. Let me give the online Python course a try.

## Objectives:
- Give my friend a simple solution to allow fast update on excel files used as datasources without copy-pasting.
- Give myself a real life usage for the Python course I've taken.
- Learn Github basics

## What's inside? 
You'll find 4 files. The Python script to be run every time you need to update your Linkedin data.
- Linkedin_data_updater.py
The other 3 files are the Excel datasources which will feed the Data Studio dashboard.
- Linkedin_followers.xlsx
- Linkedin_visitors.xlsx
- Linkedin_updates.xlsx

## Who can find this useful?
Social media communicators, marketing people, community managers or anyone who doesn't want to pay for a premium Linkedin reporting tool.
Anyone writting their first piece of code.

## Comments
- For the script to run all other files need to be in the same folder.
- I haven't started with Linkedin developer tools and this is my sort of mvp for my "client".
- The script is supposed to be run every other week or monthly. On several sheets inside the excel files the script saves data for a particular day. Example. On Linkedin_followers.xlsx/Location the script will store the followers data for every follower who declares their location to Linkedin on a particular day (the day on which data is downloaded). This feature will give you several *pictures* of where your followers are and be able to check how many new followers you have.

## Next steps
- [X] Unify the three scripts in one. There is a lag when Python imports numpy and pandas libraries.
- [ ] Explore Linkedin developer tools to get real time information
- [ ] Explore similar tools for Twitter, Instagram, Facebook, etc.. (Break this step after completing previous 2.
