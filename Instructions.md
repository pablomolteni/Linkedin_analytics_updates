# Installation
Copy the 4 files (Linkedin_data_updater.py and 3 xxxxx.xlsx) to a folder on your computer. 
The 3 Excel files ('Linkedin_followers.xlsx', 'Linkedin_visitors.xlsx' and 'Linkedin_updates.xlsx') will be the datasources to use on a dashboard (Data Studio, Power BI, Tableau, etc.)

# Updating the excel datasources
- Download the 3 files from Linkedin to the same folder as above. Linkedin gives them a name with the following logic linkedinaccountname_followers_xxxxxxx.xls or lnkdaccountname_visitors_xxxxxxx.xls or lnkdaccountname_updates_xxxxxxx.xls
- Back up the 3 datasources files.
- Open terminal.
- Find your way to the folder by using the cd command on Terminal.
- Run Linkedin_data_updater.py by tipping the name of the file (don't forget the extension .py)
- Enter the names of the downloaded files that will update the excel datasources.
