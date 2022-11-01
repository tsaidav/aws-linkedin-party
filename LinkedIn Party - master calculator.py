# -*- coding: utf-8 -*-
"""
alias: @tsaidav
for: LinkedIn Party
"""
# Importing function from function library
from openpyxl import load_workbook

def main():
    # Change file name here - ensure that the Excel file and Python file are in
    # the same folder
    
    # Change the file name here
    # Ensure that the CSV file and Python file are in the same folder and/or
    # specify the desired working directory
    data_file = 'FILENAME.xlsx'
    
    # Load in workbook
    wb = load_workbook(data_file, data_only=True)
    # Select the first sheet of data
    # Fieldsense form data is on Sheet1 by default
    ws = wb['Sheet1']
    # Separate data into a list of rows
    all_rows = list(ws.rows)
    # Create a point counting dictionary
    point_dict = {}
    # Create a lead dictionary to ensure no duplicates
    lead_dict = {}
    # Create a profile dictionary to ensure no duplicates
    prof_dict = {}
    
    # Create a loop through to see amount of points assigned to candidate
    for row in all_rows[1:len(all_rows)]:
        # Note that column/row counting starts from 0
        # Assign column value for alias
        alias_s = str(row[4].value)
        alias_j = ''.join(filter(str.isalnum, alias_s))
        alias = alias_j.lower()
        # Assign column value for status
        status = str(row[8].value)
        # Assign column value for lead
        lead = str(row[5].value)
        # Assign column value for profile
        prof = str(row[7].value)
        
        # If lead is a duplicate, assign no points and add occurance to lead
        # If lead does not exist yet, add lead key into lead dictionary and
        # set current occurance to 1
        # Likewise for lead's job profile
        if lead in lead_dict:
            lead_dict[lead] += 1
            prof_dict[prof] += 1
        else:
            lead_dict[lead] = 1
            prof_dict[prof] = 1
            # Set initial point number
            p = 0
            # Allocate points in accordance to contact status
            if status == "Not contacted":
                p = 1
            elif status == "I haven't messaged yet, but I will contact them in the next 7 days":
                p = 2
            elif status == "I emailed them; potential referral":
                p = 3
            elif status == "I emailed them; potential diverse referral":
                p = 5
            
            # Check if alias key in dictionary
            # If alias in dictionary - then add profile point
            # If alias not in dictionary - then set profile point
            if alias in point_dict:
                point_dict[alias] += p
            else:
                point_dict[alias] = p
    
    # Uncomment the item for the information that you wish to output to the
    # terminal
    '''
    # Sort data dictionary in ranking of value
    ranking = sorted(point_dict.items(), key=lambda x: x[1], reverse=True)
    # Print leaderboard rankings
    for i in ranking:
        	print(i[0], i[1])
    '''
    '''
    # Sort data dictionary in ranking of value
    lead_occur = sorted(lead_dict.items(), key=lambda x: x[1], reverse=True)
    # Print top leads in value of occurance
    for j in lead_occur:
        	print(j[0], j[1])
    '''
    '''
    # Sort data dictionary in ranking of value
    prof_occur = sorted(prof_dict.items(), key=lambda x: x[1], reverse=True)
    # Print top job profiles in value of occurance
    for k in prof_occur:
        	print(k[0], k[1])
    '''
    
if __name__ == "__main__":
    main()
