import time
import requests
import pandas as pd

bearer_token = "BEARER_TOKEN"  # replace with your bearer token
headers = {
    'Authorization': 'Bearer ' + bearer_token,
}

def get_group_details(group_id):
    url = f"https://webexapis.com/v1/groups/{group_id}?includeMembers=true"
    response = requests.get(url, headers=headers)
    return response.json()

def get_person_details(person_id):
    while True:  # keep trying until successful
        url = f"https://webexapis.com/v1/people/{person_id}"
        response = requests.get(url, headers=headers)
        if response.status_code == 429:  # rate limit error
            print("Rate limit exceeded. Sleeping for 5 seconds.")
            time.sleep(5)  # wait for 5 seconds before trying again
        else:
            return response.json()

def get_group_members(group_id):
    count = 500
    emails = []
    memberSize = get_group_details(group_id)["memberSize"]

    for startIndex in range(1, memberSize + 1, count):  
        while True:  # keep trying until successful
            url = f"https://webexapis.com/v1/groups/{group_id}/members?startIndex={startIndex}&count={count}"
            response = requests.get(url, headers=headers)
            if response.status_code == 429:  # rate limit error
                print("Rate limit exceeded. Sleeping for 5 seconds.")
                time.sleep(5)  # wait for 5 seconds before trying again
            else:
                data = response.json()
                for member in data["members"]:  # updated this line
                    person_details = get_person_details(member["id"])  
                    emails.append(person_details["emails"][0])
                break  # exit the loop if the request was successful
    return emails

# Load group IDs from CSV file
input_df = pd.read_csv('input.csv')  # replace with your input CSV file path

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('group_emails.xlsx', engine='xlsxwriter')

# Process each group
for i, row in input_df.iterrows():
    group_id = row['group_id']  # replace with your actual group_id column name
    group_details = get_group_details(group_id)  # fetch group details
    emails = get_group_members(group_id)

    # Write to a new sheet in the Excel file
    output_df = pd.DataFrame(emails, columns=["Emails"])
    output_df.to_excel(writer, sheet_name=group_details['displayName'], index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.close()
