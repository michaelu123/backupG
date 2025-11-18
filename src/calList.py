import os
import json
import csv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class excel2(csv.Dialect):
    """Describe the usual properties of Excel-generated CSV files."""
    delimiter = ';'
    quotechar = '"'
    doublequote = True
    skipinitialspace = False
    lineterminator = '\n'
    quoting = csv.QUOTE_MINIMAL


class CsvWriter:
    def __init__(self, f):
        self.fieldNames = [
            "user",
            "id",
            "summary",
            "description",
            "access_role",
        ]

        csv.register_dialect("excel2", excel2)
        self.writer = csv.DictWriter(
            f, self.fieldNames, dialect="excel2", extrasaction="ignore")
        self.writer.writeheader()

    def write(self, entry):
        self.writer.writerow(entry)


# --- Configuration ---
# Path to your service account key file
# 'path/to/your/service_account.json'
SERVICE_ACCOUNT_FILE = "./driveperms-366414-815736cd7941.json"

# The Google Workspace domain (e.g., 'your-company.com')
DOMAIN = "adfc-muenchen.de"  # your-domain.com'

# The scopes you authorized in the Admin Console
SCOPES = [
    'https://www.googleapis.com/auth/admin.directory.user.readonly',
    'https://www.googleapis.com/auth/calendar.readonly'
]
# --- End Configuration ---


def get_all_domain_users(service):
    """Fetches all users in the Google Workspace domain."""
    users = []
    page_token = None
    print("Fetching all users from the domain...")
    while True:
        try:
            results = service.users().list(
                domain=DOMAIN,
                orderBy='email',
                maxResults=500,  # Max allowed per page
                pageToken=page_token
            ).execute()

            users.extend(results.get('users', []))
            page_token = results.get('nextPageToken')
            if not page_token:
                break
        except HttpError as error:
            print(f"An error occurred while fetching users: {error}")
            return None
    print(f"Found {len(users)} users in the domain.")
    return users


def get_calendars_for_user(user_email):
    """Lists all calendars for a single, impersonated user."""
    try:
        # Create new credentials impersonating the user
        impersonated_creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES).with_subject(user_email)

        # Build a new calendar service for the impersonated user
        calendar_service = build(
            'calendar', 'v3', credentials=impersonated_creds)

        page_token = None
        rows = []
        while True:
            calendar_list = calendar_service.calendarList().list(
                pageToken=page_token).execute()
            for calendar_list_entry in calendar_list.get('items', []):
                # print(json.dumps(calendar_list_entry))
                access_role = calendar_list_entry.get(
                    'accessRole', "")
                summary = calendar_list_entry.get('summary', '')
                description = calendar_list_entry.get(
                    'description', '')
                cal_id = calendar_list_entry.get('id')
                row = {
                    "user": user_email,
                    "id": cal_id,
                    "summary": summary,
                    "description": description,
                    "access_role": access_role,
                }
                rows.append(row)
                print(
                    f"  - Summary:'{summary}' Description: '{description}' AccessRole:{access_role} ID:{cal_id}")
                # cal_resource = calendar_service.calendars().get(calendarId=cal_id).execute()
                # # print(json.dumps(cal_resource, indent=2))
                # assert (cal_resource.get('summary', 'No Summary') == summary)
                # assert (cal_resource.get('description',
                #         'No Description') == description)
                # dataOwner = cal_resource.get("dataOwner", "No Data Owner")
                # print(
                #     f"  - Summary:'{summary}' Description: '{description}' AccessRole:{access_role} ID:{cal_id} DataOwner:{dataOwner}")
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break
        return rows
    except HttpError as error:
        # Common errors include a user not having Calendar enabled.
        print(f"  - Could not retrieve calendars. Error: {error.reason}")
        return []
    except Exception as e:
        print(f"  - An unexpected error occurred: {e}")
        return []


def main():
    """
    Main function to enumerate all calendars for all users in a domain.
    """
    # 1. Authenticate as the service account to use the Admin SDK
    admin_creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # Note: The Admin SDK requires a subject to impersonate, even for domain-level tasks.
    # You must use the email of an admin user in your domain.
    admin_user_email = 'michael.uhlenberg.admin@adfc-muenchen.de'  # CHANGE THIS
    delegated_admin_creds = admin_creds.with_subject(admin_user_email)

    try:
        admin_service = build('admin', 'directory_v1',
                              credentials=delegated_admin_creds)
    except Exception as e:
        print(
            f"Failed to build Admin SDK service. Check your credentials and scopes. Error: {e}")
        return

    # 2. Get all users in the domain
    users = get_all_domain_users(admin_service)
    if not users:
        print("No users found or an error occurred. Exiting.")
        return

    # 3. Iterate through each user and list their calendars
    rows = []
    for user in users:
        user_email = user['primaryEmail']
        print(f"\nProcessing calendars for user: {user_email}")
        rows.extend(get_calendars_for_user(user_email))

    with open("cal.csv", "w", encoding="utf-8-sig") as csvOutfile:
        csvWriter = CsvWriter(csvOutfile)
        for row in rows:
            csvWriter.write(row)
    print("Ausgabedatei geschrieben, ", len(rows), "Einträge")


if __name__ == '__main__':
    main()


"""
Q: As an admin, can I enumerate all calenders, not just those of a specific user?
Gemini:
Excellent question! This is a common and important administrative task. The short answer is yes, you can, but not with a single API call.

Unlike the user-centric approach, there isn't a single "list all calendars in the domain" endpoint. Instead, as an administrator, you need to combine two powerful concepts:

The Admin SDK Directory API: To get a list of all users in your Google Workspace domain.
Service Account with Domain-Wide Delegation: To act on behalf of each of those users to list their respective calendars using the Calendar API.
This process allows you to programmatically iterate through every user and enumerate their calendars one by one.

The High-Level Process
Set Up a Service Account: Create a service account in your Google Cloud project.
Enable Domain-Wide Delegation: Configure the service account to have domain-wide authority. This is the key step that allows it to impersonate users.
Authorize the Service Account in Google Workspace: A Google Workspace Super Admin must go into the Admin Console and grant the service account the specific API scopes it needs to access user data.
Write the Script:
Step A: Use the service account to call the Admin SDK Directory API to fetch all users in your domain.
Step B: Loop through each user.
Step C: For each user, create a new, impersonated Calendar API client.
Step D: Use this impersonated client to call the Calendar API's calendarList.list endpoint, which will return the calendars for that specific user.
Step-by-Step Guide with Code
Here is a detailed walkthrough with a Python code example.

1. Prerequisites: Service Account Setup
If you haven't done this before, it's the most critical part.

Create a Service Account:

Go to the Google Cloud Console.
Select your project and click "+ CREATE SERVICE ACCOUNT".
Give it a name (e.g., calendar-enumerator) and click "CREATE AND CONTINUE".
You can skip granting roles for now. Click "DONE".
Find your new service account, click the three-dot menu (Actions), and select "Manage keys".
Click "ADD KEY" -> "Create new key", choose JSON, and click "CREATE". A JSON key file will be downloaded. Keep this file secure! We'll refer to it as service_account.json.
Enable Domain-Wide Delegation:

In the list of service accounts, click on the one you just created.
Go to the "Details" tab.
Copy the Unique ID (it's a long number).
Expand the "Advanced settings" section.
Under "Domain-wide Delegation", click "VIEW GOOGLE WORKSPACE ADMIN CONSOLE".
Authorize the Client in the Google Workspace Admin Console:

You will be taken to admin.google.com -> Security -> API controls -> Domain-wide Delegation.
Click "Add new".
Paste the Unique ID of your service account into the "Client ID" field.
In the "OAuth scopes" field, you must add the permissions your script will need. For this task, you need two:
https://www.googleapis.com/auth/admin.directory.user.readonly (To list users)
https://www.googleapis.com/auth/calendar.readonly (To list calendars for each user)
Click "Authorize".

<<< Python code from above >>>


Key Improvements and Considerations in the Code
Separation of Concerns: The logic is broken into distinct functions (get_all_domain_users, get_calendars_for_user) for clarity and reusability.

Impersonation (with_subject): The core of domain-wide delegation is using .with_subject(user_email). 
This tells the Google API infrastructure, "I am this service account, but for this one request, 
I want you to act as if I am user_email."

Robust Pagination: Both the user listing and calendar listing functions include while loops to 
handle pagination correctly, ensuring you get all results even if there are thousands.

Error Handling: The code includes try...except HttpError blocks. This is crucial because you might 
encounter users who are suspended, don't have the Calendar service enabled, or have other permission issues. 
The script will report the error and continue to the next user instead of crashing.

Admin SDK Impersonation: A subtle but critical point: even to list all users with the Admin SDK, 
the service account must impersonate a domain administrator. That's why delegated_admin_creds is created before calling the Directory API.
This approach provides a powerful and scalable way to audit and manage all calendar resources across your entire Google Workspace organization.



"""
