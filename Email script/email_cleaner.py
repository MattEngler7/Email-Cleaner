import os
import pickle
import pytz
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import re

# Set the redirect URI and scopes
REDIRECT_URI = 'https://localhost:5689/'
SCOPES = (
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://mail.google.com/'
)

# Use the client to generate a refresh token
creds = None

# Look and see if file is already created. If it is created and empty delete it and create a new one with saved credentials.
# Construct file path dynamically based on operating system
if os.name == 'posix':  # Unix-based system (Mac or Linux)
    file_path = '/Users/matthewengler/Desktop/Email script/token.pickle'
elif os.name == 'nt':  # Windows system
    file_path = 'C:\\Users\\matthewengler\\Desktop\\Email script\\token.pickle'
else:
    raise Exception("Unsupported operating system")

if os.path.exists(file_path):
    with open(file_path, 'r', encoding='ISO-8859-1') as f:
        file_content = f.read()
        if not file_content:
            os.remove(file_path)
            print('File not found...')
            print('Creating a new one...')

if os.path.exists('token.pickle'):
    os.remove('token.pickle')

if os.path.exists('token.pickle'):
    print('Loading Saved Data from File...')
    with open('token.pickle', 'rb') as f:
        creds = pickle.load(f)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        print('Refreshing Token...')
        creds.refresh(Request())
    else:
        print('Please Sign in...')

        script_directory = os.path.dirname(os.path.abspath(__file__))
        creds_file = os.path.join(script_directory, 'credentials.json')

        flow = InstalledAppFlow.from_client_secrets_file(
            creds_file, SCOPES)

        # Use the redirect URI when running the local server
        flow.redirect_uri = REDIRECT_URI
        flow.run_local_server(port=5689, prompt='consent', authorization_prompt_message="")

        creds = flow.credentials

    # Save the credentials for the next run
    with open('token.pickle', 'wb') as f:
        print('Saving...')
        print(" ")
        pickle.dump(creds, f)

# Create the service object
service = build('gmail', 'v1', credentials=creds)

# Get the total number of emails in the inbox
inbox_result = service.users().messages().list(userId='me', q='is:inbox').execute()
total_emails = inbox_result.get('resultSizeEstimate', 0)

# Set the batch size based on the total number of emails
batch_size = min(total_emails, 500)

# Prompt user for their preferred option
option = ""
while option not in ['1', '2', '3']:
    option = input("Enter ( 1 ) to delete ALL unread emails or ( 2 ) to delete emails OLDER THAN A MONTH or ( 3 ) to EXIT:  ")
    if option not in ['1', '2', '3']:
        print(" ")
        print("Invalid option. Please try again.")

# Get the list of all unread emails in the inbox
if option == '1':
    query = "is:Unread"
elif option == '2':
    # Get the current UTC time
    now = datetime.now(pytz.utc)

    # Calculate the difference in days
    difference = timedelta(days=30)
    d = now - difference

    query = "is:unread before:" + d.strftime("%Y/%m/%d")

# Gives the user a choice to close the program    
elif option == '3':
    print("Exiting the program, have a nice day!")
    exit()    

result = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query).execute()
print('Reading emails to delete...')

# Check if the 'messages' key is present in the result
if 'messages' in result:
    # Get the list of email IDs
    email_ids = [email['id'] for email in result['messages']]

    # Initialize the counter
    deleted_count = 0
    print(f"Deleting...")

    while email_ids:
        batch_request = {'ids': email_ids[:batch_size]}
        response = service.users().messages().batchDelete(userId='me', body=batch_request).execute()

        if 'deleted' in response:
            deleted_count += len(response['deleted'])

        email_ids = email_ids[batch_size:]

    print(f"Deleted {deleted_count} emails.")
    print(f"All selected emails deleted successfully.")
else:
    print(f"No emails found to delete.")
    print(f"Enjoy a cleaner email. Have a nice day!")
