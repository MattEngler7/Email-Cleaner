import os
import pickle
import pytz
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import time

# Set the redirect URI and scopes
REDIRECT_URI = 'https://localhost:5689/'
SCOPES = (
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://mail.google.com/'
)

# Use the client to generate a refresh token
creds = None

# Look and see if the file is already created. If it is created and empty, delete it and create a new one with saved credentials.
# Construct file path dynamically based on the operating system
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

# Initialize counters
total_emails_deleted = 0
batch_size = 900  # Set your desired batch size here

# Measure the time it takes to delete one batch
start_time = time.time()

# Define time_taken_for_batch here
time_taken_for_batch = 0

# Prompt the user for their preferred option
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

while True:
    result = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query, maxResults=batch_size).execute()
    messages = result.get('messages', [])

    if not messages:
        print(f"No more emails found to delete.")
        break

    for message in messages:
        email_id = message['id']
        service.users().messages().delete(userId='me', id=email_id).execute()
        total_emails_deleted += 1

    print(f"Deleted {len(messages)} emails. Total deleted so far: {total_emails_deleted}")

    # Calculate time taken for one batch
    time_taken_for_batch = time.time() - start_time

    # Check if there are more emails to fetch
    if 'nextPageToken' in result:
        page_token = result['nextPageToken']
    else:
        print(f"No more emails found to delete.")
        break

# Calculate the estimated time to delete all emails
estimated_total_time = time_taken_for_batch * (total_emails_deleted / batch_size)
print(f"All selected emails deleted successfully. Total emails deleted: {total_emails_deleted}")
print(f"Estimated time to delete all emails: {estimated_total_time} seconds")
