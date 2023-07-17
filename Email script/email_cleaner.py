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

        creds_file = os.path.join(os.getcwd(), 'credentials.json')

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

    # Iterate through the list of email IDs
    for email_id in email_ids:
        # Get the data of the email
        email_data = service.users().messages().get(userId='me', id=email_id).execute()

        # Extract the email header information
        headers = email_data['payload']['headers']

        sender = ""
        date_str = ""

        # Get the header that contains the sender name
        for header in headers:
            if header['name'] == 'From':
                sender = header['value']
                break

        # Get the header that contains the email date
        for header in headers:
            if header['name'] == 'Date':
                date_str = header['value']
                break

        def clean_date(date_str):
            # Remove timezone offset
            date_str = re.sub(r'([-+]\d{2}):?(\d{2})$', r'\1\2', date_str)

            try:
                date = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
            except ValueError:
                try:
                    date = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S[ %z]')
                except ValueError:
                    try:
                        date = datetime.strptime(date_str[:-6], '%a, %d %b %Y %H:%M:%S %z')
                    except ValueError:
                        if date_str[-1] == ':':
                            date_str += '00'
                        date = datetime.strptime(date_str[:-6], '%a, %d %b %Y %H:%M:%S')
                        timezone_str = date_str[-6:]
                        sign = timezone_str[0]
                        hours = int(timezone_str[1:3])
                        minutes = int(timezone_str[3:])
                        if sign == '+':
                            delta = timedelta(hours=hours, minutes=minutes)
                        else:
                            delta = timedelta(hours=-hours, minutes=-minutes)
                        date = date + delta

            return date

        all_timezones = pytz.all_timezones
        timezone_name = all_timezones[0]
        tz = pytz.timezone(timezone_name)
        new_date = pytz.utc.normalize(datetime.now(pytz.utc))
        date = new_date.astimezone(tz)

        # Query the API for all unread emails
        results = service.users().messages().list(userId='me', q=query).execute()

        messages = []

        if 'messages' in results:
            messages = results['messages']

        for message in messages:
            email_data = service.users().messages().get(userId='me', id=message['id']).execute()
            headers = email_data['payload']['headers']
            sender = ""
            email_id = ""
            for header in headers:
                if header['name'] == 'From':
                    sender = header['value']
                    break
            service.users().messages().trash(userId='me', id=message['id']).execute()
            print(f'Deleted email from {sender} with ID: {email_id}')
    else:
        def cleanup():
            os.system('sudo pkill python')
            os.system('sudo service apache2 reload')
        print(f"No emails for me to delete!")
        print("Enjoy a cleaner mailbox!")
else:
    print("No emails found left for me to delete.")
