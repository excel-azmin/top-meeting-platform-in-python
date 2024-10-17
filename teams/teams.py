import requests
from msal import ConfidentialClientApplication

# Define the constants required for OAuth2
CLIENT_ID = "b924aa6b-85d0-4c72-84f1-2d84f2c32805"  # Replace with your application client ID
CLIENT_SECRET = ".eQ8Q~XQ1HaNHzd5WJ3V2BEvtZAF.6cj0r2bxbd6"  # Replace with your client secret
AUTHORITY = "https://login.microsoftonline.com/35590502-c139-4026-99ae-be0ce861ad8c"  # Replace with your actual tenant ID
SCOPES = ["https://graph.microsoft.com/.default"]  # Define the required scopes
USER_ID = "sahaaninda@averbangladesh.onmicrosoft.com"  # Replace with the specific user's email or ID
GRAPH_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/events"  # API endpoint to get a specific user's calendar events

# Function to authenticate using ConfidentialClientApplication
def authenticate():
    # Initialize ConfidentialClientApplication
    client = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

    # Acquire a token for the app (client credentials flow)
    result = client.acquire_token_for_client(scopes=SCOPES)

    # Check if the authentication was successful
    if "access_token" in result:
        print("Authentication successful!")
        return result["access_token"]
    else:
        raise Exception(f"Authentication failed: {result}")

# Function to get calendar events
def get_calendar_events(access_token):
    # Set up the request headers with the access token
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Send a GET request to the Microsoft Graph API to get calendar events
    response = requests.get(GRAPH_ENDPOINT, headers=headers)

    # Check if the request was successful
    if response.status_code == 200:
        events = response.json()  # Parse the JSON response
        return events
    else:
        raise Exception(f"Failed to get calendar events: {response.status_code} - {response.text}")

# Call the function to authenticate and get the access token
token = authenticate()

# Get calendar events using the access token
events = get_calendar_events(token)

# Print the calendar events (for testing purposes)
print("Calendar Events:")
for event in events['value']:
    print(f"Event: {event['subject']}, Start: {event['start']['dateTime']}, End: {event['end']['dateTime']}")
