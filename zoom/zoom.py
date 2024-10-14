import requests
import base64

# Replace with your Client ID, Client Secret, and Account ID from Zoom
ACCOUNT_ID = '80CCORwoS_6-R-g-cwQQsg'
CLIENT_ID = 'b9ukZW20Q6KjhdPT2maoFg'
CLIENT_SECRET = 'ecv46Kkt9IKXeWs1erjy7oG53koCkUav'

# Zoom OAuth 2.0 Token URL
TOKEN_URL = 'https://zoom.us/oauth/token'

# Function to get access token
def get_access_token():
    response = requests.post(TOKEN_URL, {"grant_type": "account_credentials",
                                         "account_id": ACCOUNT_ID,
                                         "client_id": CLIENT_ID,
                                         "client_secret": CLIENT_SECRET})
    return  response.json()['access_token']
# Function to retrieve Zoom events based on role and event status
def get_zoom_events(access_token):
    api_url = 'https://api.zoom.us/v2/users/me/meetings'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    params = {
        'type': 'upcoming',  # You can change this to 'scheduled', 'live', 'upcoming'
        'page_size': 30,  # Adjust page size as needed
    }
    return requests.get(api_url, headers=headers,params=params).json()


# Main function to execute the flow
if __name__ == "__main__":
    # Step 1: Get the access token using client credentials
    token = get_access_token()
    print(token)

    meetings = get_zoom_events(token)
    if meetings:
        print("Meetings:")
        for meeting in meetings.get('meetings', []):
            print(f"Meeting ID: {meeting['id']}")
            print(f"Topic: {meeting['topic']}")
            print(f"Start Time: {meeting['start_time']}")
            print(f"Join URL: {meeting['join_url']}")
            print("-" * 40)
    else:
        print("No meetings found.")