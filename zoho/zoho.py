import requests

# Function to exchange authorization code for access and refresh tokens
def exchange_auth_code_for_tokens(client_id, client_secret, code, redirect_uri):
    url = "https://accounts.zoho.com/oauth/v2/auth"
    payload = {
        'grant_type': 'authorization_code',
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': redirect_uri,
        'code': code
    }

    response = requests.post(url, data=payload)
    if response.status_code == 200:
        tokens = response.json()
        access_token = tokens.get('access_token')
        refresh_token = tokens.get('refresh_token')
        print("Access Token:", access_token)
        print("Refresh Token:", refresh_token)
        return tokens
    else:
        print("Error exchanging code for tokens:", response.text)
        return None

# Function to refresh access token using the refresh token
def refresh_access_token(client_id, client_secret, refresh_token):
    url = "https://accounts.zoho.com/oauth/v2/token"
    payload = {
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'grant_type': 'refresh_token'
    }

    response = requests.post(url, data=payload)
    if response.status_code == 200:
        access_token = response.json().get('access_token')
        print("New Access Token:", access_token)
        return access_token
    else:
        print("Error refreshing access token:", response.text)
        return None

# Function to retrieve events from Zoho Calendar
def get_zoho_calendar_events(access_token, calendar_id):
    url = f"https://calendar.zoho.com/api/v2/calendars/{calendar_id}/events"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    params = {
        'from': '2024-01-01T00:00:00Z',  # Adjust the date range as per your requirement
        'to': '2024-12-31T23:59:59Z'
    }

    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        events = response.json().get('data', [])
        print(f"Retrieved {len(events)} events.")
        for event in events:
            print(f"Event: {event['title']}, Start: {event['start']}, End: {event['end']}")
        return events
    else:
        print("Error fetching events:", response.text)
        return None

# Example usage
if __name__ == "__main__":
    # Step 1: Replace with your app's credentials
    client_id = "1004.6LAMT9ZO2P7I6R6IR1ODZ26VZABOTV"
    client_secret = "4b903e0c1b53bbc0dbda779bf7086193e5cdc12fb7"
    redirect_uri = "http://localhost:8000/callback"

    # Step 2: Authorization code from Zoho (you'll get this from the URL after authorization)
    authorization_code = "AUTHORIZATION_CODE_FROM_URL"

    # Step 3: Exchange authorization code for access and refresh tokens
    tokens = exchange_auth_code_for_tokens(client_id, client_secret, authorization_code, redirect_uri)

    print(tokens)

    if tokens:
        access_token = tokens['access_token']
        refresh_token = tokens['refresh_token']

        # Step 4: Retrieve calendar events
        calendar_id = 'primary'  # Default calendar ID. Replace if using a specific calendar.
        events = get_zoho_calendar_events(access_token, calendar_id)

        # Step 5: Optionally, refresh the access token when it expires
        new_access_token = refresh_access_token(client_id, client_secret, refresh_token)


https://accounts.zoho.com/oauth/v2/auth

?response_type=code

&client_id=1004.6LAMT9ZO2P7I6R6IR1ODZ26VZABOTV

&scope=AaaServer.profile.READ%2CAaaServer.profile.UPDATE

&redirect_uri=http%3A%2F%2Flocalhost%3A8080%2FZohoOAuth%2Findex.jsp

&state=-5466400890088961855