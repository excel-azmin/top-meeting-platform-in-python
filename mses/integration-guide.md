# Python Integration with On-Premises Exchange Server
## Accessing Calendar Events and Meetings via EWS

---

## Overview
This guide shows how to access calendar events and meeting room bookings from an on-premises Exchange Server using Python. Unlike Microsoft Graph API (which requires internet/cloud), this uses Exchange Web Services (EWS) which works completely offline within your local network.

---

## Architecture Comparison

### Your Current Setup (Microsoft Graph - Cloud):
```
Python App → Internet → Azure AD → Graph API → Exchange Online
```

### New Setup (EWS - On-Premises):
```
Python App → Local Network → Exchange Server (EWS) → Mailbox Database
```

---

## Prerequisites

### 1. Python Libraries Required
```bash
# Install the required library
pip install exchangelib

# Optional: For NTLM authentication (Windows domain auth)
pip install pywin32  # Windows only
# OR
pip install requests-ntlm  # Cross-platform
```

### 2. Exchange Server Requirements
- Exchange Server 2010 SP2 or later
- EWS enabled and configured
- Service account with appropriate permissions

---

## Implementation Guide

### Step 1: Basic Connection to Exchange Server

```python
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib import EWSDateTime, EWSTimeZone, CalendarItem
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from datetime import datetime, timedelta
import urllib3

# Disable SSL warnings for self-signed certificates (development only)
urllib3.disable_warnings()

# For self-signed certificates in development
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

# Define your Exchange server settings
EXCHANGE_SERVER = 'exchange.company.local'  # Your Exchange server hostname
EXCHANGE_USERNAME = 'COMPANY\\svc-roommgmt'  # Domain\Username format
EXCHANGE_PASSWORD = 'YourServiceAccountPassword'
USER_EMAIL = 'user@company.local'  # The mailbox you want to access

def connect_to_exchange():
    """
    Establish connection to on-premises Exchange Server
    """
    # Set up credentials
    credentials = Credentials(
        username=EXCHANGE_USERNAME,
        password=EXCHANGE_PASSWORD
    )
    
    # Configure the Exchange server connection
    config = Configuration(
        server=EXCHANGE_SERVER,
        credentials=credentials
    )
    
    # Connect to the account
    account = Account(
        primary_smtp_address=USER_EMAIL,
        config=config,
        autodiscover=False,  # Set to False for on-premises
        access_type=DELEGATE
    )
    
    return account

# Connect to Exchange
account = connect_to_exchange()
print(f"Connected to mailbox: {account.primary_smtp_address}")
```

### Step 2: Get Calendar Events (Direct Replacement for Your Current Code)

```python
from exchangelib import Account, Credentials, Configuration, DELEGATE
from exchangelib import EWSDateTime, EWSTimeZone
from datetime import datetime, timedelta
import pytz

class ExchangeCalendarClient:
    def __init__(self, server, username, password, email=None):
        """
        Initialize Exchange connection
        
        Args:
            server: Exchange server hostname (e.g., 'exchange.company.local')
            username: Domain\\Username or username@domain.local
            password: Account password
            email: Email address of mailbox to access (if different from username)
        """
        self.server = server
        self.username = username
        self.password = password
        self.email = email or username.split('\\')[1] + '@company.local'
        self.account = None
        
    def authenticate(self):
        """
        Authenticate to Exchange Server (equivalent to your authenticate() function)
        """
        try:
            credentials = Credentials(self.username, self.password)
            config = Configuration(server=self.server, credentials=credentials)
            
            self.account = Account(
                primary_smtp_address=self.email,
                config=config,
                autodiscover=False,
                access_type=DELEGATE
            )
            
            print("Authentication successful!")
            return True
            
        except Exception as e:
            raise Exception(f"Authentication failed: {e}")
    
    def get_calendar_events(self, start_date=None, end_date=None):
        """
        Get calendar events (equivalent to your get_calendar_events() function)
        
        Args:
            start_date: Start date for filtering events
            end_date: End date for filtering events
            
        Returns:
            List of calendar events
        """
        if not self.account:
            raise Exception("Not authenticated. Call authenticate() first.")
        
        # Set default date range if not provided
        tz = EWSTimeZone('UTC')
        if not start_date:
            start_date = datetime.now(pytz.UTC)
        if not end_date:
            end_date = start_date + timedelta(days=30)
        
        # Convert to EWS datetime
        start = EWSDateTime.from_datetime(start_date)
        end = EWSDateTime.from_datetime(end_date)
        
        # Get calendar items
        calendar_items = self.account.calendar.view(
            start=start,
            end=end
        )
        
        # Format events similar to Graph API response
        events = []
        for item in calendar_items:
            event = {
                'subject': item.subject,
                'start': {
                    'dateTime': item.start.isoformat() if item.start else None,
                    'timeZone': str(item.start.tzinfo) if item.start else None
                },
                'end': {
                    'dateTime': item.end.isoformat() if item.end else None,
                    'timeZone': str(item.end.tzinfo) if item.end else None
                },
                'location': item.location,
                'organizer': {
                    'emailAddress': {
                        'address': item.organizer.email_address if item.organizer else None,
                        'name': item.organizer.name if item.organizer else None
                    }
                },
                'attendees': [
                    {
                        'emailAddress': {
                            'address': attendee.mailbox.email_address,
                            'name': attendee.mailbox.name
                        },
                        'type': 'required' if attendee.response_type == 'Organizer' else 'optional'
                    } for attendee in item.required_attendees
                ] if item.required_attendees else [],
                'body': {
                    'content': item.body,
                    'contentType': 'HTML' if item.body else 'Text'
                },
                'isAllDay': item.is_all_day,
                'isCancelled': item.is_cancelled,
                'isRecurring': item.is_recurring,
                'categories': item.categories if item.categories else []
            }
            events.append(event)
        
        return {'value': events}

# Usage - Direct replacement for your current code
def main():
    # Replace these with your on-premises Exchange details
    EXCHANGE_SERVER = "exchange.company.local"
    USERNAME = "COMPANY\\svc-roommgmt"  # or "svc-roommgmt@company.local"
    PASSWORD = "YourServiceAccountPassword"
    USER_EMAIL = "sahaaninda@company.local"  # User whose calendar to access
    
    # Initialize client
    client = ExchangeCalendarClient(
        server=EXCHANGE_SERVER,
        username=USERNAME,
        password=PASSWORD,
        email=USER_EMAIL
    )
    
    # Authenticate (equivalent to your authenticate() function)
    client.authenticate()
    
    # Get calendar events (equivalent to your get_calendar_events() function)
    events = client.get_calendar_events()
    
    # Print the calendar events (same as your current code)
    print("Calendar Events:")
    for event in events['value']:
        print(f"Event: {event['subject']}, Start: {event['start']['dateTime']}, End: {event['end']['dateTime']}")

if __name__ == "__main__":
    main()
```

### Step 3: Working with Meeting Rooms

```python
from exchangelib import Account, Credentials, Configuration, DELEGATE
from exchangelib import CalendarItem, Room, Attendee, Mailbox
from exchangelib import EWSDateTime, EWSTimeZone
from datetime import datetime, timedelta
import pytz

class MeetingRoomManager:
    def __init__(self, server, username, password):
        """
        Initialize Meeting Room Manager for on-premises Exchange
        """
        self.server = server
        self.credentials = Credentials(username, password)
        self.config = Configuration(server=server, credentials=self.credentials)
        self.rooms = {}
        
    def connect_to_room(self, room_email):
        """
        Connect to a specific meeting room mailbox
        """
        try:
            room_account = Account(
                primary_smtp_address=room_email,
                config=self.config,
                autodiscover=False,
                access_type=DELEGATE
            )
            self.rooms[room_email] = room_account
            return room_account
        except Exception as e:
            print(f"Failed to connect to room {room_email}: {e}")
            return None
    
    def get_room_availability(self, room_email, start_time, end_time):
        """
        Check if a room is available during specified time
        """
        if room_email not in self.rooms:
            self.connect_to_room(room_email)
        
        room = self.rooms.get(room_email)
        if not room:
            return None
        
        # Get calendar items for the specified time range
        calendar_items = room.calendar.view(
            start=EWSDateTime.from_datetime(start_time),
            end=EWSDateTime.from_datetime(end_time)
        )
        
        # Check if any events exist in this time range
        events = list(calendar_items)
        is_available = len(events) == 0
        
        return {
            'room': room_email,
            'is_available': is_available,
            'conflicts': [
                {
                    'subject': item.subject,
                    'start': item.start.isoformat(),
                    'end': item.end.isoformat(),
                    'organizer': item.organizer.email_address if item.organizer else None
                } for item in events
            ] if not is_available else []
        }
    
    def book_room(self, room_email, subject, start_time, end_time, 
                  organizer_email, attendees=None, body=None):
        """
        Book a meeting room
        """
        if room_email not in self.rooms:
            self.connect_to_room(room_email)
        
        room = self.rooms.get(room_email)
        if not room:
            return {'success': False, 'error': 'Could not connect to room'}
        
        try:
            # Create calendar item
            item = CalendarItem(
                account=room,
                subject=subject,
                start=start_time,
                end=end_time,
                body=body or f"Meeting booked by {organizer_email}",
                location=room_email,
                required_attendees=[
                    Attendee(
                        mailbox=Mailbox(email_address=organizer_email),
                        response_type='Organizer'
                    )
                ]
            )
            
            # Add additional attendees if provided
            if attendees:
                for attendee_email in attendees:
                    item.required_attendees.append(
                        Attendee(
                            mailbox=Mailbox(email_address=attendee_email),
                            response_type='Required'
                        )
                    )
            
            # Save the meeting
            item.save()
            
            return {
                'success': True,
                'meeting_id': item.id,
                'change_key': item.changekey,
                'message': f'Room {room_email} booked successfully'
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
    
    def get_all_room_schedules(self, room_emails, date=None):
        """
        Get schedules for multiple rooms on a specific date
        """
        if not date:
            date = datetime.now(pytz.UTC).date()
        
        start = datetime.combine(date, datetime.min.time()).replace(tzinfo=pytz.UTC)
        end = datetime.combine(date, datetime.max.time()).replace(tzinfo=pytz.UTC)
        
        schedules = {}
        
        for room_email in room_emails:
            if room_email not in self.rooms:
                self.connect_to_room(room_email)
            
            room = self.rooms.get(room_email)
            if room:
                events = room.calendar.view(
                    start=EWSDateTime.from_datetime(start),
                    end=EWSDateTime.from_datetime(end)
                )
                
                schedules[room_email] = [
                    {
                        'subject': item.subject,
                        'start': item.start.isoformat(),
                        'end': item.end.isoformat(),
                        'organizer': item.organizer.email_address if item.organizer else None,
                        'status': item.legacy_free_busy_status
                    } for item in events
                ]
        
        return schedules

# Example usage
def example_room_management():
    # Initialize manager
    manager = MeetingRoomManager(
        server="exchange.company.local",
        username="COMPANY\\svc-roommgmt",
        password="YourPassword"
    )
    
    # Define room emails
    rooms = [
        "mainconf@company.local",
        "boardroom@company.local",
        "training@company.local"
    ]
    
    # Check room availability
    start = datetime.now(pytz.UTC) + timedelta(hours=1)
    end = start + timedelta(hours=2)
    
    for room in rooms:
        availability = manager.get_room_availability(room, start, end)
        if availability:
            print(f"{room}: {'Available' if availability['is_available'] else 'Busy'}")
    
    # Book a room if available
    result = manager.book_room(
        room_email="mainconf@company.local",
        subject="Project Review Meeting",
        start_time=start,
        end_time=end,
        organizer_email="john.doe@company.local",
        attendees=["jane.smith@company.local", "bob.wilson@company.local"],
        body="Quarterly project review and planning session"
    )
    
    print(f"Booking result: {result}")
    
    # Get all room schedules for today
    schedules = manager.get_all_room_schedules(rooms)
    for room, events in schedules.items():
        print(f"\n{room} - {len(events)} meetings today")
        for event in events:
            print(f"  - {event['subject']} ({event['start']} to {event['end']})")

if __name__ == "__main__":
    example_room_management()
```

### Step 4: Advanced Features - Impersonation for Multiple Users

```python
from exchangelib import Account, Credentials, Configuration, IMPERSONATION
from exchangelib import Identity, Mailbox

class ExchangeImpersonationClient:
    """
    Use impersonation to access multiple user calendars
    Requires ApplicationImpersonation role on service account
    """
    
    def __init__(self, server, service_account, service_password):
        self.credentials = Credentials(service_account, service_password)
        self.config = Configuration(
            server=server,
            credentials=self.credentials
        )
    
    def get_user_calendar(self, user_email):
        """
        Access a user's calendar using impersonation
        """
        # Use impersonation to access the user's mailbox
        account = Account(
            primary_smtp_address=user_email,
            config=self.config,
            autodiscover=False,
            access_type=IMPERSONATION
        )
        
        return account
    
    def get_multiple_user_events(self, user_emails, start_date=None, end_date=None):
        """
        Get calendar events for multiple users
        """
        all_events = {}
        
        for email in user_emails:
            try:
                account = self.get_user_calendar(email)
                
                # Get events
                tz = EWSTimeZone('UTC')
                start = start_date or datetime.now(pytz.UTC)
                end = end_date or (start + timedelta(days=7))
                
                events = account.calendar.view(
                    start=EWSDateTime.from_datetime(start),
                    end=EWSDateTime.from_datetime(end)
                )
                
                all_events[email] = [
                    {
                        'subject': item.subject,
                        'start': item.start.isoformat(),
                        'end': item.end.isoformat(),
                        'location': item.location
                    } for item in events
                ]
                
            except Exception as e:
                all_events[email] = {'error': str(e)}
        
        return all_events

# Usage
client = ExchangeImpersonationClient(
    server="exchange.company.local",
    service_account="COMPANY\\svc-roommgmt",
    service_password="YourPassword"
)

# Get events for multiple users
users = ["user1@company.local", "user2@company.local", "room1@company.local"]
events = client.get_multiple_user_events(users)
```

### Step 5: Error Handling and Connection Pooling

```python
import logging
from exchangelib import Account, Credentials, Configuration, DELEGATE
from exchangelib.errors import ErrorAccessDenied, ErrorItemNotFound
from exchangelib.protocol import BaseProtocol
import ssl

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class RobustExchangeClient:
    def __init__(self, server, username, password, 
                 verify_ssl=True, max_connections=4):
        """
        Robust Exchange client with error handling
        """
        self.server = server
        self.username = username
        self.password = password
        
        # Configure SSL
        if not verify_ssl:
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
        
        # Set connection pool size
        BaseProtocol.SESSION_POOLSIZE = max_connections
        
        self.account = None
        
    def connect(self, email):
        """
        Connect with error handling and retry logic
        """
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                credentials = Credentials(self.username, self.password)
                config = Configuration(
                    server=self.server,
                    credentials=credentials,
                    max_connections=4
                )
                
                self.account = Account(
                    primary_smtp_address=email,
                    config=config,
                    autodiscover=False,
                    access_type=DELEGATE
                )
                
                # Test connection
                _ = self.account.root.effective_rights
                
                logger.info(f"Successfully connected to {email}")
                return True
                
            except ErrorAccessDenied:
                logger.error(f"Access denied to {email}. Check permissions.")
                raise
                
            except Exception as e:
                retry_count += 1
                logger.warning(f"Connection attempt {retry_count} failed: {e}")
                if retry_count >= max_retries:
                    logger.error(f"Failed to connect after {max_retries} attempts")
                    raise
                    
        return False
    
    def safe_get_events(self, start_date=None, end_date=None):
        """
        Get events with proper error handling
        """
        if not self.account:
            raise Exception("Not connected. Call connect() first.")
        
        try:
            start = start_date or datetime.now(pytz.UTC)
            end = end_date or (start + timedelta(days=30))
            
            events = self.account.calendar.view(
                start=EWSDateTime.from_datetime(start),
                end=EWSDateTime.from_datetime(end)
            )
            
            return list(events)
            
        except ErrorItemNotFound:
            logger.warning("No calendar found for this account")
            return []
            
        except Exception as e:
            logger.error(f"Error fetching events: {e}")
            raise

# Usage with error handling
try:
    client = RobustExchangeClient(
        server="exchange.company.local",
        username="COMPANY\\svc-roommgmt",
        password="YourPassword",
        verify_ssl=False  # For self-signed certificates
    )
    
    client.connect("user@company.local")
    events = client.safe_get_events()
    
    print(f"Found {len(events)} events")
    
except ErrorAccessDenied:
    print("Access denied. Please check service account permissions.")
except Exception as e:
    print(f"An error occurred: {e}")
```

---

## Configuration Requirements

### 1. Exchange Server Configuration

```powershell
# PowerShell commands to configure Exchange for Python access

# 1. Create service account (if not already created)
New-ADUser -Name "svc-python" `
    -UserPrincipalName "svc-python@company.local" `
    -AccountPassword (ConvertTo-SecureString "ComplexP@ssw0rd" -AsPlainText -Force) `
    -Enabled $true `
    -PasswordNeverExpires $true

# 2. Enable mailbox for service account
Enable-Mailbox -Identity "svc-python@company.local"

# 3. Grant ApplicationImpersonation role
New-ManagementRoleAssignment -Role ApplicationImpersonation `
    -User "svc-python@company.local"

# 4. Grant access to specific mailboxes (if not using impersonation)
Add-MailboxPermission -Identity "user@company.local" `
    -User "svc-python@company.local" `
    -AccessRights FullAccess

# 5. Check EWS virtual directory
Get-WebServicesVirtualDirectory | Format-List Name, *Url*, *Authentication*
```

### 2. Network Requirements

```yaml
# Your Python application needs network access to:
Exchange_Server: exchange.company.local
Port: 443 (HTTPS)
Protocol: HTTPS/EWS

# DNS Resolution must work for:
- exchange.company.local
- autodiscover.company.local (if using autodiscover)

# Firewall rules:
Source: Python_App_Server
Destination: Exchange_Server
Port: 443
Protocol: TCP
```

### 3. SSL/TLS Considerations

```python
# For production with proper certificates:
config = Configuration(
    server='exchange.company.local',
    credentials=credentials,
    # No special SSL configuration needed
)

# For development with self-signed certificates:
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import urllib3

# Disable warnings (development only!)
urllib3.disable_warnings()
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

# For custom certificate validation:
import ssl
from exchangelib.protocol import BaseProtocol

class CustomHTTPAdapter(BaseProtocol.HTTP_ADAPTER_CLS):
    def init_poolmanager(self, *args, **kwargs):
        kwargs['ssl_context'] = ssl.create_default_context()
        kwargs['ssl_context'].load_cert_chain(
            certfile='/path/to/client_cert.pem',
            keyfile='/path/to/client_key.pem'
        )
        return super().init_poolmanager(*args, **kwargs)

BaseProtocol.HTTP_ADAPTER_CLS = CustomHTTPAdapter
```

---

## Migration Checklist

### From Graph API to EWS:

- [ ] Install `exchangelib` library
- [ ] Replace Graph API authentication with EWS credentials
- [ ] Update API endpoints to use local Exchange server
- [ ] Modify data models (Graph API response → EWS objects)
- [ ] Update error handling for EWS-specific errors
- [ ] Test with local Exchange server
- [ ] Remove internet/Azure dependencies

### Key Differences:

| Aspect | Graph API (Current) | EWS (On-Premises) |
|--------|-------------------|-------------------|
| Authentication | OAuth2/Azure AD | NTLM/Basic/Kerberos |
| Library | `msal` + `requests` | `exchangelib` |
| Server | graph.microsoft.com | exchange.company.local |
| Internet | Required | Not Required |
| Credentials | App ID + Secret | Username + Password |
| API Format | REST/JSON | SOAP/XML (handled by library) |

---

## Testing Your Implementation

```python
# test_exchange_connection.py
import sys
from exchangelib import Account, Credentials, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import urllib3

def test_connection():
    # Test configuration
    TEST_CONFIG = {
        'server': 'exchange.company.local',
        'username': 'COMPANY\\testuser',
        'password': 'TestPassword',
        'email': 'testuser@company.local'
    }
    
    print("Testing Exchange Server Connection...")
    print(f"Server: {TEST_CONFIG['server']}")
    print(f"Username: {TEST_CONFIG['username']}")
    
    try:
        # Disable SSL verification for testing
        urllib3.disable_warnings()
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
        
        # Connect
        credentials = Credentials(
            TEST_CONFIG['username'], 
            TEST_CONFIG['password']
        )
        config = Configuration(
            server=TEST_CONFIG['server'],
            credentials=credentials
        )
        account = Account(
            primary_smtp_address=TEST_CONFIG['email'],
            config=config,
            autodiscover=False,
            access_type=DELEGATE
        )
        
        # Test access
        print(f"✓ Connected to: {account.primary_smtp_address}")
        
        # Get folder info
        print(f"✓ Calendar folder: {account.calendar.name}")
        
        # Count items
        item_count = account.calendar.total_count
        print(f"✓ Calendar items: {item_count}")
        
        print("\n✅ Connection test PASSED!")
        return True
        
    except Exception as e:
        print(f"\n❌ Connection test FAILED!")
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    sys.exit(0 if test_connection() else 1)
```

---

## Troubleshooting

### Common Issues and Solutions:

1. **401 Unauthorized**
   - Check username format (DOMAIN\\username or username@domain.local)
   - Verify password is correct
   - Check account is not locked

2. **Certificate Error**
   - Add certificate to trusted store
   - Use NoVerifyHTTPAdapter for development
   - Install proper certificates for production

3. **Access Denied**
   - Verify ApplicationImpersonation role is granted
   - Check mailbox permissions
   - Ensure account is enabled

4. **Connection Timeout**
   - Check firewall rules
   - Verify DNS resolution
   - Test with telnet: `telnet exchange.company.local 443`

5. **Module Import Error**
   ```bash
   # Ensure all dependencies are installed
   pip install exchangelib
   pip install python-dateutil
   pip install pytz
   ```

---

## Performance Optimization

```python
# Batch operations for better performance
from exchangelib import Account, CalendarItem
from exchangelib.queryset import Q

class OptimizedExchangeClient:
    def __init__(self, account):
        self.account = account
    
    def bulk_get_events(self, subject_filter=None):
        """
        Get events with filtering
        """
        query = self.account.calendar.all()
        
        if subject_filter:
            query = query.filter(subject__contains=subject_filter)
        
        # Limit fields for better performance
        return query.only(
            'subject', 
            'start', 
            'end', 
            'location'
        ).order_by('-start')
    
    def bulk_create_events(self, events_data):
        """
        Create multiple events efficiently
        """
        calendar_items = []
        
        for data in events_data:
            item = CalendarItem(
                account=self.account,
                subject=data['subject'],
                start=data['start'],
                end=data['end']
            )
            calendar_items.append(item)
        
        # Bulk save
        return self.account.bulk_create(items=calendar_items)
```

---

## Summary

This implementation provides:
- ✅ Complete offline operation
- ✅ Direct integration with on-premises Exchange
- ✅ No internet/cloud dependencies
- ✅ Same functionality as Graph API
- ✅ Better performance (local network)
- ✅ Full control over data

The main change is replacing `msal` + `requests` with `exchangelib`, and pointing to your local Exchange server instead of Microsoft Graph API.
