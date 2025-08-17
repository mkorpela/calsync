import json
import msal
import atexit
import requests
import os.path
import re
from datetime import datetime, timedelta
from icalendar import Calendar
import pytz

# --- CONFIGURATION ---
# Set to True to see what the script would do without making any changes.
DRY_RUN = False
# The number of days forward from today to sync events.
SYNC_DAYS = 30
# The subject to use for events created in the work calendar.
OUTLOOK_EVENT_SUBJECT = "Personal Commitment"
# A robust regex to find our unique ID inside a hidden HTML tag.
UID_REGEX = re.compile(r"SourceUID::([\s\S]*?)<\/p>")
# File locations
TOKEN_CACHE_FILE = ".msal_token_cache.json"
CONFIG_FILE = "config.json"


# --- AUTHENTICATION ---
def get_graph_token():
    """Authenticates with Microsoft Graph using MSAL and returns an access token."""
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: {CONFIG_FILE} not found. Please create it based on config.example.json.")
        return None

    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())

    atexit.register(lambda:
        open(TOKEN_CACHE_FILE, "w").write(cache.serialize())
        if cache.has_state_changed else None)

    app = msal.PublicClientApplication(
        config['client_id'],
        authority=config['authority'],
        token_cache=cache
    )

    token_result = None
    accounts = app.get_accounts()
    if accounts:
        print("Found existing account. Trying to get a token silently...")
        token_result = app.acquire_token_silent(config['scopes'], account=accounts[0])

    if not token_result:
        print("No cached token. Starting interactive login flow...")
        token_result = app.acquire_token_interactive(scopes=config['scopes'])

    if "access_token" in token_result:
        print("Successfully acquired access token!")
        return token_result['access_token']
    else:
        print(f"Failed to acquire token. Error: {token_result.get('error')}")
        return None


# --- ICAL DATA HANDLING ---
def get_ical_events(ical_url):
    """Downloads the iCal file from the provided URL."""
    try:
        response = requests.get(ical_url)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"ERROR: Could not download iCal file. {e}")
        return None

def parse_ical(ical_data):
    """Parses raw iCal data into a list of event dictionaries."""
    cal = Calendar.from_ical(ical_data)
    events = []
    utc = pytz.UTC
    for component in cal.walk():
        if component.name == "VEVENT":
            start_dt = component.get('dtstart').dt
            end_dt = component.get('dtend').dt
            if isinstance(start_dt, datetime) and start_dt.tzinfo is None:
                start_dt = utc.localize(start_dt)
            if isinstance(end_dt, datetime) and end_dt.tzinfo is None:
                end_dt = utc.localize(end_dt)
            events.append({
                "uid": str(component.get('uid')),
                "summary": str(component.get('summary')),
                "start": start_dt,
                "end": end_dt,
                "transp": str(component.get('transp')), # Transparency (busy/free)
                "status": str(component.get('status')), # Status (confirmed, tentative)
            })
    return events

def filter_and_deduplicate_events(events):
    """
    Applies business logic to the raw list of parsed iCal events.
    1. De-duplicates events based on UID.
    2. Filters for events that should be synced (i.e., are 'busy').
    """
    # Stage 1: De-duplicate by UID, keeping the first one seen.
    unique_events_by_uid = {}
    for event in events:
        if event['uid'] not in unique_events_by_uid:
            unique_events_by_uid[event['uid']] = event
    
    deduplicated_events = list(unique_events_by_uid.values())
    
    # Stage 2: Filter for events that block time ('OPAQUE').
    # We also check that the status is 'CONFIRMED'.
    filtered_events = [
        event for event in deduplicated_events
        if event.get('transp') == 'OPAQUE' and event.get('status') == 'CONFIRMED'
    ]
    
    print(f"Initial event count: {len(events)}. After de-duplication: {len(deduplicated_events)}. After filtering for busy/confirmed: {len(filtered_events)}.")
    return filtered_events

# --- GRAPH API HELPERS ---
def get_outlook_events(access_token, start_date, end_date):
    """
    Fetches synced events from Outlook within a date range by looking for a UID tag.
    Handles API pagination to retrieve all events.
    """
    headers = {'Authorization': f'Bearer {access_token}'}
    all_events_data = []
    
    # Construct the initial URL with parameters
    base_url = "https://graph.microsoft.com/v1.0/me/calendarview"
    params = {
        'startDateTime': start_date.isoformat(),
        'endDateTime': end_date.isoformat(),
        '$select': 'id,subject,start,end,body'
    }
    
    next_url = f"{base_url}?{requests.compat.urlencode(params)}"
    
    print("Fetching existing events from Outlook...")
    
    # Loop until there are no more pages of results
    while next_url:
        response = requests.get(next_url, headers=headers)
        if response.status_code != 200:
            print(f"Error fetching Outlook events: {response.text}")
            return []

        data = response.json()
        all_events_data.extend(data.get('value', []))
        next_url = data.get('@odata.nextLink')

    # Now, parse the full list of events retrieved from all pages
    parsed_events = []
    for event in all_events_data:
        body_content = event.get('body', {}).get('content', '')
        if event.get('body', {}).get('contentType') == 'html' and 'SourceUID::' in body_content:
            match = UID_REGEX.search(body_content)
            if match:
                parsed_events.append({
                    'outlook_id': event['id'],
                    'uid': match.group(1),
                    'start': datetime.fromisoformat(event['start']['dateTime'].replace('Z', '+00:00')),
                    'end': datetime.fromisoformat(event['end']['dateTime'].replace('Z', '+00:00'))
                })
    return parsed_events

def create_outlook_event(access_token, event):
    """Creates a new event in Outlook."""
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    body = {
        "subject": OUTLOOK_EVENT_SUBJECT,
        "body": { "contentType": "HTML", "content": f"Synced from private calendar.<p style=\"display:none;\">SourceUID::{event['uid']}</p>" },
        "start": {"dateTime": event['start'].isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": event['end'].isoformat(), "timeZone": "UTC"},
        "showAs": "busy",
        "isReminderOn": False
    }
    response = requests.post('https://graph.microsoft.com/v1.0/me/events', headers=headers, json=body)
    return response.status_code == 201, response.text

def update_outlook_event(access_token, outlook_id, event):
    """Updates an existing Outlook event's time."""
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    body = {
        "start": {"dateTime": event['start'].isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": event['end'].isoformat(), "timeZone": "UTC"}
    }
    url = f'https://graph.microsoft.com/v1.0/me/events/{outlook_id}'
    response = requests.patch(url, headers=headers, json=body)
    return response.status_code == 200, response.text

def delete_outlook_event(access_token, outlook_id):
    """Deletes an Outlook event."""
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f'https://graph.microsoft.com/v1.0/me/events/{outlook_id}'
    response = requests.delete(url, headers=headers)
    return response.status_code == 204, response.text


# --- CORE RECONCILIATION LOGIC ---
def reconcile_events(ical_events, outlook_events):
    """Compares iCal and Outlook events and determines actions needed."""
    to_create, to_update, to_delete = [], [], []
    outlook_map = {event['uid']: event for event in outlook_events}
    ical_uids = {event['uid'] for event in ical_events}

    for ical_event in ical_events:
        if ical_event['uid'] not in outlook_map:
            to_create.append(ical_event)
        else:
            outlook_event = outlook_map[ical_event['uid']]
            if ical_event['start'] != outlook_event['start'] or ical_event['end'] != outlook_event['end']:
                to_update.append({'outlook_id': outlook_event['outlook_id'], **ical_event})
    
    for outlook_event in outlook_events:
        if outlook_event['uid'] not in ical_uids:
            to_delete.append(outlook_event)

    return to_create, to_update, to_delete


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    access_token = get_graph_token()
    if not access_token:
        print("\n--- Authentication Failed. Exiting. ---")
        exit()
    
    print("\n--- Starting Calendar Sync ---")
    with open(CONFIG_FILE, 'r') as f: config = json.load(f)

    ical_data = get_ical_events(config['calendar_url'])
    if not ical_data: exit()
    
    all_ical_events = parse_ical(ical_data)
    
    # Apply our new filtering and de-duplication logic
    processed_ical_events = filter_and_deduplicate_events(all_ical_events)
    
    now = datetime.now(pytz.UTC)
    sync_end_date = now + timedelta(days=SYNC_DAYS)
    ical_events_in_window = [e for e in processed_ical_events if now <= e['start'] < sync_end_date]
    print(f"Found {len(ical_events_in_window)} source events in the next {SYNC_DAYS} days.")

    outlook_events_in_window = get_outlook_events(access_token, now, sync_end_date)
    print(f"Found {len(outlook_events_in_window)} existing '{OUTLOOK_EVENT_SUBJECT}' events in Outlook.")

    to_create, to_update, to_delete = reconcile_events(ical_events_in_window, outlook_events_in_window)
    print(f"\nReconciliation complete: {len(to_create)} to create, {len(to_update)} to update, {len(to_delete)} to delete.")
    
    if DRY_RUN:
        print("\n--- [DRY RUN] No changes will be made. ---")
    else:
        print("\n--- Applying changes... ---")
        for event in to_create:
            success, _ = create_outlook_event(access_token, event)
            print(f"  - CREATE {'Success' if success else 'Failed'}: {event['summary']} at {event['start'].strftime('%Y-%m-%d %H:%M')}")
        for event in to_update:
            success, _ = update_outlook_event(access_token, event['outlook_id'], event)
            print(f"  - UPDATE {'Success' if success else 'Failed'}: Event for UID {event['uid']} to {event['start'].strftime('%Y-%m-%d %H:%M')}")
        for event in to_delete:
            success, _ = delete_outlook_event(access_token, event['outlook_id'])
            print(f"  - DELETE {'Success' if success else 'Failed'}: Event for UID {event['uid']} at {event['start'].strftime('%Y-%m-%d %H:%M')}")
    
    print("\n--- Sync Finished ---")