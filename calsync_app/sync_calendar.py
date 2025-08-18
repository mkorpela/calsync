import json
import msal
import atexit
import requests
import os.path
import re
from datetime import datetime, timedelta
from icalendar import Calendar, vCalAddress
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
    """Parses raw iCal data into a list of event dictionaries, including attendees and recurrence."""
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
            
            attendees = []
            for prop in component.get('attendee', []):
                if isinstance(prop, vCalAddress):
                    attendees.append({
                        "email": prop.replace('mailto:', ''),
                        "status": prop.params.get('PARTSTAT')
                    })

            events.append({
                "uid": str(component.get('uid')),
                "summary": str(component.get('summary')),
                "start": start_dt,
                "end": end_dt,
                "transp": str(component.get('transp')),
                "status": str(component.get('status')),
                "attendees": attendees,
                "recurrence-id": component.get('recurrence-id')
            })
    return events

def process_ical_events(events, config):
    """
    Applies business logic: de-duplicates (prioritizing instances), 
    filters based on user participation, and filters for working hours.
    """
    user_email = config.get("user_email", "").lower()
    if not user_email:
        print("Warning: 'user_email' not set in config. Filtering will be based on event transparency only.")

    # Stage 1: De-duplicate by UID, prioritizing events with a recurrence-id (instances).
    unique_events_by_uid = {}
    for event in events:
        uid = event['uid']
        is_instance = event.get('recurrence-id') is not None
        
        if uid not in unique_events_by_uid or (is_instance and unique_events_by_uid[uid].get('recurrence-id') is None):
            unique_events_by_uid[uid] = event
    
    deduplicated_events = list(unique_events_by_uid.values())
    
    # Stage 2: Filter based on user's participation status.
    filtered_events = []
    for event in deduplicated_events:
        if event.get('status') != 'CONFIRMED':
            continue

        is_user_attendee = False
        user_partstat = None
        if user_email:
            for attendee in event.get('attendees', []):
                if attendee.get('email', '').lower() == user_email:
                    is_user_attendee = True
                    user_partstat = attendee.get('status')
                    break
        
        if is_user_attendee:
            if user_partstat != 'DECLINED':
                filtered_events.append(event)
        else:
            if event.get('transp') == 'OPAQUE':
                filtered_events.append(event)

    print(f"Initial event count: {len(events)}. After de-duplication: {len(deduplicated_events)}. After participation/busy filtering: {len(filtered_events)}.")
    
    # Stage 3: Filter for events within working hours.
    try:
        working_hours = config['working_hours']
        local_tz = pytz.timezone(config['timezone'])
    except (KeyError, pytz.exceptions.UnknownTimeZoneError) as e:
        print(f"Warning: Could not apply working hours filter. Missing or invalid 'working_hours'/'timezone' in config. {e}")
        return filtered_events

    working_hours_events = []
    for event in filtered_events:
        event_start_local = event['start'].astimezone(local_tz)
        event_end_local = event['end'].astimezone(local_tz)
        
        current_date = event_start_local.date()
        is_in_working_hours = False
        while current_date <= event_end_local.date():
            day_name = current_date.strftime("%A").lower()
            if day_name in working_hours:
                for slot in working_hours[day_name]:
                    try:
                        slot_start_time = datetime.strptime(slot['start'], '%H:%M').time()
                        slot_end_time = datetime.strptime(slot['end'], '%H:%M').time()
                        
                        slot_start_dt = local_tz.localize(datetime.combine(current_date, slot_start_time))
                        slot_end_dt = local_tz.localize(datetime.combine(current_date, slot_end_time))

                        if max(event_start_local, slot_start_dt) < min(event_end_local, slot_end_dt):
                            working_hours_events.append(event)
                            is_in_working_hours = True
                            break
                    except (ValueError, KeyError):
                        continue
            if is_in_working_hours:
                break
            current_date += timedelta(days=1)
            
    print(f"After filtering for working hours: {len(working_hours_events)} events remain.")
    return working_hours_events

# --- GRAPH API HELPERS ---
def get_outlook_events(access_token, start_date, end_date):
    """
    Fetches synced events from Outlook within a date range by looking for a UID tag.
    Handles API pagination to retrieve all events.
    """
    headers = {'Authorization': f'Bearer {access_token}'}
    all_events_data = []
    
    base_url = "https://graph.microsoft.com/v1.0/me/calendarview"
    params = {
        'startDateTime': start_date.isoformat(),
        'endDateTime': end_date.isoformat(),
        '$select': 'id,subject,start,end,body',
        '$filter': f"subject eq '{OUTLOOK_EVENT_SUBJECT}'"
    }
    
    next_url = f"{base_url}?{requests.compat.urlencode(params)}"
    
    print("Fetching existing events from Outlook...")
    
    while next_url:
        response = requests.get(next_url, headers=headers)
        if response.status_code != 200:
            print(f"Error fetching Outlook events: {response.text}")
            return []

        data = response.json()
        all_events_data.extend(data.get('value', []))
        next_url = data.get('@odata.nextLink')

    parsed_events = []
    for event in all_events_data:
        body_content = event.get('body', {}).get('content', '')
        if 'SourceUID::' in body_content:
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
    try:
        with open(CONFIG_FILE, 'r') as f: config = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: {CONFIG_FILE} not found. Please create it based on config.example.json.")
        exit()

    ical_data = get_ical_events(config['calendar_url'])
    if not ical_data: exit()
    
    all_ical_events = parse_ical(ical_data)
    
    now = datetime.now(pytz.UTC)
    sync_end_date = now + timedelta(days=SYNC_DAYS)
    
    ical_events_in_window = [
        e for e in all_ical_events if now <= e.get('start', now) < sync_end_date
    ]
    print(f"\nTotal events parsed from iCal: {len(all_ical_events)}.")
    print(f"Found {len(ical_events_in_window)} events in the next {SYNC_DAYS}-day sync window.")
    
    processed_ical_events = process_ical_events(ical_events_in_window, config)
    print(f"After all filtering, {len(processed_ical_events)} source events remain to be synced.")

    outlook_events_in_window = get_outlook_events(access_token, now, sync_end_date)
    print(f"Found {len(outlook_events_in_window)} existing '{OUTLOOK_EVENT_SUBJECT}' events in Outlook.")

    to_create, to_update, to_delete = reconcile_events(processed_ical_events, outlook_events_in_window)
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