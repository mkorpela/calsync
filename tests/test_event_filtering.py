import pytest
from datetime import datetime
import pytz
from calsync_app.sync_calendar import process_ical_events

# Define a timezone for consistent testing
HELSINKI = pytz.timezone("Europe/Helsinki")

# A fixed Monday for predictable weekday calculations
MONDAY = datetime(2025, 8, 18) 
SATURDAY = datetime(2025, 8, 23)

@pytest.fixture
def sample_config():
    """Provides a sample configuration for testing working hours."""
    return {
        "timezone": "Europe/Helsinki",
        "working_hours": {
            "monday": [
                {"start": "09:00", "end": "17:00"}, 
                {"start": "21:00", "end": "23:00"}
            ],
            "tuesday": [
                {"start": "09:00", "end": "12:00"}
            ]
        }
    }

@pytest.fixture
def sample_raw_events():
    """Provides a list of raw iCal events with various conditions."""
    events = [
        # 1. Standard event, completely within working hours
        {'uid': 'inside_working_hours', 'start': HELSINKI.localize(MONDAY.replace(hour=10)), 'end': HELSINKI.localize(MONDAY.replace(hour=11)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},
        
        # 2. Evening event, within the second working slot
        {'uid': 'evening_slot', 'start': HELSINKI.localize(MONDAY.replace(hour=21, minute=30)), 'end': HELSINKI.localize(MONDAY.replace(hour=22)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},
        
        # 3. Event that starts before working hours but ends within them (should be included)
        {'uid': 'starts_before', 'start': HELSINKI.localize(MONDAY.replace(hour=8, minute=30)), 'end': HELSINKI.localize(MONDAY.replace(hour=9, minute=30)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},
        
        # 4. Event that starts within working hours but ends after (should be included)
        {'uid': 'ends_after', 'start': HELSINKI.localize(MONDAY.replace(hour=16, minute=30)), 'end': HELSINKI.localize(MONDAY.replace(hour=17, minute=30)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},

        # --- Events that should BE FILTERED OUT ---
        # 5. Event on a non-working day
        {'uid': 'weekend_event', 'start': HELSINKI.localize(SATURDAY.replace(hour=10)), 'end': HELSINKI.localize(SATURDAY.replace(hour=11)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},
        
        # 6. Event between the two working slots on a Monday
        {'uid': 'between_slots', 'start': HELSINKI.localize(MONDAY.replace(hour=18)), 'end': HELSINKI.localize(MONDAY.replace(hour=19)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},

        # 7. A "free" event (TRANSPARENT) during working hours (should be filtered out first)
        {'uid': 'free_event', 'start': HELSINKI.localize(MONDAY.replace(hour=14)), 'end': HELSINKI.localize(MONDAY.replace(hour=15)), 'transp': 'TRANSPARENT', 'status': 'CONFIRMED'},

        # 8. A duplicate event to test de-duplication
        {'uid': 'inside_working_hours', 'start': HELSINKI.localize(MONDAY.replace(hour=10)), 'end': HELSINKI.localize(MONDAY.replace(hour=11)), 'transp': 'OPAQUE', 'status': 'CONFIRMED'},
    ]
    return events

def test_full_event_processing(sample_raw_events, sample_config):
    """
    Tests the entire chain of processing: de-duplication, busy/free filtering,
    and working hours filtering.
    """
    processed = process_ical_events(sample_raw_events, sample_config)
    
    # We expect 4 events to remain
    assert len(processed) == 4
    
    processed_uids = {e['uid'] for e in processed}
    
    # Check that the correct events were kept
    assert 'inside_working_hours' in processed_uids
    assert 'evening_slot' in processed_uids
    assert 'starts_before' in processed_uids
    assert 'ends_after' in processed_uids
    
    # Check that the correct events were filtered out
    assert 'weekend_event' not in processed_uids
    assert 'between_slots' not in processed_uids
    assert 'free_event' not in processed_uids