import pytest
from datetime import datetime, timedelta
import pytz
from calsync_app.sync_calendar import reconcile_events # Import the function to test

# Define some fixed timestamps for predictable tests
NOW = datetime.now(pytz.UTC)
TOMORROW = NOW + timedelta(days=1)
DAY_AFTER = NOW + timedelta(days=2)

@pytest.fixture
def sample_events():
    """Provides sample data for tests."""
    ical_events = [
        {'uid': 'uid1', 'start': NOW, 'end': NOW + timedelta(hours=1)}, # Unchanged
        {'uid': 'uid2', 'start': TOMORROW, 'end': TOMORROW + timedelta(hours=1)}, # Modified
        {'uid': 'uid4', 'start': DAY_AFTER, 'end': DAY_AFTER + timedelta(hours=1)}, # New
    ]
    outlook_events = [
        {'outlook_id': 'o1', 'uid': 'uid1', 'start': NOW, 'end': NOW + timedelta(hours=1)}, # Unchanged
        {'outlook_id': 'o2', 'uid': 'uid2', 'start': TOMORROW - timedelta(hours=2), 'end': TOMORROW - timedelta(hours=1)}, # Old version
        {'outlook_id': 'o3', 'uid': 'uid3', 'start': DAY_AFTER, 'end': DAY_AFTER + timedelta(hours=1)}, # Deleted
    ]
    return ical_events, outlook_events

def test_reconciliation(sample_events):
    """Tests the main reconciliation logic with a mix of changes."""
    ical_events, outlook_events = sample_events
    
    to_create, to_update, to_delete = reconcile_events(ical_events, outlook_events)

    # Assertions for TO_CREATE
    assert len(to_create) == 1
    assert to_create[0]['uid'] == 'uid4'

    # Assertions for TO_UPDATE
    assert len(to_update) == 1
    assert to_update[0]['uid'] == 'uid2'
    assert to_update[0]['outlook_id'] == 'o2'
    assert to_update[0]['start'] == TOMORROW # Check that it has the new time

    # Assertions for TO_DELETE
    assert len(to_delete) == 1
    assert to_delete[0]['uid'] == 'uid3'
    assert to_delete[0]['outlook_id'] == 'o3'

def test_no_changes():
    """Tests that no actions are generated when calendars are in sync."""
    ical_events = [
        {'uid': 'uid1', 'start': NOW, 'end': NOW + timedelta(hours=1)}
    ]
    outlook_events = [
        {'outlook_id': 'o1', 'uid': 'uid1', 'start': NOW, 'end': NOW + timedelta(hours=1)}
    ]
    to_create, to_update, to_delete = reconcile_events(ical_events, outlook_events)
    assert not to_create and not to_update and not to_delete

def test_all_new():
    """Tests that all events are marked for creation on the first run."""
    ical_events = [
        {'uid': 'uid1', 'start': NOW, 'end': NOW + timedelta(hours=1)}
    ]
    outlook_events = []
    to_create, to_update, to_delete = reconcile_events(ical_events, outlook_events)
    assert len(to_create) == 1
    assert not to_update and not to_delete