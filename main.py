import sys
import time
import requests
from datetime import datetime, timezone, timedelta
from auth_handler import AuthHandler

def get_outlook_events(auth_handler: AuthHandler, calendar_id: str):
    """Get Outlook events for a specific calendar."""
    if not auth_handler.verify_outlook_token():
        print("Failed to verify Outlook token")
        return None

    try:
        schedule = auth_handler.outlook_account.schedule()
        calendar = schedule.get_calendar(calendar_id)
        if not calendar:
            print(f"Failed to get calendar with ID: {calendar_id}")
            return None

        # Get current time in UTC
        now = datetime.now(timezone.utc)
        end_time = now + timedelta(days=365)
        
        print(f"Fetching events between: {now.isoformat()} and {end_time.isoformat()}")

        try:
            query = calendar.new_query('start').greater_equal(now)
            query.chain('and').on_attribute('end').less_equal(end_time)
            query.select('subject', 'start', 'end', 'location', 'body', 'is_cancelled')
            
            events = list(calendar.get_events(
                query=query,
                include_recurring=True,
                batch=50
            ))
            
            formatted_events = []
            seen_events = {}
            
            for event in events:
                try:
                    start_time = event.start.astimezone(timezone.utc)
                    end_time = event.end.astimezone(timezone.utc)
                    
                    event_key = f"{event.subject}|{int(start_time.timestamp())}|{int(end_time.timestamp())}"
                    
                    if event_key not in seen_events:
                        formatted_event = {
                            'event_id': event.object_id,
                            'summary': event.subject,
                            'description': event.body or '',
                            'start_time': {'timestamp': str(int(start_time.timestamp()))},
                            'end_time': {'timestamp': str(int(end_time.timestamp()))},
                            'location': event.location or '',
                            'status': 'confirmed' if not event.is_cancelled else 'cancelled'
                        }
                        formatted_events.append(formatted_event)
                        seen_events[event_key] = formatted_event
                except Exception as e:
                    print(f"Error processing individual event: {e}")
                    continue
            
            return formatted_events
            
        except Exception as e:
            print(f"Error during event retrieval: {e}")
            return None
            
    except Exception as e:
        print(f"Error fetching Outlook events: {e}")
        return None


def get_feishu_events(auth_handler: AuthHandler, calendar_id: str):
    """Get Feishu events with proper token verification."""
    if not auth_handler.verify_feishu_tokens():
        print("Failed to verify Feishu tokens")
        return None

    try:
        user_token = auth_handler.get_feishu_user_token()
        if not user_token:
            print("Failed to get Feishu user token")
            return None

        payload = {'start_time': str(int(time.time()))}
        response = requests.get(
            f'https://open.feishu.cn/open-apis/calendar/v4/calendars/{calendar_id}/events',
            headers={'Authorization': f'Bearer {user_token}'},
            params=payload
        )

        if response.status_code != 200:
            print(f"Failed to get Feishu events: {response.status_code}")
            return None

        response_data = response.json()
        return response_data.get('data', {}).get('items', [])

    except Exception as e:
        print(f"Error fetching Feishu events: {e}")
        return None

def filter_future_events(events):
    """Filter out past events from Feishu events list."""
    now = int(time.time())
    future_events = []
    
    for event in events:
        try:
            start_timestamp = int(float(event['start_time']['timestamp']))
            if start_timestamp >= now:
                future_events.append(event)
        except (KeyError, ValueError) as e:
            print(f"Error processing event timestamp: {e}")
            continue
            
    return future_events

def sync_calendar_events(auth_handler: AuthHandler, feishu_events, outlook_events, outlook_calendar_id: str):
    """Sync events from Feishu to a specific Outlook calendar."""
    synced_count = 0
    skipped_count = 0
    failed_count = 0

    # Create lookup map for existing events
    existing_events = {}
    for event in (outlook_events or []):
        try:
            key = (
                event.get('summary', ''),
                int(float(event['start_time']['timestamp']))
            )
            existing_events[key] = event['event_id']
        except Exception as e:
            print(f"Error processing existing event: {e}")

    # Get specific Outlook calendar
    schedule = auth_handler.outlook_account.schedule()
    calendar = schedule.get_calendar(outlook_calendar_id)
    if not calendar:
        print(f"Failed to get calendar with ID: {outlook_calendar_id}")
        return 0, 0, 0
        
    for event in (feishu_events or []):
        try:
            # Skip cancelled events
            if event.get('status') == 'cancelled':
                print("Skipping cancelled event")
                continue

            # Validate required fields
            summary = event.get('summary')
            if not summary:
                print("Skipping event with no summary")
                failed_count += 1
                continue

            start_timestamp = event.get('start_time', {}).get('timestamp')
            end_timestamp = event.get('end_time', {}).get('timestamp')
            if not start_timestamp or not end_timestamp:
                print("Skipping event with invalid timestamps")
                failed_count += 1
                continue

            event_key = (summary, int(float(start_timestamp)))
            event_start = datetime.fromtimestamp(int(float(start_timestamp)), tz=timezone.utc)
            
            if event_key in existing_events:
                print(f"\nSkipping existing event: {summary}")
                print(f"Start: {event_start}")
                skipped_count += 1
                continue

            print(f"\nNew event found: {summary}")
            print(f"Start: {event_start}")
            
            new_event = calendar.new_event()
            
            # Set required fields
            new_event.subject = summary.strip()
            
            # Convert timestamps
            start_time = datetime.fromtimestamp(int(float(start_timestamp)))
            end_time = datetime.fromtimestamp(int(float(end_timestamp)))
            
            new_event.start = start_time
            new_event.end = end_time

            # Build event body
            body_parts = []
            
            # Add description if present
            description = event.get('description')
            if description:
                body_parts.append(description.strip())

            # Add meeting URL if present
            vchat = event.get('vchat', {})
            if vchat and vchat.get('meeting_url'):
                body_parts.append(f"Meeting URL: {vchat['meeting_url']}")

            # Set body if we have any content
            if body_parts:
                new_event.body = "\n\n".join(body_parts)
            
            # Handle location - extract just the location name
            location = event.get('location', {})
            if isinstance(location, dict) and location.get('name'):
                new_event.location = location['name']
            elif isinstance(location, str) and location:
                new_event.location = location

            # Save the event with detailed error logging
            try:
                print(f"Saving event:")
                print(f"  Subject: {new_event.subject}")
                print(f"  Start: {new_event.start}")
                print(f"  End: {new_event.end}")
                print(f"  Body: {new_event.body}")
                if hasattr(new_event, 'location'):
                    print(f"  Location: {new_event.location}")

                if new_event.save():
                    print(f"Successfully created event: {summary}")
                    print(f"Successfully synced: {summary}")
                    synced_count += 1
                else:
                    print(f"Failed to sync: {summary}")
                    failed_count += 1
            except Exception as e:
                print(f"Error saving event: {str(e)}")
                failed_count += 1
                    
        except Exception as e:
            print(f"Error processing event for sync: {e}")
            failed_count += 1

    return synced_count, skipped_count, failed_count

def sync_calendars(auth_handler: AuthHandler):
    """Main sync function that handles all calendar pairs."""
    # Verify Feishu tokens first
    if not auth_handler.verify_feishu_tokens():
        print("Feishu token verification failed")
        return False

    # Verify Outlook token
    if not auth_handler.verify_outlook_token():
        print("Outlook token verification failed")
        return False

    try:
        total_synced = 0
        total_skipped = 0
        total_failed = 0

        for pair in auth_handler.calendar_pairs:
            feishu_id = pair['feishu']['id']
            feishu_name = pair['feishu']['name']
            outlook_id = pair['outlook']['id']
            outlook_name = pair['outlook']['name']

            print(f"\nProcessing calendar pair:")
            print(f"Feishu: {feishu_name}")
            print(f"Outlook: {outlook_name}")
            
            # Get Feishu events
            feishu_events = get_feishu_events(auth_handler, feishu_id)
            if feishu_events is None:
                print(f"Failed to fetch events from Feishu calendar: {feishu_name}")
                continue

            # Get Outlook events for this specific calendar
            outlook_events = get_outlook_events(auth_handler, outlook_id)
            if outlook_events is None:
                print(f"Failed to fetch events from Outlook calendar: {outlook_name}")
                continue

            # Filter future events
            future_events = filter_future_events(feishu_events)
            print(f"Found {len(future_events)} future events in {feishu_name}")

            # Sync events to specific Outlook calendar
            synced, skipped, failed = sync_calendar_events(
                auth_handler=auth_handler,
                feishu_events=future_events,
                outlook_events=outlook_events,
                outlook_calendar_id=outlook_id
            )
            
            total_synced += synced
            total_skipped += skipped
            total_failed += failed

        print(f"\nSync Summary:")
        print(f"- Events synced: {total_synced}")
        print(f"- Events skipped: {total_skipped}")
        print(f"- Events failed: {total_failed}")

        return True
    
    except Exception as e:
        print(f"Error during sync: {e}")
        return False


def main():
    auth_handler = AuthHandler()
    
    # Verify initial setup
    if not auth_handler.is_fully_configured():
        print("Please run auth_handler.py first to setup authentication")
        sys.exit(1)

    print("Starting sync process...")
    
    try:
        # Do initial sync
        if not sync_calendars(auth_handler):
            print("\nInitial sync failed")
            sys.exit(1)
            
        print("\nInitial sync completed successfully")
        print("Starting continuous sync...")
        
        # Start continuous sync
        while True:
            print("\nWaiting 5 minutes before next sync...")
            time.sleep(300)  # Wait 5 minutes
            
            if not sync_calendars(auth_handler):
                print("\nSync failed, will retry in next cycle")
            else:
                print("Sync completed successfully")
            
    except KeyboardInterrupt:
        print("\nSync process stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"\nError during sync: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()