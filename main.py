import sys
import os
import time
import requests
from datetime import datetime, timezone, timedelta
from auth_handler import AuthHandler
from db_handler import DBHandler
from typing import Optional, Tuple
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Global database handler instance
db_handler = None

def get_outlook_events(auth_handler: AuthHandler, calendar_id: str):
    """Get Outlook events with proper query handling."""
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
            # Create query with proper date filtering
            query = calendar.new_query('start').greater_equal(now)
            query.chain('and').on_attribute('end').less_equal(end_time)
            
            # Force select all needed fields
            query.select('subject', 'start', 'end', 'location', 'body', 'is_cancelled')
            
            print(f"Generated query: {query}")
            
            events = list(calendar.get_events(
                query=query,
                include_recurring=True,
                batch=50
            ))
            print(f"Raw events retrieved: {len(events)}")
            
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
                        print(f"Processing event: {event.subject} at {start_time}")
                    else:
                        print(f"Found duplicate - New: {event.subject} at {start_time}")
                        print(f"Existing: {seen_events[event_key]['summary']} at "
                              f"{datetime.fromtimestamp(int(seen_events[event_key]['start_time']['timestamp']), tz=timezone.utc)}")
                        
                except Exception as e:
                    print(f"Error processing individual event: {e}")
                    continue
            
            print(f"Successfully processed {len(formatted_events)} events")
            return formatted_events
            
        except Exception as e:
            print(f"Error during event retrieval: {e}")
            return None
            
    except Exception as e:
        print(f"Error fetching Outlook events: {e}")
        return None

def get_feishu_events(auth_handler: AuthHandler, calendar_id: str):
    """Get Feishu events with proper token verification, including deleted events."""
    if not auth_handler.verify_feishu_tokens():
        print("Failed to verify Feishu tokens")
        return None

    try:
        user_token = auth_handler.get_feishu_user_token()
        if not user_token:
            print("Failed to get Feishu user token")
            return None

        # Get current time for pagination start
        start_time = str(int(time.time()))
        all_events = []
        page_token = None

        while True:
            # Prepare request parameters
            params = {
                'start_time': start_time,
                'page_size': 100  # Maximum allowed by API
            }
            if page_token:
                params['page_token'] = page_token

            response = requests.get(
                f'https://open.feishu.cn/open-apis/calendar/v4/calendars/{calendar_id}/events',
                headers={'Authorization': f'Bearer {user_token}'},
                params=params
            )

            if response.status_code != 200:
                print(f"Failed to get Feishu events: {response.status_code}")
                return None

            response_data = response.json()
            items = response_data.get('data', {}).get('items', [])
            all_events.extend(items)

            # Check for more pages
            page_token = response_data.get('data', {}).get('page_token')
            if not page_token:
                break

        print(f"Retrieved {len(all_events)} events from Feishu")
        return all_events

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

def sync_calendar_events(auth_handler: AuthHandler, feishu_events, outlook_events, outlook_calendar_id: str, event_sync_id: int = -1) -> Tuple[int, int, int, int]:
    """Sync events (including deletions) from Feishu to specific Outlook calendar."""
    synced_count = 0
    skipped_count = 0
    failed_count = 0
    deleted_count = 0
    
    # Get current timestamp for filtering
    current_timestamp = int(time.time())

    # Create lookup maps for existing events
    existing_events = {}
    feishu_event_map = {}
    
    # Map Outlook events by summary and start time
    for event in (outlook_events or []):
        try:
            start_timestamp = int(float(event['start_time']['timestamp']))
            
            # Only consider future events for syncing
            if start_timestamp >= current_timestamp:
                key = (
                    event.get('summary', ''),
                    start_timestamp
                )
                existing_events[key] = {
                    'id': event['event_id'],
                    'status': event.get('status', 'confirmed'),
                    'start_timestamp': start_timestamp
                }
                print(f"Future existing event found: {event.get('summary', '')} at {datetime.fromtimestamp(start_timestamp, tz=timezone.utc)}")
            else:
                print(f"Skipping past event from consideration: {event.get('summary', '')} at {datetime.fromtimestamp(start_timestamp, tz=timezone.utc)}")
        except Exception as e:
            print(f"Error processing existing event: {e}")

    # Map Feishu events by summary and start time
    for event in (feishu_events or []):
        try:
            summary = event.get('summary')
            start_timestamp = event.get('start_time', {}).get('timestamp')
            if summary and start_timestamp:
                start_timestamp = int(float(start_timestamp))
                if start_timestamp >= current_timestamp:
                    key = (summary, start_timestamp)
                    feishu_event_map[key] = event.get('status', 'confirmed')
        except Exception as e:
            print(f"Error mapping Feishu event: {e}")

    # Get specific calendar for syncing
    schedule = auth_handler.outlook_account.schedule()
    calendar = schedule.get_calendar(outlook_calendar_id)
    if not calendar:
        print(f"Failed to get calendar with ID: {outlook_calendar_id}")
        return 0, 0, 0, 0

    # Process deletions only for future events
    for key, outlook_data in existing_events.items():
        summary, start_timestamp = key
        
        # Double check timestamp before deletion
        if start_timestamp >= current_timestamp:
            # If event exists in Outlook but not in Feishu, or is cancelled in Feishu
            if key not in feishu_event_map or feishu_event_map[key] == 'cancelled':
                try:
                    print(f"\nProcessing deletion for future event: {summary} at {datetime.fromtimestamp(start_timestamp, tz=timezone.utc)}")
                    event = calendar.get_event(outlook_data['id'])
                    if event:
                        start_time = datetime.fromtimestamp(start_timestamp, tz=timezone.utc)
                        
                        # Log the deletion attempt
                        if db_handler and event_sync_id > 0:
                            db_handler.log_event_action(
                                event_sync_id=event_sync_id,
                                event_title=summary,
                                event_start_time=start_time,
                                event_end_time=start_time + timedelta(hours=1),  # Approximation if end time unavailable
                                action="delete",
                                status="started"
                            )
                            
                        if event.delete():
                            print(f"Successfully deleted event: {summary}")
                            deleted_count += 1
                            
                            # Log the deletion success
                            if db_handler and event_sync_id > 0:
                                db_handler.log_event_action(
                                    event_sync_id=event_sync_id,
                                    event_title=summary,
                                    event_start_time=start_time,
                                    event_end_time=start_time + timedelta(hours=1),  # Approximation if end time unavailable
                                    action="delete",
                                    status="success"
                                )
                        else:
                            print(f"Failed to delete event: {summary}")
                            failed_count += 1
                            
                            # Log the deletion failure
                            if db_handler and event_sync_id > 0:
                                db_handler.log_event_action(
                                    event_sync_id=event_sync_id,
                                    event_title=summary,
                                    event_start_time=start_time,
                                    event_end_time=start_time + timedelta(hours=1),  # Approximation if end time unavailable
                                    action="delete",
                                    status="failed",
                                    error_message="Delete operation returned failure"
                                )
                except Exception as e:
                    print(f"Error deleting event: {e}")
                    failed_count += 1
                    
                    # Log the deletion error
                    if db_handler and event_sync_id > 0:
                        start_time = datetime.fromtimestamp(start_timestamp, tz=timezone.utc)
                        db_handler.log_event_action(
                            event_sync_id=event_sync_id,
                            event_title=summary,
                            event_start_time=start_time,
                            event_end_time=start_time + timedelta(hours=1),  # Approximation
                            action="delete",
                            status="error",
                            error_message=str(e)
                        )

    # Process regular events (new and updates)
    for event in (feishu_events or []):
        try:
            # Skip already cancelled events
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

            # Convert to integer timestamp
            start_timestamp = int(float(start_timestamp))
            end_timestamp = int(float(end_timestamp))
            
            # Skip past events
            if start_timestamp < current_timestamp:
                print(f"Skipping past event: {summary}")
                continue

            event_key = (summary, start_timestamp)
            event_start = datetime.fromtimestamp(start_timestamp, tz=timezone.utc)
            event_end = datetime.fromtimestamp(end_timestamp, tz=timezone.utc)
            
            if event_key in existing_events:
                print(f"\nSkipping existing event: {summary}")
                print(f"Start: {event_start}")
                skipped_count += 1
                continue

            print(f"\nNew event found: {summary}")
            print(f"Start: {event_start}")
            
            # Log the creation attempt
            if db_handler and event_sync_id > 0:
                db_handler.log_event_action(
                    event_sync_id=event_sync_id,
                    event_title=summary,
                    event_start_time=event_start,
                    event_end_time=event_end,
                    action="create",
                    status="started"
                )
            
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
                    
                    # Log the creation success
                    if db_handler and event_sync_id > 0:
                        db_handler.log_event_action(
                            event_sync_id=event_sync_id,
                            event_title=summary,
                            event_start_time=event_start,
                            event_end_time=event_end,
                            action="create",
                            status="success"
                        )
                else:
                    print(f"Failed to sync: {summary}")
                    failed_count += 1
                    
                    # Log the creation failure
                    if db_handler and event_sync_id > 0:
                        db_handler.log_event_action(
                            event_sync_id=event_sync_id,
                            event_title=summary,
                            event_start_time=event_start,
                            event_end_time=event_end,
                            action="create",
                            status="failed",
                            error_message="Save operation returned failure"
                        )
            except Exception as e:
                print(f"Error saving event: {str(e)}")
                failed_count += 1
                
                # Log the creation error
                if db_handler and event_sync_id > 0:
                    db_handler.log_event_action(
                        event_sync_id=event_sync_id,
                        event_title=summary,
                        event_start_time=event_start,
                        event_end_time=event_end,
                        action="create",
                        status="error",
                        error_message=str(e)
                    )
                    
        except Exception as e:
            print(f"Error processing event for sync: {e}")
            failed_count += 1

    print(f"\nDeletion Summary for calendar {outlook_calendar_id}:")
    print(f"- Events deleted: {deleted_count}")
    
    return synced_count, skipped_count, failed_count, deleted_count

def sync_calendars(auth_handler: AuthHandler) -> bool:
    """Main sync function that handles all calendar pairs."""
    global db_handler
    
    if not auth_handler.verify_feishu_tokens():
        print("Feishu token verification failed")
        return False

    if not auth_handler.verify_outlook_token():
        print("Outlook token verification failed")
        return False

    try:
        total_synced = 0
        total_skipped = 0
        total_failed = 0
        total_deleted = 0
        
        # Start sync logging in database
        sync_id = -1
        if db_handler:
            sync_id = db_handler.log_sync_start(auth_handler.yaml_file)

        for pair in auth_handler.calendar_pairs:
            feishu_id = pair['feishu']['id']
            feishu_name = pair['feishu']['name']
            outlook_id = pair['outlook']['id']
            outlook_name = pair['outlook']['name']
            calendar_pair_name = f"{feishu_name} -> {outlook_name}"

            print(f"\nProcessing calendar pair:")
            print(f"Feishu: {feishu_name}")
            print(f"Outlook: {outlook_name}")
            
            # Initialize calendar sync entry in database
            event_sync_id = -1
            if db_handler and sync_id > 0:
                event_sync_id = db_handler.log_calendar_sync(
                    sync_id=sync_id,
                    calendar_pair=calendar_pair_name,
                    feishu_calendar_id=feishu_id,
                    feishu_calendar_name=feishu_name,
                    outlook_calendar_id=outlook_id,
                    outlook_calendar_name=outlook_name,
                    events_synced=0,
                    events_skipped=0,
                    events_failed=0,
                    events_deleted=0
                )
            
            # Get Outlook events for this specific calendar
            outlook_events = get_outlook_events(auth_handler, outlook_id)
            if outlook_events is None:
                print(f"Failed to fetch Outlook events for calendar: {outlook_name}")
                continue
                
            print(f"Found {len(outlook_events or [])} Outlook events in {outlook_name}")
            
            # Get Feishu events
            feishu_events = get_feishu_events(auth_handler, feishu_id)
            if feishu_events is None:
                print(f"Failed to fetch Feishu events for calendar: {feishu_name}")
                continue

            # Filter future events but keep deleted ones
            future_events = filter_future_events(feishu_events)
            print(f"Found {len(future_events)} future events in {feishu_name}")

            # Sync events including deletions
            synced, skipped, failed, deleted = sync_calendar_events(
                auth_handler,
                future_events,
                outlook_events,
                outlook_id,
                event_sync_id
            )
            
            total_synced += synced
            total_skipped += skipped
            total_failed += failed
            total_deleted += deleted
            
            # Update the calendar sync entry with final counts
            if db_handler and event_sync_id > 0:
                db_handler.log_calendar_sync(
                    sync_id=sync_id,
                    calendar_pair=calendar_pair_name,
                    feishu_calendar_id=feishu_id,
                    feishu_calendar_name=feishu_name,
                    outlook_calendar_id=outlook_id,
                    outlook_calendar_name=outlook_name,
                    events_synced=synced,
                    events_skipped=skipped,
                    events_failed=failed,
                    events_deleted=deleted
                )

        print(f"\nSync Summary:")
        print(f"- Events synced: {total_synced}")
        print(f"- Events skipped: {total_skipped}")
        print(f"- Events failed: {total_failed}")
        print(f"- Events deleted: {total_deleted}")
        
        # Update the sync log with final status
        if db_handler and sync_id > 0:
            message = f"Sync completed: {total_synced} synced, {total_skipped} skipped, {total_failed} failed, {total_deleted} deleted"
            db_handler.log_sync_end(sync_id, True, message)

        return True
    
    except Exception as e:
        print(f"Error during sync: {e}")
        
        # Log the error in database
        if db_handler and sync_id > 0:
            db_handler.log_sync_end(sync_id, False, f"Error during sync: {e}")
            
        return False

def run_sync(config_path: str = 'tokens.yaml') -> bool:
    """Run sync process with specified config file."""
    global db_handler
    
    try:
        auth_handler = AuthHandler(yaml_file=config_path)
        
        # Verify initial setup
        if not auth_handler.is_fully_configured():
            print(f"Configuration incomplete for {config_path}")
            return False

        print(f"\nStarting sync process for {config_path}...")
        
        # Do initial sync
        if not sync_calendars(auth_handler):
            print("\nInitial sync failed")
            return False
            
        return True

    except Exception as e:
        print(f"\nError during sync for {config_path}: {e}")
        return False

def run_continuous_sync(config_path: str = 'tokens.yaml', interval: int = 300) -> None:
    """Run continuous sync process with specified interval."""
    global db_handler
    
    try:
        # Initialize database connection if not already done
        if db_handler is None:
            db_handler = DBHandler(
                host=os.environ.get('PGHOST', 'ep-red-brook-aawc47cf-pooler.westus3.azure.neon.tech'),
                database=os.environ.get('PGDATABASE', 'CalendarSync'),
                user=os.environ.get('PGUSER', 'CalendarSync_owner'),
                password=os.environ.get('PGPASSWORD', 'npg_Had49rjAFUVg')
            )
            if not db_handler.connect():
                print("Failed to connect to database, proceeding without database logging")
                db_handler = None
        
        print(f"Starting continuous sync for {config_path}")
        while True:
            success = run_sync(config_path)
            if not success:
                print(f"\nSync failed for {config_path}, will retry in next cycle")
            else:
                print(f"Sync completed successfully for {config_path}")
                
            print(f"\nWaiting {interval} seconds before next sync...")
            time.sleep(interval)
            
    except KeyboardInterrupt:
        print("\nSync process stopped by user")
        
        # Close database connection
        if db_handler:
            db_handler.close()
            
        sys.exit(0)
    except Exception as e:
        print(f"\nError during continuous sync for {config_path}: {e}")
        
        # Close database connection
        if db_handler:
            db_handler.close()
            
        sys.exit(1)

if __name__ == "__main__":
    # Initialize database connection
    db_handler = DBHandler(
        host=os.environ.get('PGHOST', 'ep-red-brook-aawc47cf-pooler.westus3.azure.neon.tech'),
        database=os.environ.get('PGDATABASE', 'CalendarSync'),
        user=os.environ.get('PGUSER', 'CalendarSync_owner'),
        password=os.environ.get('PGPASSWORD', 'npg_Had49rjAFUVg')
    )
    
    if not db_handler.connect():
        print("Failed to connect to database, proceeding without database logging")
        db_handler = None
    
    auth_handler = AuthHandler()
    
    # Verify initial setup
    if not auth_handler.is_fully_configured():
        print("Please run auth_handler.py first to setup authentication")
        
        # Close database connection
        if db_handler:
            db_handler.close()
            
        sys.exit(1)

    print("Starting sync process...")
    
    try:
        # Run continuous sync with default config
        run_continuous_sync()
            
    except KeyboardInterrupt:
        print("\nSync process stopped by user")
        
        # Close database connection
        if db_handler:
            db_handler.close()
            
        sys.exit(0)
    except Exception as e:
        print(f"\nError during sync: {e}")
        
        # Close database connection
        if db_handler:
            db_handler.close()
            
        sys.exit(1)