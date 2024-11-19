from feishu.feishu_data import DataHandler as FeishuHandler
from outlook.outlook_data import OutlookHandler
from save_handler import SaveHandler
import sys
import time
from datetime import datetime, timezone

tkmanager = SaveHandler()

def verify_tokens():
    """Verify and refresh tokens for both services."""
    # Feishu verification
    if not tkmanager.get_feishu_app_token() or not tkmanager.is_feishu_app_token_valid():
        app_tk = feishu_handler.obtain_app_access_token()
    if not tkmanager.get_feishu_user_token() or not tkmanager.is_feishu_user_token_valid():
        oauth_code = feishu_handler.obtain_oauth_code()
        user_tk = feishu_handler.obtain_user_access_token(code=oauth_code)
    if not tkmanager.get_feishu_calendar_id():
        calendar_id = feishu_handler.get_primary_calendar(access_token=user_tk)

    # Outlook verification
    if not tkmanager.is_outlook_authenticated():
        outlook_handler.authenticate()
    if not tkmanager.get_outlook_calendar_id():
        outlook_handler.get_primary_calendar()

def sync_calendars():
    """Sync calendar events between Feishu and Outlook."""
    verify_tokens()

    now = datetime.now(timezone.utc)
    print("\nFetching Feishu events...")
    feishu_events = feishu_handler.get_future_calendar_events(
        calendar_id=tkmanager.get_feishu_calendar_id(),
        access_token=tkmanager.get_feishu_user_token()
    )
    
    # Filter out past events from Feishu
    future_feishu_events = []
    for event in (feishu_events or []):
        try:
            start_timestamp = int(float(event['start_time']['timestamp']))
            event_start = datetime.fromtimestamp(start_timestamp, tz=timezone.utc)
            if event_start >= now:
                future_feishu_events.append(event)
            else:
                print(f"Skipping past event: {event['summary']} at {event_start}")
        except Exception as e:
            print(f"Error processing Feishu event timestamp: {e}")
            continue

    print(f"Found {len(feishu_events or [])} total Feishu events")
    print(f"Found {len(future_feishu_events)} future Feishu events")

    print("\nFetching Outlook events...")
    outlook_events = outlook_handler.get_future_calendar_events()
    print(f"Found {len(outlook_events or [])} Outlook events")
    
    # Create lookup map of existing events
    existing_events = {}
    for event in (outlook_events or []):
        try:
            # Create a unique key for each event
            key = (
                event['summary'],
                int(float(event['start_time']['timestamp']))
            )
            existing_events[key] = event['event_id']
            print(f"Existing event found: {event['summary']} at {datetime.fromtimestamp(int(float(event['start_time']['timestamp'])), tz=timezone.utc)}")
        except Exception as e:
            print(f"Error processing existing event: {e}")

    print("\nStarting sync process...")
    synced_count = 0
    skipped_count = 0
    failed_count = 0

    # Sync new events
    for event in (future_feishu_events or []):
        try:
            # Create the same format key for comparison
            event_key = (
                event['summary'],
                int(float(event['start_time']['timestamp']))
            )
            
            event_start = datetime.fromtimestamp(int(float(event['start_time']['timestamp'])), tz=timezone.utc)
            
            if event_key in existing_events:
                print(f"\nSkipping existing event: {event['summary']}")
                print(f"Start: {event_start}")
                skipped_count += 1
            else:
                print(f"\nNew event found: {event['summary']}")
                print(f"Start: {event_start}")
                if outlook_handler.create_event(None, event):
                    print(f"Successfully created event: {event['summary']}")
                    print(f"Successfully synced: {event['summary']}")
                    synced_count += 1
                else:
                    print(f"Failed to sync: {event['summary']}")
                    failed_count += 1
        except Exception as e:
            print(f"Error processing event for sync: {e}")
            failed_count += 1

    print(f"\nSync Summary:")
    print(f"- Total Feishu events: {len(feishu_events or [])}")
    print(f"- Future Feishu events: {len(future_feishu_events)}")
    print(f"- Total Outlook events: {len(outlook_events or [])}")
    print(f"- Events synced: {synced_count}")
    print(f"- Events skipped: {skipped_count}")
    print(f"- Events failed: {failed_count}")

def main():
    # Initialize SaveHandler
    tkmanager = SaveHandler()
    
    # Get Feishu credentials
    feishu_id, feishu_secret = tkmanager.get_feishu_app_info()
    if not feishu_id or not feishu_secret:
        feishu_id = input("Enter your Feishu app ID: ")
        feishu_secret = input("Enter your Feishu app secret: ")
        tkmanager.set_feishu_app_info(feishu_id, feishu_secret)

    # Get Outlook credentials
    outlook_id, outlook_secret, tenant_id = tkmanager.get_outlook_app_info()
    if not outlook_id or not outlook_secret or not tenant_id:
        outlook_id = input("Enter your Outlook client ID: ")
        outlook_secret = input("Enter your Outlook client secret: ")
        tenant_id = input("Enter your Azure AD tenant ID: ")
        tkmanager.set_outlook_app_info(outlook_id, outlook_secret, tenant_id)

    # Initialize handlers
    global feishu_handler, outlook_handler
    feishu_handler = FeishuHandler(app_id=feishu_id, app_secret=feishu_secret)
    outlook_handler = OutlookHandler(client_id=outlook_id, client_secret=outlook_secret, tenant_id=tenant_id)

    # Initial authentication
    if not outlook_handler.authenticate():
        print("Failed to authenticate with Outlook")
        sys.exit(1)

    tkmanager.set_outlook_authenticated(True)

    print("Starting sync process...")
    
    try:
        # Do initial sync
        sync_calendars()
        print("\nInitial sync completed successfully")
        print("Starting continuous sync...")
        
        # Start continuous sync
        while True:
            print("\nWaiting 5 minutes before next sync...")
            time.sleep(300)  # Wait 5 minutes before next sync
            sync_calendars()
            print("Sync completed successfully")
            
    except KeyboardInterrupt:
        print("\nSync process stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"\nError during sync: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()