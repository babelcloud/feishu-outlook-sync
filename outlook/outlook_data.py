from O365 import Account
from datetime import datetime, timezone, timedelta
from typing import Optional, List, Dict
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from save_handler import SaveHandler

class OutlookHandler:
    def __init__(self, client_id, client_secret, tenant_id):
        self.CLIENT_ID = client_id
        self.CLIENT_SECRET = client_secret
        self.TENANT_ID = tenant_id
        
        # Allow http for development
        import os
        os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
        
        self.save_handler = SaveHandler()
        
        # Initialize account with offline_access scope to get refresh tokens
        self.account = Account((client_id, client_secret),
                             tenant_id=tenant_id,
                             scopes=['offline_access', 'calendar_all'])
        
        # Try to load existing tokens
        self._load_saved_token()

    def _load_saved_token(self):
        """Load and validate both access and refresh tokens."""
        access_token, refresh_token, expiration = self.save_handler.get_outlook_token()
        if access_token and refresh_token:
            self.account.connection.token_backend.token = {
                'token_type': 'Bearer',
                'access_token': access_token,
                'refresh_token': refresh_token,
                'expires_at': expiration
            }

    def _save_tokens(self, token_dict):
        """Save both access and refresh tokens."""
        self.save_handler.set_outlook_token(
            access_token=token_dict.get('access_token'),
            refresh_token=token_dict.get('refresh_token'),
            expires_in=3600  # Standard expiration for access token
        )

    def authenticate(self):
        """Authenticate with proper token refresh handling."""
        try:
            # First check if current token is valid
            if self.account.is_authenticated:
                return True

            # Check if we have a refresh token to use
            token_dict = self.account.connection.token_backend.token
            if token_dict and 'refresh_token' in token_dict:
                try:
                    # Attempt to refresh the token
                    print("Attempting to refresh token...")
                    result = self.account.connection.refresh_token()
                    if result:
                        # Save the new tokens
                        self._save_tokens(self.account.connection.token_backend.token)
                        print("Successfully refreshed token")
                        return True
                except Exception as e:
                    print(f"Token refresh failed: {e}")
                    # Continue to full authentication if refresh fails

            # If we get here, we need a full authentication
            print("\nFull authentication required. Please sign in to your Outlook account in the browser window...")
            result = self.account.authenticate()
            
            if result:
                # Save both access and refresh tokens
                self._save_tokens(self.account.connection.token_backend.token)
                return True
            return False
            
        except Exception as e:
            print(f"Authentication error: {e}")
            return False

    def get_primary_calendar(self):
        try:
            if not self.authenticate():
                return None
                
            schedule = self.account.schedule()
            calendar = schedule.get_default_calendar()
            calendar_id = calendar.calendar_id
            
            self.save_handler.set_outlook_calendar_id(calendar_id)
            return calendar_id
            
        except Exception as e:
            print(f"Error getting calendar: {e}")
            return None

    def create_event(self, calendar_id, event_data):
        try:
            if not self.authenticate():
                return False
                
            schedule = self.account.schedule()
            calendar = schedule.get_default_calendar()
            
            event = calendar.new_event()
            event.subject = event_data.get('summary')
            
            # Convert timestamps
            start_time = datetime.fromtimestamp(int(event_data['start_time']['timestamp']))
            end_time = datetime.fromtimestamp(int(event_data['end_time']['timestamp']))
            
            event.start = start_time
            event.end = end_time
            event.body = event_data.get('description', '')
            event.location = event_data.get('location', '')
            
            success = event.save()
            if success:
                print(f"Successfully created event: {event.subject}")
            return success
            
        except Exception as e:
            print(f"Error creating event: {e}")
            return False

    def _standardize_timestamp(self, dt):
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return str(int(dt.timestamp()))

    def get_future_calendar_events(self, calendar_id=None) -> List[Dict]:
        try:
            if not self.authenticate():
                print("Failed to authenticate with Outlook")
                return []

            schedule = self.account.schedule()
            calendar = schedule.get_default_calendar()
            if not calendar:
                print("Failed to get default calendar")
                return []

            now = datetime.now(timezone.utc)
            end_time = now + timedelta(days=365)
            
            print(f"Fetching events between: {now.isoformat()} and {end_time.isoformat()}")

            try:
                # Debug logging for query
                query = calendar.new_query('start').greater_equal(now)
                query.chain('and').on_attribute('end').less_equal(end_time)
                print(f"Generated query: {query}")
                
                # Force select all needed fields
                query.select('subject', 'start', 'end', 'location', 'body', 'is_cancelled')
                
                # Get events with error tracking
                try:
                    events = list(calendar.get_events(
                        query=query,
                        include_recurring=True,
                        batch=50
                    ))
                    print(f"Raw events retrieved: {len(events)}")
                except Exception as e:
                    print(f"Error during event retrieval: {e}")
                    events = []
                
                formatted_events = []
                seen_events = {}
                
                for event in events:
                    try:
                        # Debug logging for event processing
                        print(f"\nProcessing Outlook event: {event.subject}")
                        print(f"Start time (raw): {event.start}")
                        print(f"Start time (UTC): {event.start.astimezone(timezone.utc)}")
                        
                        event_key = self._create_event_key(event)
                        print(f"Generated event key: {event_key}")
                        
                        if event_key not in seen_events:
                            formatted_event = self._format_event(event)
                            if formatted_event:
                                formatted_events.append(formatted_event)
                                seen_events[event_key] = formatted_event
                                print(f"Processing event: {event.subject} at {event.start}")
                        else:
                            print(f"Found duplicate - New: {event.subject} at {event.start}")
                            print(f"Existing: {seen_events[event_key]['summary']} at {datetime.fromtimestamp(int(seen_events[event_key]['start_time']['timestamp']), tz=timezone.utc)}")
                    except Exception as e:
                        print(f"Error processing individual event: {str(e)}")
                        continue
                
                print(f"Successfully processed {len(formatted_events)} events")
                return formatted_events
                
            except Exception as e:
                print(f"Error fetching events: {str(e)}")
                return []
            
        except Exception as e:
            print(f"Error in get_future_calendar_events: {e}")
            return []

    def _create_event_key(self, event) -> str:
        try:
            # Keep times in UTC for consistent comparison
            start_time = event.start.astimezone(timezone.utc)
            end_time = event.end.astimezone(timezone.utc)
            
            # Add both start and end time to key for better uniqueness
            key = f"{event.subject}|{int(start_time.timestamp())}|{int(end_time.timestamp())}"
            return key
            
        except Exception as e:
            print(f"Error creating event key: {e}")
            return None

    def _format_event(self, event) -> Optional[Dict]:
        try:
            # Ensure times are in UTC
            start_time = event.start.astimezone(timezone.utc)
            end_time = event.end.astimezone(timezone.utc)
                
            # Convert to Unix timestamp
            start_timestamp = int(start_time.timestamp())
            end_timestamp = int(end_time.timestamp())
                
            return {
                'event_id': event.object_id,
                'summary': event.subject,
                'description': event.body or '',
                'start_time': {'timestamp': str(start_timestamp)},
                'end_time': {'timestamp': str(end_timestamp)},
                'location': event.location or '',
                'status': 'confirmed' if not event.is_cancelled else 'cancelled'
            }
        except Exception as e:
            print(f"Error formatting event: {e}")
            return None