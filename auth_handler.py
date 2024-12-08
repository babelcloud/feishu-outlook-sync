import json
import os
import time
import yaml
import requests
import lark_oapi as lark
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple
from O365 import Account
from lark_oapi.api.authen.v1 import *
from lark_oapi.api.auth.v3 import *

# Import our OAuth implementations
from feishu_oauth import FeishuOAuth

class AuthHandler:
    def __init__(self, yaml_file: str = 'tokens.yaml'):
        self.yaml_file = yaml_file
        self.config = self._load_config()
        self._setup_clients()

    def _setup_clients(self):
        """Setup API clients if credentials exist."""
        feishu_id, feishu_secret = self.get_feishu_app_info()
        if feishu_id and feishu_secret:
            self.feishu_client = lark.Client.builder() \
                .app_id(feishu_id) \
                .app_secret(feishu_secret) \
                .enable_set_token(True) \
                .log_level(lark.LogLevel.DEBUG) \
                .build()
            
            # Initialize FeishuOAuth
            self.feishu_oauth = FeishuOAuth(feishu_id, feishu_secret)
        
        outlook_id, outlook_secret, tenant_id = self.get_outlook_app_info()
        if outlook_id and outlook_secret:
            self.outlook_account = Account(
                (outlook_id, outlook_secret),
                tenant_id=tenant_id,
                scopes=['offline_access', 'calendar_all']  # Include offline_access for refresh tokens
            )
            
            # Load existing token if available
            self._load_outlook_token()

    def _load_config(self) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(self.yaml_file, 'r') as file:
                return yaml.safe_load(file) or self._get_default_config()
        except FileNotFoundError:
            return self._get_default_config()

    def _get_default_config(self) -> Dict:
        """Get default configuration structure."""
        return {
            'feishu': {
                'app_info': {
                    'app_id': None,
                    'app_secret': None
                },
                'tokens': {
                    'app_access_token': {
                        'token': None,
                        'expiration_time': None
                    },
                    'user_access_token': {
                        'token': None,
                        'refresh_token': None,
                        'expiration_time': None
                    }
                }
            },
            'outlook': {
                'app_info': {
                    'client_id': None,
                    'client_secret': None,
                    'tenant_id': None
                },
                'tokens': {
                    'access_token': None,
                    'refresh_token': None,
                    'expiration_time': None
                },
                'authenticated': False
            },
            'calendar_pairs': []  # Will store pairs of calendars
        }


    def _save_config(self) -> None:
        """Save configuration to YAML file."""
        with open(self.yaml_file, 'w') as file:
            yaml.dump(self.config, file, default_flow_style=False)

    # Feishu Token Management
    def refresh_feishu_app_token(self) -> bool:
        """Refresh Feishu app token."""
        request = InternalAppAccessTokenRequest.builder() \
            .request_body(InternalAppAccessTokenRequestBody.builder()
                       .app_id(self.get_feishu_app_info()[0])
                       .app_secret(self.get_feishu_app_info()[1])
                       .build()) \
            .build()

        response = self.feishu_client.auth.v3.app_access_token.internal(request)
        if not response.success():
            return False

        response_data = json.loads(response.raw.content.decode('utf-8'))
        token = response_data.get('app_access_token')
        expire = response_data.get('expire')
        
        if token and expire:
            self.set_feishu_app_token(token, expire)
            return True
        return False

    def refresh_feishu_user_token(self) -> bool:
        """Refresh Feishu user token using refresh token."""
        try:
            refresh_token = self.get_feishu_refresh_token()
            if not refresh_token:
                print("No refresh token available")
                return False

            # Get app credentials
            app_id, app_secret = self.get_feishu_app_info()
            if not app_id or not app_secret:
                print("Missing app credentials")
                return False

            # Prepare request
            url = 'https://open.feishu.cn/open-apis/authen/v2/oauth/token'
            headers = {
                'Content-Type': 'application/json'
            }
            payload = {
                'grant_type': 'refresh_token',
                'client_id': app_id,
                'client_secret': app_secret,
                'refresh_token': refresh_token
            }

            print("\nDebug Info:")
            print(f"URL: {url}")
            print(f"Headers: {headers}")
            print(f"Payload: {json.dumps({**payload, 'client_secret': '***', 'refresh_token': '***'}, indent=2)}")

            # Make request to refresh token
            response = requests.post(url, headers=headers, json=payload)
            
            if response.status_code != 200:
                print(f"\nError Response Status: {response.status_code}")
                try:
                    error_data = response.json()
                    print(f"Error Response Body: {json.dumps(error_data, indent=2)}")
                except:
                    print(f"Error Response Text: {response.text}")
                return False

            # Parse response
            response_data = response.json()
            if response_data.get('code') != 0:
                print(f"Error in response: {json.dumps(response_data, indent=2)}")
                return False

            # Extract token data
            access_token = response_data.get('access_token')
            new_refresh_token = response_data.get('refresh_token')
            access_token_expires_in = response_data.get('expires_in')
            refresh_token_expires_in = response_data.get('refresh_token_expires_in')

            # Use existing set_feishu_user_token method with updated structure
            self.set_feishu_user_token(
                token=access_token,
                refresh_token=new_refresh_token,
                expires_in=access_token_expires_in,
                refresh_expires_in=refresh_token_expires_in
            )

            print("Successfully refreshed Feishu tokens")
            return True

        except Exception as e:
            print(f"Error refreshing user token: {e}")
            return False

    def verify_feishu_tokens(self) -> bool:
        """Verify and refresh Feishu tokens if needed."""
        try:
            app_valid = self.is_feishu_app_token_valid()
            user_valid = self.is_feishu_user_token_valid()

            # Try to refresh app token if needed
            if not app_valid:
                print("Refreshing Feishu app token...")
                app_valid = self.refresh_feishu_app_token()
                if not app_valid:
                    print("Failed to refresh app token")
                    return False

            # Try to refresh user token if needed
            if not user_valid:
                refresh_token = self.get_feishu_refresh_token()
                if refresh_token:
                    print("Attempting to refresh Feishu user token...")
                    if self.refresh_feishu_user_token():
                        print("Successfully refreshed Feishu user token")
                        user_valid = True
                    else:
                        print("Failed to refresh user token - clearing stored tokens")
                        # Clear invalid tokens
                        self.clear_feishu_user_tokens()
                else:
                    print("No valid refresh token available for Feishu")

                # Only if refresh token doesn't exist or refresh failed, do full reauth
                if not user_valid:
                    print("\nFull Feishu authentication required...")
                    oauth_code = self.feishu_oauth.obtain_oauth_code()
                    if oauth_code:
                        user_valid = self.get_feishu_user_token_from_code(oauth_code)
                    else:
                        print("Failed to obtain Feishu OAuth code")
                        return False

            return app_valid and user_valid

        except Exception as e:
            print(f"Error verifying Feishu tokens: {e}")
            return False

    def clear_feishu_user_tokens(self) -> None:
        """Clear stored Feishu user tokens when they're invalid."""
        self.config['feishu']['tokens']['user_access_token'] = {
            'token': None,
            'refresh_token': None,
            'expiration_time': None,
            'refresh_token_expiration_time': None
        }
        self._save_config()

    # Outlook Token Management
    def _load_outlook_token(self):
        """Load token for Outlook if exists."""
        access_token, refresh_token, expiration = self.get_outlook_token()
        if access_token and refresh_token:
            self.outlook_account.connection.token_backend.token = {
                'token_type': 'Bearer',
                'access_token': access_token,
                'refresh_token': refresh_token,
                'expires_at': expiration
            }

    def refresh_outlook_token(self) -> bool:
        """Refresh Outlook token."""
        try:
            token_dict = self.outlook_account.connection.token_backend.token
            if token_dict and 'refresh_token' in token_dict:
                result = self.outlook_account.connection.refresh_token()
                if result:
                    # Save the new tokens
                    new_token = self.outlook_account.connection.token_backend.token
                    self.set_outlook_token(
                        new_token['access_token'],
                        new_token['refresh_token'],
                        3600  # Standard expiration
                    )
                    return True
        except Exception as e:
            print(f"Failed to refresh Outlook token: {e}")
        return False

    def verify_outlook_token(self) -> bool:
        """Verify and refresh Outlook token if needed."""
        try:
            # First check if we have a valid token
            if self.outlook_account.is_authenticated:
                return True

            # Check if we have a refresh token to use
            token_dict = self.outlook_account.connection.token_backend.token
            if token_dict and token_dict.get('refresh_token'):
                try:
                    print("Attempting to refresh Outlook token...")
                    result = self.outlook_account.connection.refresh_token()
                    if result:
                        # Save the new tokens
                        new_token = self.outlook_account.connection.token_backend.token
                        self.set_outlook_token(
                            new_token['access_token'],
                            new_token.get('refresh_token', ''),
                            3600
                        )
                        print("Successfully refreshed Outlook token")
                        return True
                    else:
                        print("Token refresh failed")
                except Exception as e:
                    print(f"Error during token refresh: {e}")

            # Only do full reauth if refresh failed or no refresh token
            print("\nFull Outlook authentication required...")
            
            # Get credentials from config
            client_id, client_secret, tenant_id = self.get_outlook_app_info()
            
            # Reinitialize account with proper credentials
            self.outlook_account = Account(
                (client_id, client_secret),
                tenant_id=tenant_id,
                scopes=['offline_access', 'Calendars.ReadWrite']
            )
            
            result = self.outlook_account.authenticate()
            
            if result:
                token = self.outlook_account.connection.token_backend.token
                self.set_outlook_token(
                    token['access_token'],
                    token.get('refresh_token', ''),
                    3600
                )
                self.set_outlook_authenticated(True)
                return True
            
            print("Outlook authentication failed")
            return False
                
        except Exception as e:
            print(f"Error verifying Outlook token: {e}")
            return False

    # Setup Methods
    def setup_feishu(self, app_id: str, app_secret: str) -> bool:
        """Initial Feishu setup."""
        try:
            print("\nInitializing Feishu setup...")
            self.set_feishu_app_info(app_id, app_secret)
            self._setup_clients()
            
            # Get initial app token
            print("Getting app token...")
            if not self.refresh_feishu_app_token():
                print("Failed to get app token")
                return False
                
            # Get user token through OAuth
            print("Starting OAuth process...")
            oauth_code = self.feishu_oauth.obtain_oauth_code()
            if not oauth_code:
                print("Failed to get OAuth code")
                return False

            print(f"Got OAuth code: {oauth_code}")
            
            # Get user token
            print("Getting user token...")
            if not self.get_feishu_user_token_from_code(oauth_code):
                print("Failed to get user token")
                return False

            # List calendars
            print("Fetching calendars...")
            calendars = self.list_feishu_calendars()
            if not calendars:
                print("No calendars found")
                return False

            print("\nAvailable Feishu Calendars:")
            for i, cal in enumerate(calendars, 1):
                name = cal.get('calendar', {}).get('summary') or cal.get('summary', 'Unnamed Calendar')
                description = cal.get('calendar', {}).get('description') or cal.get('description', 'No description')
                print(f"{i}. {name} ({description})")
            
            selections = input("\nEnter calendar numbers to sync (comma-separated) or 'all': ").strip()
            selected_calendars = {}
            
            if selections.lower() == 'all':
                for cal in calendars:
                    cal_id = cal.get('calendar', {}).get('calendar_id') or cal.get('calendar_id')
                    name = cal.get('calendar', {}).get('summary') or cal.get('summary', 'Unnamed Calendar')
                    if cal_id:
                        selected_calendars[cal_id] = name
            else:
                try:
                    indices = [int(i)-1 for i in selections.split(',')]
                    for idx in indices:
                        cal = calendars[idx]
                        cal_id = cal.get('calendar', {}).get('calendar_id') or cal.get('calendar_id')
                        name = cal.get('calendar', {}).get('summary') or cal.get('summary', 'Unnamed Calendar')
                        if cal_id:
                            selected_calendars[cal_id] = name
                except (ValueError, IndexError):
                    print("Invalid selection")
                    return False

            if not selected_calendars:
                print("No calendars selected")
                return False

            print(f"\nSelected {len(selected_calendars)} calendars")
            self.config['feishu']['calendars'] = selected_calendars
            self._save_config()
            return True

        except Exception as e:
            print(f"Setup failed with error: {e}")
            return False

    def setup_outlook(self, client_id: str, client_secret: str, tenant_id: str) -> bool:
        """Initial Outlook setup."""
        try:
            print("\nInitializing Outlook setup...")
            
            # Store the credentials first
            self.set_outlook_app_info(client_id, client_secret, tenant_id)
            
            # Initialize the account with the provided credentials
            self.outlook_account = Account(
                (client_id, client_secret),
                tenant_id=tenant_id,
                scopes=['offline_access', 'Calendars.ReadWrite']
            )
            
            # Authenticate with Outlook
            print("\nStarting Outlook authentication...")
            if not self.authenticate_outlook():
                print("Failed to authenticate with Outlook")
                return False

            print("Successfully authenticated with Outlook")
            return True

        except Exception as e:
            print(f"Outlook setup failed with error: {e}")
            return False
    
    def authenticate_outlook(self) -> bool:
        """Authenticate Outlook using built-in O365 authentication."""
        try:
            # First check if current token is valid
            if self.outlook_account.is_authenticated:
                return True

            # Get credentials from config
            client_id, client_secret, tenant_id = self.get_outlook_app_info()
            
            # Check if we have a refresh token to use
            token_dict = self.outlook_account.connection.token_backend.token
            if token_dict and 'refresh_token' in token_dict:
                try:
                    # Attempt to refresh the token
                    print("Attempting to refresh Outlook token...")
                    result = self.outlook_account.connection.refresh_token()
                    if result:
                        # Save the new tokens
                        new_token = self.outlook_account.connection.token_backend.token
                        self.set_outlook_token(
                            new_token['access_token'],
                            new_token.get('refresh_token', ''),
                            3600
                        )
                        print("Successfully refreshed Outlook token")
                        return True
                except Exception as e:
                    print(f"Token refresh failed: {e}")

            # If we get here, we need a full authentication
            print("\nFull authentication required. Please sign in to your Outlook account in the browser window...")
            
            try:
                # Create a new account instance with the stored credentials
                self.outlook_account = Account(
                    (client_id, client_secret),
                    tenant_id=tenant_id,
                    scopes=['offline_access', 'Calendars.ReadWrite']
                )
                
                result = self.outlook_account.authenticate()
                
                if result:
                    # Save both access and refresh tokens
                    token = self.outlook_account.connection.token_backend.token
                    if 'refresh_token' not in token:
                        print("Warning: No refresh token received. Token will expire in 1 hour.")
                        
                    self.set_outlook_token(
                        token['access_token'],
                        token.get('refresh_token', ''),
                        3600
                    )
                    self.set_outlook_authenticated(True)
                    return True
                
                print("Authentication failed")
                return False
                
            except Exception as e:
                print(f"Error during authentication: {e}")
                return False
                
        except Exception as e:
            print(f"Outlook authentication error: {e}")
            return False
    
    def list_outlook_calendars(self) -> List[Dict]:
        """List all available Outlook calendars."""
        try:
            if not self.verify_outlook_token():
                print("Failed to verify Outlook token")
                return []

            schedule = self.outlook_account.schedule()
            calendars = []
            
            # Get default calendar
            default_calendar = schedule.get_default_calendar()
            if default_calendar:
                calendars.append({
                    'id': default_calendar.calendar_id,
                    'name': 'Default Calendar',
                    'is_default': True
                })

            # Get other calendars
            for calendar in schedule.list_calendars():
                if calendar.calendar_id != calendars[0]['id']:
                    calendars.append({
                        'id': calendar.calendar_id,
                        'name': calendar.name,
                        'is_default': False
                    })

            return calendars

        except Exception as e:
            print(f"Error listing Outlook calendars: {e}")
            return []

    def list_feishu_calendars(self) -> List[Dict]:
        """List all available Feishu calendars."""
        try:
            user_token = self.get_feishu_user_token()
            if not user_token:
                print("No valid user token available")
                return []

            # First, get primary calendar
            response = requests.post(
                'https://open.feishu.cn/open-apis/calendar/v4/calendars/primary',
                headers={'Authorization': f'Bearer {user_token}'}
            )
            
            if response.status_code != 200:
                print(f"Failed to get primary calendar: {response.status_code}")
                return []
                
            data = response.json()
            primary_calendar = data.get('data', {}).get('calendars', [])[0]
            calendars = [primary_calendar]

            # Then get other calendars
            list_response = requests.get(
                'https://open.feishu.cn/open-apis/calendar/v4/calendars',
                headers={'Authorization': f'Bearer {user_token}'}
            )

            if list_response.status_code == 200:
                list_data = list_response.json()
                extra_calendars = list_data.get('data', {}).get('calendars', [])
                calendars.extend(extra_calendars)

            return calendars

        except Exception as e:
            print(f"Error listing calendars: {e}")
            return []

    def is_fully_configured(self) -> bool:
        """Check if all necessary configurations are present."""
        feishu_id, feishu_secret = self.get_feishu_app_info()
        outlook_id, outlook_secret, tenant_id = self.get_outlook_app_info()
        
        return all([
            feishu_id, feishu_secret,
            outlook_id, outlook_secret, tenant_id,
            self.calendar_pairs  # Check for configured calendar pairs
        ])

    @property
    def selected_calendars(self) -> Dict[str, str]:
        """Get selected Feishu calendars."""
        return self.config['feishu']['calendars']
    
    @property
    def calendar_pairs(self) -> List[Dict]:
        """Get configured calendar pairs."""
        return self.config.get('calendar_pairs', [])

    def get_feishu_user_token_from_code(self, code: str) -> bool:
        """Get Feishu user token from OAuth code."""
        try:
            # Create token request using v2 endpoint
            url = 'https://open.feishu.cn/open-apis/authen/v2/oauth/token'
            app_id, app_secret = self.get_feishu_app_info()
            redirect_uri = "http://127.0.0.1:5000/callback"  # Match the authorization URL
            
            headers = {
                'Content-Type': 'application/json'
            }
            
            payload = {
                'grant_type': 'authorization_code',
                'client_id': app_id,
                'client_secret': app_secret,
                'code': code,
                'redirect_uri': redirect_uri  # Add the redirect_uri parameter
            }

            # Debug logging
            print("\nDebug - Token Exchange Request:")
            print(f"URL: {url}")
            print(f"Headers: {headers}")
            debug_payload = {**payload, 'client_secret': '***'}
            print(f"Payload: {json.dumps(debug_payload, indent=2)}")

            response = requests.post(url, headers=headers, json=payload)
            
            if response.status_code != 200:
                print(f"Failed to get access token: {response.status_code}")
                try:
                    error_data = response.json()
                    print(f"Error details: {json.dumps(error_data, indent=2)}")
                except:
                    print(f"Error response: {response.text}")
                return False

            response_data = response.json()
            if response_data.get('code') != 0:
                print(f"Error in response: {response_data}")
                return False

            # Extract token data
            access_token = response_data.get('access_token')
            refresh_token = response_data.get('refresh_token')
            access_token_expires_in = response_data.get('expires_in')
            refresh_token_expires_in = response_data.get('refresh_token_expires_in')

            if not all([access_token, refresh_token, access_token_expires_in, refresh_token_expires_in]):
                print("Missing required token data in response")
                print("Response data:", json.dumps(response_data, indent=2))
                return False

            # Store the new tokens
            self.set_feishu_user_token(
                token=access_token,
                refresh_token=refresh_token,
                expires_in=access_token_expires_in,
                refresh_expires_in=refresh_token_expires_in
            )
            
            print("Successfully obtained new Feishu tokens")
            return True

        except Exception as e:
            print(f"Error getting user token: {e}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            return False

    def get_outlook_token_from_code(self, code: str) -> bool:
        """Get Outlook token from OAuth code."""
        try:
            result = self.outlook_account.connection.request_token(code)
            if result:
                token_dict = self.outlook_account.connection.token_backend.token
                self.set_outlook_token(
                    token_dict['access_token'],
                    token_dict['refresh_token'],
                    3600  # Standard expiration
                )
                self.set_outlook_authenticated(True)
                return True
        except Exception as e:
            print(f"Failed to get Outlook token: {e}")
        return False

    # Token getters and setters
    def set_feishu_app_info(self, app_id: str, app_secret: str) -> None:
        """Set Feishu app credentials."""
        self.config['feishu']['app_info'] = {
            'app_id': app_id,
            'app_secret': app_secret
        }
        self._save_config()

    def get_feishu_app_info(self) -> Tuple[Optional[str], Optional[str]]:
        """Get Feishu app credentials."""
        app_info = self.config['feishu']['app_info']
        return app_info['app_id'], app_info['app_secret']

    def set_feishu_app_token(self, token: str, expires_in: int) -> None:
        """Set Feishu app access token with expiration."""
        expiration = int(time.time()) + expires_in
        self.config['feishu']['tokens']['app_access_token'] = {
            'token': token,
            'expiration_time': expiration
        }
        self._save_config()

    def set_feishu_user_token(self, token: str, refresh_token: str, expires_in: int, refresh_expires_in: int = None) -> None:
        """Set Feishu user access token with refresh token."""
        current_time = int(time.time())
        token_data = {
            'token': token,
            'refresh_token': refresh_token,
            'expiration_time': current_time + expires_in
        }
        
        # Add refresh token expiration if provided
        if refresh_expires_in:
            token_data['refresh_token_expiration_time'] = current_time + refresh_expires_in
        
        self.config['feishu']['tokens']['user_access_token'] = token_data
        self._save_config()

    def get_feishu_app_token(self) -> Optional[str]:
        """Get Feishu app token if valid."""
        token_data = self.config['feishu']['tokens']['app_access_token']
        if not token_data['token'] or not token_data['expiration_time']:
            return None
        
        if int(time.time()) > token_data['expiration_time']:
            return None
        
        return token_data['token']

    def get_feishu_user_token(self) -> Optional[str]:
        """Get Feishu user token if valid."""
        token_data = self.config['feishu']['tokens']['user_access_token']
        if not token_data.get('token') or not token_data.get('expiration_time'):
            return None
        
        if int(time.time()) > token_data['expiration_time']:
            return None
        
        return token_data['token']
    
    def get_feishu_refresh_token(self) -> Optional[str]:
        """Get Feishu refresh token if not expired."""
        try:
            token_data = self.config['feishu']['tokens']['user_access_token']
            refresh_token = token_data.get('refresh_token')
            
            # Basic validation of token format
            if not refresh_token or len(refresh_token.strip()) < 10:  # Simple length check
                print("Invalid refresh token format in storage")
                return None
                
            refresh_expiration = token_data.get('refresh_token_expiration_time')
            
            # If expiration time exists, check it
            if refresh_expiration and int(time.time()) >= refresh_expiration:
                print("Refresh token has expired")
                return None
                
            return refresh_token.strip()
            
        except Exception as e:
            print(f"Error retrieving refresh token: {e}")
            return None


    def is_feishu_app_token_valid(self) -> bool:
        """Check if Feishu app token is valid."""
        return self.get_feishu_app_token() is not None

    def is_feishu_user_token_valid(self) -> bool:
        """Check if Feishu user token is valid and not expired."""
        try:
            token = self.get_feishu_user_token()
            if not token:
                return False
                
            # Add additional validation by making a test API call
            response = requests.get(
                'https://open.feishu.cn/open-apis/calendar/v4/calendars',
                headers={'Authorization': f'Bearer {token}'}
            )
            
            return response.status_code == 200
                
        except Exception:
            return False
        
    def set_outlook_app_info(self, client_id: str, client_secret: str, tenant_id: str) -> None:
        """Set Outlook app credentials."""
        self.config['outlook']['app_info'] = {
            'client_id': client_id,
            'client_secret': client_secret,
            'tenant_id': tenant_id
        }
        self._save_config()

    def get_outlook_app_info(self) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """Get Outlook app credentials."""
        app_info = self.config['outlook']['app_info']
        return app_info['client_id'], app_info['client_secret'], app_info['tenant_id']

    def set_outlook_token(self, access_token: str, refresh_token: str, expires_in: int) -> None:
        """Set Outlook tokens with expiration."""
        expiration = int(time.time()) + expires_in
        self.config['outlook']['tokens'] = {
            'access_token': access_token,
            'refresh_token': refresh_token,
            'expiration_time': expiration
        }
        self._save_config()

    def get_outlook_token(self) -> Tuple[Optional[str], Optional[str], Optional[int]]:
        """Get Outlook token info if valid."""
        token_data = self.config['outlook']['tokens']
        
        # Check if we have all required token data
        if (not token_data['access_token'] or 
            not token_data['refresh_token'] or
            not token_data['expiration_time']):
            return None, None, None
        
        # Check if token is expired
        if int(time.time()) > token_data['expiration_time']:
            return None, None, None
        
        return (token_data['access_token'], 
                token_data['refresh_token'],
                token_data['expiration_time'])

    def set_outlook_authenticated(self, status: bool) -> None:
        """Set Outlook authentication status."""
        self.config['outlook']['authenticated'] = status
        self._save_config()

    def setup_calendar_pairs(self) -> bool:
        """Setup calendar pairs for syncing."""
        try:
            print("\nSetting up calendar pairs...")
            
            # Get Feishu calendars
            feishu_calendars = self.list_feishu_calendars()
            if not feishu_calendars:
                print("No Feishu calendars found")
                return False

            # Get Outlook calendars
            outlook_calendars = self.list_outlook_calendars()
            if not outlook_calendars:
                print("No Outlook calendars found")
                return False

            print("\nAvailable Feishu Calendars:")
            for i, cal in enumerate(feishu_calendars, 1):
                name = cal.get('calendar', {}).get('summary') or cal.get('summary', 'Unnamed Calendar')
                calendar_id = cal.get('calendar', {}).get('calendar_id') or cal.get('calendar_id')
                print(f"{i}. {name} (ID: {calendar_id})")

            print("\nAvailable Outlook Calendars:")
            for i, cal in enumerate(outlook_calendars, 1):
                print(f"{i}. {cal['name']} (ID: {cal['id']})")

            calendar_pairs = []
            while True:
                try:
                    print("\nEnter a calendar pair (or press Enter to finish):")
                    feishu_input = input("Enter Feishu calendar number: ").strip()
                    if not feishu_input:
                        break

                    outlook_input = input("Enter Outlook calendar number: ").strip()
                    if not outlook_input:
                        break

                    feishu_idx = int(feishu_input) - 1
                    outlook_idx = int(outlook_input) - 1

                    if (0 <= feishu_idx < len(feishu_calendars) and 
                        0 <= outlook_idx < len(outlook_calendars)):
                        
                        feishu_cal = feishu_calendars[feishu_idx]
                        outlook_cal = outlook_calendars[outlook_idx]

                        pair = {
                            'feishu': {
                                'id': feishu_cal.get('calendar', {}).get('calendar_id') or feishu_cal.get('calendar_id'),
                                'name': feishu_cal.get('calendar', {}).get('summary') or feishu_cal.get('summary')
                            },
                            'outlook': {
                                'id': outlook_cal['id'],
                                'name': outlook_cal['name']
                            }
                        }
                        calendar_pairs.append(pair)
                        print(f"Pair added: {pair['feishu']['name']} -> {pair['outlook']['name']}")
                    else:
                        print("Invalid calendar numbers")

                except ValueError:
                    print("Invalid input. Please enter numbers.")
                except Exception as e:
                    print(f"Error adding pair: {e}")

            if not calendar_pairs:
                print("No calendar pairs configured")
                return False

            self.config['calendar_pairs'] = calendar_pairs
            self._save_config()
            print(f"\nSuccessfully configured {len(calendar_pairs)} calendar pairs")
            return True

        except Exception as e:
            print(f"Error setting up calendar pairs: {e}")
            return False

if __name__ == '__main__':
    auth_handler = AuthHandler()
    
    if not auth_handler.is_fully_configured():
        print("First-time setup required")
        
        print("\nFeishu Setup:")
        feishu_id = input("Enter Feishu App ID: ")
        feishu_secret = input("Enter Feishu App Secret: ")
        
        if not auth_handler.setup_feishu(feishu_id, feishu_secret):
            print("Feishu setup failed")
            exit(1)
        
        print("\nOutlook Setup:")
        outlook_id = input("Enter Outlook Client ID: ")
        outlook_secret = input("Enter Outlook Client Secret: ")
        tenant_id = input("Enter Azure AD Tenant ID: ")
        
        if not auth_handler.setup_outlook(outlook_id, outlook_secret, tenant_id):
            print("Outlook setup failed")
            exit(1)
            
        if not auth_handler.setup_calendar_pairs():
            print("Calendar pair setup failed")
            exit(1)
        
        print("\nSetup completed successfully!")
    else:
        print("Verifying tokens...")
        if auth_handler.verify_feishu_tokens() and auth_handler.verify_outlook_token():
            print("All tokens are valid!")
            
            # Allow reconfiguring calendar pairs if requested
            reconfigure = input("\nWould you like to reconfigure calendar pairs? (y/N): ").lower().strip()
            if reconfigure == 'y':
                if auth_handler.setup_calendar_pairs():
                    print("Calendar pairs updated successfully!")
                else:
                    print("Failed to update calendar pairs")
            
        else:
            print("Some tokens need refresh. Run sync script to handle this automatically.")