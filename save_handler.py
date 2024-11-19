import yaml
import time
from typing import Dict, Optional, Tuple

class SaveHandler:
    def __init__(self, yaml_file: str = 'tokens.yaml'):
        self.yaml_file = yaml_file
        self.config = self._load_config()

    def _load_config(self) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(self.yaml_file, 'r') as file:
                return yaml.safe_load(file) or {}
        except FileNotFoundError:
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
                            'expiration_time': None
                        }
                    },
                    'calendar_id': {
                        'id': None
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
                        'expiration_time': None
                    },
                    'calendar_id': {
                        'id': None
                    },
                    'authenticated': False
                }
            }

    def _save_config(self) -> None:
        """Save configuration to YAML file."""
        with open(self.yaml_file, 'w') as file:
            yaml.dump(self.config, file, default_flow_style=False)

    def set_feishu_app_info(self, app_id: str, app_secret: str) -> None:
        """
        Set Feishu app credentials.
        
        Args:
            app_id: The Feishu app ID
            app_secret: The Feishu app secret
        """
        self.config['feishu']['app_info'] = {
            'app_id': app_id,
            'app_secret': app_secret
        }
        self._save_config()

    def get_feishu_app_info(self) -> Tuple[Optional[str], Optional[str]]:
        """
        Get Feishu app credentials.
        
        Returns:
            Tuple of (app_id, app_secret)
        """
        app_info = self.config['feishu']['app_info']
        return app_info['app_id'], app_info['app_secret']

    def set_feishu_app_token(self, token: str, expires_in: int) -> None:
        """
        Set Feishu application access token with expiration time.
        
        Args:
            token: The token string
            expires_in: Expiration time in seconds
        """
        expiration = int(time.time()) + expires_in
        self.config['feishu']['tokens']['app_access_token'] = {
            'token': token,
            'expiration_time': expiration
        }
        self._save_config()

    def set_feishu_user_token(self, token: str, expires_in: int) -> None:
        """
        Set Feishu user access token with expiration time.
        
        Args:
            token: The token string
            expires_in: Expiration time in seconds
        """
        expiration = int(time.time()) + expires_in
        self.config['feishu']['tokens']['user_access_token'] = {
            'token': token,
            'expiration_time': expiration
        }
        self._save_config()

    def set_feishu_calendar_id(self, calendar_id: str) -> None:
        """Set calendar ID."""
        self.config['feishu']['calendar_id'] = {'id': calendar_id}
        self._save_config()

    def get_feishu_app_token(self) -> Optional[str]:
        """
        Get Feishu application access token if valid.
        Returns None if token is expired or not set.
        """
        token_data = self.config['feishu']['tokens']['app_access_token']
        if not token_data['token'] or not token_data['expiration_time']:
            return None
        
        if int(time.time()) > token_data['expiration_time']:
            return None
        
        return token_data['token']

    def get_feishu_user_token(self) -> Optional[str]:
        """
        Get Feishu user access token if valid.
        Returns None if token is expired or not set.
        """
        token_data = self.config['feishu']['tokens']['user_access_token']
        if not token_data['token'] or not token_data['expiration_time']:
            return None
        
        if int(time.time()) > token_data['expiration_time']:
            return None
        
        return token_data['token']

    def get_feishu_calendar_id(self) -> Optional[str]:
        """Get calendar ID."""
        return self.config['feishu']['calendar_id']['id']

    def is_feishu_app_token_valid(self) -> bool:
        """Check if Feishu app token is valid and not expired."""
        return self.get_feishu_app_token() is not None

    def is_feishu_user_token_valid(self) -> bool:
        """Check if Feishu user token is valid and not expired."""
        return self.get_feishu_user_token() is not None
    
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

    def set_outlook_authenticated(self, authenticated: bool) -> None:
        """Set Outlook authentication status."""
        self.config['outlook']['authenticated'] = authenticated
        self._save_config()

    def is_outlook_authenticated(self) -> bool:
        """Check if Outlook is authenticated."""
        return self.config.get('outlook', {}).get('authenticated', False)

    def set_outlook_calendar_id(self, calendar_id: str) -> None:
        """Set Outlook calendar ID."""
        self.config['outlook']['calendar_id'] = {'id': calendar_id}
        self._save_config()

    def get_outlook_calendar_id(self) -> Optional[str]:
        """Get Outlook calendar ID."""
        return self.config['outlook']['calendar_id']['id']
    
    def set_outlook_token(self, access_token: str, expires_in: int) -> None:
        """Set Outlook tokens with expiration."""
        expiration = int(time.time()) + expires_in
        self.config['outlook']['tokens'] = {
            'access_token': access_token,
            'expiration_time': expiration
        }
        self._save_config()

    def get_outlook_token(self) -> Tuple[Optional[str], Optional[int]]:
        """Get Outlook token and expiration if valid."""
        token_data = self.config['outlook']['tokens']
        if (not token_data['access_token'] or 
            not token_data['expiration_time']):
            return None, None
        
        return (token_data['access_token'], 
                token_data['expiration_time'])