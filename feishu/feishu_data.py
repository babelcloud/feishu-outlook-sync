import json
import time
import requests
import lark_oapi as lark
from datetime import datetime, timezone
from lark_oapi.api.authen.v1 import *
from lark_oapi.api.auth.v3 import *
from lark_oapi.api.calendar.v4 import *

class DataHandler:
    def __init__(self, app_id, app_secret):
        self.APP_ID = app_id
        self.APP_SECRET = app_secret
        self.save_handler = SaveHandler()

        self.client = lark.Client.builder() \
            .app_id(self.APP_ID) \
            .app_secret(self.APP_SECRET) \
            .enable_set_token(True) \
            .log_level(lark.LogLevel.DEBUG) \
            .build()

    def obtain_app_access_token(self):
        request: InternalAppAccessTokenRequest = InternalAppAccessTokenRequest.builder() \
            .request_body(InternalAppAccessTokenRequestBody.builder()
                        .app_id(self.APP_ID)
                        .app_secret(self.APP_SECRET)
                        .build()) \
            .build()

        response: InternalAppAccessTokenResponse = self.client.auth.v3.app_access_token.internal(request)

        if not response.success():
            lark.logger.error(
                f"Failed to get app access token: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return None

        response_data = json.loads(response.raw.content.decode('utf-8'))
        
        # Extract token and expire time
        token = response_data.get('app_access_token')
        expire = response_data.get('expire')
        
        lark.logger.info(f"App Access Token obtained")
        self.save_handler.set_feishu_app_token(token, expire)
        
        return token

    def obtain_user_access_token(self, code: str):
        request: CreateAccessTokenRequest = CreateAccessTokenRequest.builder() \
            .request_body(CreateAccessTokenRequestBody.builder()
                        .grant_type("authorization_code")
                        .code(code)
                        .build()) \
            .build()

        response: CreateAccessTokenResponse = self.client.authen.v1.access_token.create(request)

        if not response.success():
            lark.logger.error(
                f"Failed to get user access token: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return None

        response_data = json.loads(response.raw.content.decode('utf-8'))
        
        token_data = response_data['data']
        token = token_data.get('access_token')
        refresh_token = token_data.get('refresh_token')
        expire = token_data.get('expires_in')
        
        lark.logger.info("User tokens obtained")
        
        self.save_handler.set_feishu_user_token(token, refresh_token, expire)
        return token

    def refresh_user_token(self):
        """Refresh the user access token using the refresh token."""
        refresh_token = self.save_handler.get_feishu_refresh_token()
        if not refresh_token:
            lark.logger.error("No refresh token available")
            return None

        request: RefreshAccessTokenRequest = RefreshAccessTokenRequest.builder() \
            .request_body(RefreshAccessTokenRequestBody.builder()
                        .grant_type("refresh_token")
                        .refresh_token(refresh_token)
                        .build()) \
            .build()

        response: RefreshAccessTokenResponse = self.client.authen.v1.access_token.refresh(request)

        if not response.success():
            lark.logger.error(
                f"Failed to refresh token: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return None

        response_data = json.loads(response.raw.content.decode('utf-8'))
        token_data = response_data['data']
        
        token = token_data.get('access_token')
        new_refresh_token = token_data.get('refresh_token')
        expire = token_data.get('expires_in')

        lark.logger.info("User tokens refreshed")
        self.save_handler.set_feishu_user_token(token, new_refresh_token, expire)
        return token

    def verify_and_refresh_tokens(self):
        """Verify tokens and refresh if needed."""
        # Check user token first
        if not self.save_handler.is_feishu_user_token_valid():
            # Try to refresh the token
            if self.save_handler.get_feishu_refresh_token():
                lark.logger.info("Attempting to refresh user token")
                if self.refresh_user_token():
                    lark.logger.info("Successfully refreshed user token")
                else:
                    lark.logger.error("Failed to refresh user token")
                    return False
            else:
                lark.logger.error("No refresh token available and user token is invalid")
                return False

        # Check app token
        if not self.save_handler.is_feishu_app_token_valid():
            lark.logger.info("App token invalid, obtaining new one")
            if not self.obtain_app_access_token():
                lark.logger.error("Failed to obtain new app token")
                return False

        return True

    def get_future_calendar_events(self, calendar_id: str, access_token: str):
        """Get calendar events with automatic token refresh."""
        # Verify and refresh tokens if needed
        if not self.verify_and_refresh_tokens():
            lark.logger.error("Failed to verify/refresh tokens")
            return None

        # Use the verified/refreshed token
        access_token = self.save_handler.get_feishu_user_token()
        
        payload = {'start_time': str(int(time.time()))}
        response = requests.get(
            f'https://open.feishu.cn/open-apis/calendar/v4/calendars/{calendar_id}/events',
            headers={'Authorization': f'Bearer {access_token}'},
            params=payload
        )
        
        if response.status_code == 401:  # Unauthorized
            lark.logger.warning("Token unauthorized, attempting refresh")
            if self.verify_and_refresh_tokens():
                # Retry with new token
                access_token = self.save_handler.get_feishu_user_token()
                response = requests.get(
                    f'https://open.feishu.cn/open-apis/calendar/v4/calendars/{calendar_id}/events',
                    headers={'Authorization': f'Bearer {access_token}'},
                    params=payload
                )

        response_data = response.json()

        if response.status_code != 200:
            lark.logger.error(f"Failed to get calendar events: {response.status_code}, {response_data}")
            return None
        
        events = response_data.get('data', {}).get('items', [])
        lark.logger.info(f"Retrieved {len(events)} events")
        return events