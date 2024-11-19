import json
import time
import requests
import sys
import os
import lark_oapi as lark
from .feishu_oauth import get_oauth_code
from lark_oapi.api.authen.v1 import *
from lark_oapi.api.auth.v3 import *
from lark_oapi.api.calendar.v4 import *
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from save_handler import SaveHandler

class DataHandler:
    def __init__(self, app_id, app_secret):
        self.APP_ID = app_id
        self.APP_SECRET = app_secret

        # Initialize the SaveHandler
        self.save_handler = SaveHandler()

        # Initialize the client
        self.client = lark.Client.builder() \
            .app_id(self.APP_ID) \
            .app_secret(self.APP_SECRET) \
            .enable_set_token(True) \
            .log_level(lark.LogLevel.DEBUG) \
            .build()
        
    def obtain_oauth_code(self):
        return get_oauth_code(self.APP_ID, self.APP_SECRET)
    
    def obtain_app_access_token(self):
        # Taken partially from oapi-python-sdk/samples/api/auth/v3/internal_app_access_token_sample.py
        request: InternalAppAccessTokenRequest = InternalAppAccessTokenRequest.builder() \
            .request_body(InternalAppAccessTokenRequestBody.builder()
                        .app_id(self.APP_ID)
                        .app_secret(self.APP_SECRET)
                        .build()) \
            .build()

        response: InternalAppAccessTokenResponse = self.client.auth.v3.app_access_token.internal(request)

        if not response.success():
            lark.logger.error(
                f"client.auth.v3.app_access_token.internal failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return

        response_data = json.loads(response.raw.content.decode('utf-8'))
        
        # Extract token and expire time
        token = response_data.get('app_access_token')
        expire = response_data.get('expire')
        
        lark.logger.info(f"App Access Token: {token}")
        lark.logger.info(f"Expires in: {expire} seconds")
        
        self.save_handler.set_feishu_app_token(token, expire)
        
        return token
    
    def obtain_user_access_token(self, code: str):
        # Taken partially from oapi-python-sdk/samples/api/authen/v1/create_access_token_sample.py
        request: CreateAccessTokenRequest = CreateAccessTokenRequest.builder() \
            .request_body(CreateAccessTokenRequestBody.builder()
                        .grant_type("authorization_code")
                        .code(code)
                        .build()) \
            .build()

        response: CreateAccessTokenResponse = self.client.authen.v1.access_token.create(request)

        if not response.success():
            lark.logger.error(
                f"client.authen.v1.access_token.create failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return

        response_data = json.loads(response.raw.content.decode('utf-8'))
        
        # Extract token and expire time
        token_data = response_data['data']
        token = token_data.get('access_token')
        refresh_token = token_data.get('refresh_token')
        expire = token_data.get('expires_in')
        
        lark.logger.info(f"User Access Token: {token}")
        lark.logger.info(f"User Access Refresh Token: {refresh_token}")
        lark.logger.info(f"Expires in: {expire} seconds")

        self.save_handler.set_feishu_user_token(token, expire)
        
        return token

    def get_primary_calendar(self, access_token: str):
        # Lark's SDK uses the tenant access token for some reason, so we need to use requests for this
        response = requests.post('https://open.feishu.cn/open-apis/calendar/v4/calendars/primary', headers={'Authorization': f'Bearer {access_token}'})
        response_data = response.json()

        lark.logger.info(response_data)

        if response.status_code != 200:
            lark.logger.error(f"Failed to get primary calendar, status code: {response.status_code}, response: {response_data}")
            return
        
        calendar = response_data.get('data').get('calendars')[0].get('calendar')
        lark.logger.info(calendar)

        self.save_handler.set_feishu_calendar_id(calendar['calendar_id'])

        return calendar['calendar_id']
    
    def get_future_calendar_events(self, calendar_id: str, access_token: str):
        payload = {'start_time': str(int(time.time()))}
        response = requests.get(f'https://open.feishu.cn/open-apis/calendar/v4/calendars/{calendar_id}/events', headers={'Authorization': f'Bearer {access_token}'}, params=payload)
        response_data = response.json()

        lark.logger.info(response_data)

        if response.status_code != 200:
            lark.logger.error(f"Failed to get calendar events, status code: {response.status_code}, response: {response_data}")
            return
        
        events = response_data.get('data').get('items')

        lark.logger.info(events)

        return events
