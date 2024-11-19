import urllib.parse
import os
import base64
import uvicorn
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.responses import RedirectResponse

class OutlookOAuth:
    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        self.CLIENT_ID = client_id
        self.CLIENT_SECRET = client_secret
        self.TENANT_ID = tenant_id
        self.REDIRECT_URI = "http://localhost:5001/callback"
        self.SCOPE = "https://graph.microsoft.com/Calendars.ReadWrite offline_access"
        self.STATE = base64.urlsafe_b64encode(os.urandom(32)).rstrip(b'=').decode('utf-8')
        
        # Initialize FastAPI
        self.app = FastAPI()
        
        # OAuth code storage
        self.oauth_code = None
        
        # Setup routes
        self._setup_routes()
    
    def _setup_routes(self):
        @self.app.get("/")
        def home():
            return RedirectResponse(self.construct_oauth_url())

        @self.app.get("/callback")
        def callback(code: Optional[str] = None):
            if not code:
                raise HTTPException(status_code=400, detail="No OAuth code received")
            self.oauth_code = code
            return {"message": "OAuth code received. You can close this window now and keyboard interrupt the server in the terminal."}

    def construct_oauth_url(self) -> str:
        base_url = f"https://login.microsoftonline.com/{self.TENANT_ID}/oauth2/v2.0/authorize"
        params = {
            "client_id": self.CLIENT_ID,
            "redirect_uri": urllib.parse.quote(self.REDIRECT_URI, safe=''),
            "scope": urllib.parse.quote(self.SCOPE, safe=''),
            "response_type": "code",
            "state": self.STATE,
            "response_mode": "query"
        }
        return f"{base_url}?{'&'.join(f'{k}={v}' for k, v in params.items())}"

    def obtain_oauth_code(self) -> str:
        print("\nPlease visit the following URL to authorize the app and Ctrl-C when you get the success page:\n")
        print(self.construct_oauth_url())
        
        # Run the server
        uvicorn.run(self.app, host="127.0.0.1", port=5001, log_level="error")
        
        # After server stops (when callback received), return the code
        return self.oauth_code

def get_oauth_code(client_id: str, client_secret: str, tenant_id: str) -> str:
    # Allow insecure transport for local development
    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
    
    oauth = OutlookOAuth(client_id, client_secret, tenant_id)
    return oauth.obtain_oauth_code()

if __name__ == "__main__":
    try:
        code = get_oauth_code(
            client_id="your_client_id",
            client_secret="your_client_secret",
            tenant_id="your_tenant_id"
        )
        print(f"\nReceived OAuth code: {code}")
    except Exception as e:
        print(f"\nError: {e}")