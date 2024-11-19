import urllib.parse
from typing import Optional
import uvicorn
import lark_oapi as lark
from fastapi import FastAPI, HTTPException
from fastapi.responses import RedirectResponse

class FeishuOAuth:
    def __init__(self, app_id: str, app_secret: str):
        self.APP_ID = app_id
        self.APP_SECRET = app_secret
        self.REDIRECT_URI = "http://127.0.0.1:5000/callback"
        self.SCOPE = "calendar:calendar:readonly calendar:calendar:read calendar:calendar.event:read"
        self.STATE = "RANDOMSTATE"
        
        # Initialize FastAPI and lark client
        self.app = FastAPI()
        self.client = lark.Client.builder() \
            .app_id(self.APP_ID) \
            .app_secret(self.APP_SECRET) \
            .enable_set_token(True) \
            .log_level(lark.LogLevel.DEBUG) \
            .build()
        
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
        base_url = "https://open.feishu.cn/open-apis/authen/v1/authorize"
        params = {
            "app_id": self.APP_ID,
            "redirect_uri": urllib.parse.quote(self.REDIRECT_URI, safe=''),
            "scope": urllib.parse.quote(self.SCOPE, safe=''),
            "state": self.STATE
        }
        return f"{base_url}?{'&'.join(f'{k}={v}' for k, v in params.items())}"

    def obtain_oauth_code(self) -> str:
        """Get OAuth code by providing a URL to visit."""
        print("\nPlease visit the following URL to authorize the app and Ctrl-C when you get the success page:\n")
        print(self.construct_oauth_url())
        
        # Run the server
        uvicorn.run(self.app, host="127.0.0.1", port=5000, log_level="error")
        
        # After server stops (when callback received), return the code
        return self.oauth_code

def get_oauth_code(app_id: str, app_secret: str) -> str:
    """Helper function to get an OAuth code."""
    oauth = FeishuOAuth(app_id, app_secret)
    return oauth.obtain_oauth_code()

if __name__ == "__main__":
    try:
        code = get_oauth_code(
            app_id="",
            app_secret=""
        )
        print(f"\nReceived OAuth code: {code}")
    except Exception as e:
        print(f"\nError: {e}")