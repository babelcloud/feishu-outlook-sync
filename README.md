# Feishu-Outlook Calendar Sync

A Python application that provides real-time synchronization between Feishu (Lark) and Microsoft Outlook calendars.

## Features

- One-way sync from Feishu to Outlook
- Support for multiple Feishu calendars
- Real-time continuous synchronization (5-minute intervals)
- Automatic timezone handling
- Duplicate event detection
- Token management with automatic refresh
- Support for recurring events
- Support for meeting URLs and locations

## Prerequisites

### Required Python Version
- Recommended Python 3.12
  - **NOTE**: Python 3.13 removed the `cgi` module, which the lark sdk relies on, if you are using 3.13+, install `legacy-cgi` with pip. (`pip install legacy-cgi`)
- Not tested on other versions, cannot confirm stability

### Dependencies
```bash
pip install pyyaml
pip install O365
pip install lark-oapi
pip install pytz
pip install fastapi
pip install uvicorn
```

### Required Credentials

#### Feishu (Lark) Setup
1. Create a Feishu application in the [Feishu Open Platform](https://open.feishu.cn/)
2. Add the following API scopes:
   - calendar:calendar.event:read
   - calendar:calendar:read
   - calendar:calendar:readonly
   - offline_access
3. Add a redirect URL in security settings: `http://127.0.0.1:5000/callback`
4. Note down the following:
   - App ID
   - App Secret

#### Microsoft Outlook Setup
1. Register an application in the [Azure Portal](https://portal.azure.com/)
2. Add the following API permissions:
   - Microsoft Graph > Calendars.ReadWrite
   - Microsoft Graph > Calendars.Read
3. Note down the following:
   - Client ID
   - Client Secret
   - Tenant ID (for enterprise applications)
4. Add a redirect URL in Authentication > Web: `https://login.microsoftonline.com/common/oauth2/nativeclient`

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/feishu-outlook-sync.git
cd feishu-outlook-sync
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration
The application uses a YAML file for storing credentials and tokens. Running `auth_handler.py` will automatically configure this for you.

### tokens.yaml
```yaml
feishu:
  app_info:
    app_id: your_feishu_app_id
    app_secret: your_feishu_app_secret
  tokens:
    app_access_token:
      token: null
      expiration_time: null
    user_access_token:
      token: null
      refresh_token: null
      expiration_time: null
  calendars:
    calendar_id_1: "Calendar Name 1"
    calendar_id_2: "Calendar Name 2"

outlook:
  app_info:
    client_id: your_outlook_client_id
    client_secret: your_outlook_client_secret
    tenant_id: your_tenant_id
  tokens:
    access_token: null
    refresh_token: null
    expiration_time: null
  calendar_id:
    id: null
  authenticated: false
```

## Usage

### Setup & Usage

1. Run the authentication setup:
```bash
python auth_handler.py
```
This will:
- Prompt for credentials
- Handle OAuth authentication for both services
- Display available calendars from both Feishu and Outlook
- Allow you to create calendar pairs for syncing
- Store configuration in tokens.yaml

When creating calendar pairs:

- You can pair the same Feishu calendar with multiple Outlook calendars
- You can pair multiple Feishu calendars with the same Outlook calendar
- Enter pairs one at a time, press Enter without input to finish


1. Start the sync process:
```bash
python main.py
```

By default, the main.py will run continously, syncing every 5 minutes, you can modify the code to have it simply sync once per run, such that it can be ran with a cronjob or similar.

```python
if __name__ == "__main__":
    if not auth_handler.is_fully_configured():
        print("Please run auth_handler.py first to setup authentication")
        sys.exit(1)

    auth_handler = AuthHandler()
    sync_calendars(auth_handler)
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and feature requests, please use the GitHub issue tracker.