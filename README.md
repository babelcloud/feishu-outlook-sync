# Feishu-Outlook Calendar Sync

A Python application that provides real-time synchronization between Feishu (Lark) and Microsoft Outlook calendars.

## Features

- One-way sync from Feishu to Outlook
- Real-time continuous synchronization
- Automatic timezone handling
- Duplicate event detection
- Support for recurring events
- Token-based authentication
- Persistent token storage

## Prerequisites

### Required Python Version
- Python 3.9 or higher

### Dependencies
```bash
pip install O365
pip install lark-oapi
pip install pytz
pip install fastapi
pip install uvicorn
```

### Required Credentials

#### Feishu (Lark) Setup
1. Create a Feishu application in the [Feishu Open Platform](https://open.feishu.cn/)
2. Enable Calendar permissions, and add a Bot
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

The application uses a YAML file for storing credentials and tokens. On first run, you'll be prompted to enter your credentials.

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
      expiration_time: null
  calendar_id:
    id: null

outlook:
  app_info:
    client_id: your_outlook_client_id
    client_secret: your_outlook_client_secret
    tenant_id: your_tenant_id
  tokens:
    access_token: null
    expiration_time: null
  calendar_id:
    id: null
  authenticated: false
```

## Usage

### Basic Usage

1. Run the application:
```bash
python main.py
```

2. On first run:
   - You'll be prompted to enter your Feishu and Outlook credentials
   - Links will be given for authentication (OAuth)
   - Follow the prompts to grant calendar access

3. The application will:
   - Start syncing immediately
   - Check for new events every 5 minutes
   - Log sync activities to the console


## Behavior Specifications

### Event Synchronization
- Only future events are synchronized
- Events are matched based on title and start time
- Timezone differences are automatically handled
- Duplicate events are skipped

### Authentication
- Tokens are automatically refreshed
- Failed authentications trigger re-authorization
- Credentials are securely stored in tokens.yaml

### Error Handling
- Failed syncs are logged but don't stop the process
- Network interruptions are handled gracefully
- Authentication errors trigger automatic token refresh

## Limitations

- One-way sync only (Feishu → Outlook)
- No support for event updates (only creation)
- No attendee synchronization
- Does not sync event cancellations

## Troubleshooting

### Common Issues

1. Authentication Failures
   - Verify credentials in tokens.yaml
   - Check token expiration
   - Ensure proper API permissions
   - Restart the script after one run

2. Missing Events
   - Check timezone settings
   - Verify event dates are in the future
   - Check for duplicate detection issues

3. Sync Issues
   - Verify network connectivity
   - Check both calendars are accessible
   - Verify token permissions

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and feature requests, please use the GitHub issue tracker.