import os
import sys
import yaml
import time
import threading
from typing import Dict, List, Optional
from dataclasses import dataclass
from main import run_sync, run_continuous_sync
from auth_handler import AuthHandler
from db_handler import DBHandler
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

@dataclass
class SyncConfig:
    path: str
    name: str
    valid: bool
    error: Optional[str] = None

class MultiSync:
    """Class that handles multiple sync configurations."""
    
    def __init__(self, config_dir: str = 'configs'):
        self.config_dir = config_dir
        self.configs: List[SyncConfig] = []
        self.threads: Dict[str, threading.Thread] = {}
        self.stop_events: Dict[str, threading.Event] = {}
        
        # Initialize database handler
        self.db_handler = DBHandler(
            host=os.environ.get('PGHOST', 'ep-red-brook-aawc47cf-pooler.westus3.azure.neon.tech'),
            database=os.environ.get('PGDATABASE', 'CalendarSync'),
            user=os.environ.get('PGUSER', 'CalendarSync_owner'),
            password=os.environ.get('PGPASSWORD', 'npg_Had49rjAFUVg')
        )
        
        if not self.db_handler.connect():
            print("Failed to connect to database, proceeding without database logging")
            self.db_handler = None
        
        # Create config directory if it doesn't exist
        os.makedirs(self.config_dir, exist_ok=True)

    def validate_yaml(self, path: str) -> SyncConfig:
        """Validate that a YAML file is a valid configuration."""
        try:
            # Parse filename to get sync name
            filename = os.path.basename(path)
            name = os.path.splitext(filename)[0]
            
            with open(path, 'r') as f:
                config = yaml.safe_load(f)
            
            # Basic validation
            if not isinstance(config, dict):
                return SyncConfig(path=path, name=name, valid=False, error="Configuration must be a dictionary")
            
            # Check for required sections
            if 'feishu' not in config:
                return SyncConfig(path=path, name=name, valid=False, error="Missing 'feishu' section")
            
            if 'outlook' not in config:
                return SyncConfig(path=path, name=name, valid=False, error="Missing 'outlook' section")
            
            # Check for required Feishu credentials
            feishu = config['feishu']
            if not isinstance(feishu, dict):
                return SyncConfig(path=path, name=name, valid=False, error="Feishu configuration must be a dictionary")
            
            required_feishu_keys = ['app_id', 'app_secret', 'app_token', 'user_token', 'refresh_token']
            missing_keys = [key for key in required_feishu_keys if key not in feishu]
            if missing_keys:
                return SyncConfig(path=path, name=name, valid=False, error=f"Missing Feishu keys: {', '.join(missing_keys)}")
            
            # Check for required Outlook credentials
            outlook = config['outlook']
            if not isinstance(outlook, dict):
                return SyncConfig(path=path, name=name, valid=False, error="Outlook configuration must be a dictionary")
            
            required_outlook_keys = ['client_id', 'client_secret', 'tenant_id', 'token', 'refresh_token']
            missing_keys = [key for key in required_outlook_keys if key not in outlook]
            if missing_keys:
                return SyncConfig(path=path, name=name, valid=False, error=f"Missing Outlook keys: {', '.join(missing_keys)}")
            
            # Check for calendars
            if 'calendars' not in config:
                return SyncConfig(path=path, name=name, valid=False, error="Missing 'calendars' section with calendar pairs")
            
            return SyncConfig(path=path, name=name, valid=True)
            
        except Exception as e:
            return SyncConfig(path=path, name=name, valid=False, error=str(e))

    def load_configs(self) -> None:
        """Load and validate all configs in the config directory."""
        self.configs = []
        
        if not os.path.isdir(self.config_dir):
            print(f"Config directory {self.config_dir} does not exist")
            return
        
        for filename in os.listdir(self.config_dir):
            if filename.endswith('.yaml') or filename.endswith('.yml'):
                path = os.path.join(self.config_dir, filename)
                config = self.validate_yaml(path)
                self.configs.append(config)
                
        # Log the loaded configurations to the database
        if self.db_handler:
            sync_id = self.db_handler.log_sync_start("multi_sync_load_configs")
            message = f"Loaded {len(self.configs)} configurations: {', '.join([c.name for c in self.configs])}"
            self.db_handler.log_sync_end(sync_id, True, message)

    def print_status(self) -> None:
        """Print current status of all configurations."""
        print("\nSync Configurations:")
        print("---------------------")
        
        valid_count = 0
        invalid_count = 0
        
        for idx, config in enumerate(self.configs, 1):
            status = "VALID" if config.valid else "INVALID"
            print(f"{idx}. {config.name} - {status}")
            if not config.valid and config.error:
                print(f"   Error: {config.error}")
                
            if config.valid:
                valid_count += 1
            else:
                invalid_count += 1
                
        print("\nSummary:")
        print(f"- Total configurations: {len(self.configs)}")
        print(f"- Valid: {valid_count}")
        print(f"- Invalid: {invalid_count}")

    def run_sync_for_config(self, config: SyncConfig) -> None:
        """Run sync for a specific configuration."""
        if not config.valid:
            print(f"Cannot sync {config.name} due to invalid configuration: {config.error}")
            return
            
        print(f"\nStarting sync for {config.name}")
        
        # Log the sync start to database
        sync_id = -1
        if self.db_handler:
            sync_id = self.db_handler.log_sync_start(config.name)
            
        success = main.run_sync(config.path)
        
        # Log the sync result to database
        if self.db_handler and sync_id > 0:
            status_message = "Sync completed successfully" if success else "Sync failed"
            self.db_handler.log_sync_end(sync_id, success, status_message)

    def start_sync(self) -> None:
        """Start sync process for all valid configurations."""
        # Make sure main.py has access to our database handler
        main.db_handler = self.db_handler
        
        valid_configs = [config for config in self.configs if config.valid]
        
        if not valid_configs:
            print("No valid configurations found. Please check your config files.")
            return
            
        print(f"\nStarting sync process for {len(valid_configs)} configuration(s)")
        
        # Log the overall sync start to database
        overall_sync_id = -1
        if self.db_handler:
            overall_sync_id = self.db_handler.log_sync_start("multi_sync_all")
        
        for config in valid_configs:
            self.run_sync_for_config(config)
            
        # Log the overall sync completion to database
        if self.db_handler and overall_sync_id > 0:
            message = f"Completed sync for {len(valid_configs)} configuration(s)"
            self.db_handler.log_sync_end(overall_sync_id, True, message)
            
    def start_continuous_sync(self, interval: int = 300) -> None:
        """Start continuous sync process for all valid configurations."""
        # Make sure main.py has access to our database handler
        main.db_handler = self.db_handler
        
        valid_configs = [config for config in self.configs if config.valid]
        
        if not valid_configs:
            print("No valid configurations found. Please check your config files.")
            return
            
        print(f"\nStarting continuous sync process for {len(valid_configs)} configuration(s)")
        print(f"Sync interval: {interval} seconds")
        
        # Create stop events for each config
        for config in valid_configs:
            self.stop_events[config.name] = threading.Event()
        
        try:
            while True:
                # Log the overall sync cycle start to database
                overall_sync_id = -1
                if self.db_handler:
                    overall_sync_id = self.db_handler.log_sync_start("multi_sync_cycle")
                
                for config in valid_configs:
                    self.run_sync_for_config(config)
                
                # Log the overall sync cycle completion to database
                if self.db_handler and overall_sync_id > 0:
                    message = f"Completed sync cycle for {len(valid_configs)} configuration(s)"
                    self.db_handler.log_sync_end(overall_sync_id, True, message)
                
                print(f"\nWaiting {interval} seconds before next sync cycle...")
                time.sleep(interval)
                
        except KeyboardInterrupt:
            print("\nSync process stopped by user")
            
            # Close database connection
            if self.db_handler:
                self.db_handler.close()
                
            sys.exit(0)
        except Exception as e:
            print(f"\nError during continuous sync: {e}")
            
            # Log the error to database
            if self.db_handler:
                self.db_handler.log_sync_end(overall_sync_id, False, f"Error during continuous sync: {e}")
                self.db_handler.close()
                
            sys.exit(1)

def main():
    # Allow specifying custom config directory
    config_dir = 'configs'
    if len(sys.argv) > 1:
        config_dir = sys.argv[1]
        
    multi_sync = MultiSync(config_dir=config_dir)
    multi_sync.load_configs()
    multi_sync.print_status()
    
    interval = 300  # Default interval in seconds
    if len(sys.argv) > 2:
        try:
            interval = int(sys.argv[2])
        except ValueError:
            print(f"Invalid interval '{sys.argv[2]}', using default 300 seconds")
    
    # Start continuous sync
    multi_sync.start_continuous_sync(interval=interval)

if __name__ == "__main__":
    main()