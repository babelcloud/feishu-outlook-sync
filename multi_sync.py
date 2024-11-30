import os
import sys
import yaml
import time
import threading
from typing import Dict, List, Optional
from dataclasses import dataclass
from main import run_sync, run_continuous_sync

@dataclass
class SyncConfig:
    path: str
    name: str
    valid: bool
    error: Optional[str] = None

class MultiSync:
    def __init__(self, config_dir: str = 'configs'):
        self.config_dir = config_dir
        self.configs: List[SyncConfig] = []
        self.sync_threads: Dict[str, threading.Thread] = {}
        
    def validate_yaml(self, path: str) -> SyncConfig:
        """Validate YAML file format and content."""
        try:
            # Get name from filename
            name = os.path.splitext(os.path.basename(path))[0]
            
            # Validate file extension
            if not path.endswith('.yaml'):
                return SyncConfig(path, name, False, "Not a YAML file")
                
            with open(path, 'r') as file:
                config = yaml.safe_load(file)
                
            # Validate required sections
            required_sections = ['feishu', 'outlook', 'calendar_pairs']
            missing_sections = [s for s in required_sections if s not in config]
            
            if missing_sections:
                return SyncConfig(
                    path, name, False, 
                    f"Missing required sections: {', '.join(missing_sections)}"
                )
                
            # Validate subsections
            if not config['feishu'].get('app_info'):
                return SyncConfig(path, name, False, "Missing Feishu app info")
                
            if not config['outlook'].get('app_info'):
                return SyncConfig(path, name, False, "Missing Outlook app info")
                
            # Validate calendar pairs format
            if not isinstance(config['calendar_pairs'], list):
                return SyncConfig(
                    path, name, False, 
                    "calendar_pairs must be a list"
                )
                
            for pair in config['calendar_pairs']:
                if not all(k in pair for k in ['feishu', 'outlook']):
                    return SyncConfig(
                        path, name, False, 
                        "Invalid calendar pair format"
                    )
                    
            return SyncConfig(path, name, True)
            
        except yaml.YAMLError as e:
            return SyncConfig(path, name, False, f"Invalid YAML format: {str(e)}")
        except Exception as e:
            return SyncConfig(path, name, False, str(e))

    def load_configs(self) -> None:
        """Load and validate all YAML files in config directory."""
        self.configs = []
        
        # Create config directory if it doesn't exist
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            print(f"Created config directory: {self.config_dir}")
            return

        # Load and validate each YAML file
        for filename in os.listdir(self.config_dir):
            if filename.endswith('.yaml'):
                path = os.path.join(self.config_dir, filename)
                config = self.validate_yaml(path)
                self.configs.append(config)
                
        # Sort configs by name for consistent ordering
        self.configs.sort(key=lambda x: x.name)

    def print_status(self) -> None:
        """Print status of all configurations."""
        print("\nConfiguration Status:")
        print("=" * 50)
        
        if not self.configs:
            print(f"No YAML files found in {self.config_dir}")
            return
            
        for config in self.configs:
            status = "Valid" if config.valid else "Invalid"
            print(f"\nConfig: {config.name}")
            print(f"Status: {status}")
            print(f"Path: {config.path}")
            if config.error:
                print(f"Error: {config.error}")
        print("\n" + "=" * 50)

    def run_sync_for_config(self, config: SyncConfig) -> None:
        """Run continuous sync for a single configuration."""
        try:
            run_continuous_sync(config.path)
        except Exception as e:
            print(f"Error in sync thread for {config.name}: {e}")

    def start_sync(self) -> None:
        """Start sync process for all valid configurations."""
        self.load_configs()
        self.print_status()
        
        valid_configs = [c for c in self.configs if c.valid]
        
        if not valid_configs:
            print("\nNo valid configurations found")
            return
            
        print(f"\nStarting sync for {len(valid_configs)} configurations...")
        
        # Start a thread for each valid configuration
        for config in valid_configs:
            print(f"\nStarting sync thread for {config.name}")
            thread = threading.Thread(
                target=self.run_sync_for_config,
                args=(config,),
                name=f"sync_{config.name}"
            )
            thread.daemon = True
            thread.start()
            self.sync_threads[config.name] = thread
            
        try:
            # Monitor threads and report status
            while True:
                alive_threads = {
                    name: thread 
                    for name, thread in self.sync_threads.items() 
                    if thread.is_alive()
                }
                
                if not alive_threads:
                    print("\nAll sync threads have stopped")
                    break
                    
                self.sync_threads = alive_threads
                time.sleep(60)  # Check thread status every minute
                
        except KeyboardInterrupt:
            print("\nStopping all sync threads...")
            sys.exit(0)

def main():
    # Allow specifying custom config directory
    config_dir = sys.argv[1] if len(sys.argv) > 1 else 'configs'
    
    multi_sync = MultiSync(config_dir)
    multi_sync.start_sync()

if __name__ == "__main__":
    main()