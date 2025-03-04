#!/usr/bin/env python3
import os
import sys
from db_handler import DBHandler
from dotenv import load_dotenv
from tabulate import tabulate as tabulate_fn

# Load environment variables
load_dotenv()

def query_sync_logs(db):
    """Query and display all sync logs."""
    print("\n=== SYNC LOGS ===")
    try:
        db.cursor.execute("SELECT * FROM sync_logs ORDER BY timestamp DESC")
        columns = [desc[0] for desc in db.cursor.description]
        rows = db.cursor.fetchall()
        
        if not rows:
            print("No sync logs found.")
            return
            
        print(tabulate_fn(rows, headers=columns, tablefmt="grid"))
    except Exception as e:
        print(f"Error querying sync logs: {e}")

def query_event_syncs(db):
    """Query and display all event syncs."""
    print("\n=== EVENT SYNCS ===")
    try:
        db.cursor.execute("""
            SELECT 
                es.id, 
                es.sync_id, 
                sl.config_name,
                es.calendar_pair, 
                es.feishu_calendar_name, 
                es.outlook_calendar_name,
                es.events_synced, 
                es.events_skipped, 
                es.events_failed, 
                es.events_deleted,
                es.timestamp
            FROM event_syncs es
            JOIN sync_logs sl ON es.sync_id = sl.id
            ORDER BY es.timestamp DESC
        """)
        columns = [desc[0] for desc in db.cursor.description]
        rows = db.cursor.fetchall()
        
        if not rows:
            print("No event syncs found.")
            return
            
        print(tabulate_fn(rows, headers=columns, tablefmt="grid"))
    except Exception as e:
        print(f"Error querying event syncs: {e}")

def query_event_logs(db):
    """Query and display all event logs."""
    print("\n=== EVENT LOGS ===")
    try:
        db.cursor.execute("""
            SELECT 
                el.id, 
                el.event_sync_id, 
                es.calendar_pair,
                el.event_title, 
                el.event_start_time, 
                el.event_end_time,
                el.action, 
                el.status, 
                el.error_message,
                el.timestamp
            FROM event_logs el
            JOIN event_syncs es ON el.event_sync_id = es.id
            ORDER BY el.timestamp DESC
        """)
        columns = [desc[0] for desc in db.cursor.description]
        rows = db.cursor.fetchall()
        
        if not rows:
            print("No event logs found.")
            return
            
        print(tabulate_fn(rows, headers=columns, tablefmt="grid"))
    except Exception as e:
        print(f"Error querying event logs: {e}")

def main():
    """Query and display database tables."""
    print("=== DATABASE QUERY TOOL ===\n")
    
    # Connect to database
    db = DBHandler()
    if not db.connect():
        print("❌ Failed to connect to database!")
        return False
    
    print("✅ Connected to database")
    
    # Query tables
    query_sync_logs(db)
    query_event_syncs(db)
    query_event_logs(db)
    
    # Close connection
    db.close()
    
    return True

if __name__ == "__main__":
    try:
        # First install tabulate if not already installed
        from tabulate import tabulate
    except ImportError:
        print("Installing tabulate package...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tabulate"])
        from tabulate import tabulate
        
    success = main()
    sys.exit(0 if success else 1) 