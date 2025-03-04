import os
import psycopg2
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

class DBHandler:
    """Handler for database operations and logging sync activities."""
    
    def __init__(self, 
                 host: str = None, 
                 database: str = None, 
                 user: str = None, 
                 password: str = None):
        """Initialize database connection."""
        # Use environment variables if parameters not provided
        self.host = host or os.environ.get('PGHOST')
        self.database = database or os.environ.get('PGDATABASE')
        self.user = user or os.environ.get('PGUSER')
        self.password = password or os.environ.get('PGPASSWORD')
        self.conn = None
        self.cursor = None
        
    def connect(self) -> bool:
        """Establish connection to PostgreSQL database."""
        try:
            self.conn = psycopg2.connect(
                host=self.host,
                database=self.database,
                user=self.user,
                password=self.password
            )
            self.cursor = self.conn.cursor()
            print("Successfully connected to PostgreSQL database")
            self._initialize_tables()
            return True
        except Exception as e:
            print(f"Database connection error: {e}")
            return False
    
    def _initialize_tables(self) -> None:
        """Initialize database tables if they don't exist."""
        try:
            # Create sync logs table
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS sync_logs (
                    id SERIAL PRIMARY KEY,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    config_name VARCHAR(255),
                    status VARCHAR(50),
                    message TEXT
                )
            """)
            
            # Create event sync details table
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS event_syncs (
                    id SERIAL PRIMARY KEY,
                    sync_id INTEGER REFERENCES sync_logs(id),
                    calendar_pair VARCHAR(255),
                    feishu_calendar_id VARCHAR(255),
                    feishu_calendar_name VARCHAR(255),
                    outlook_calendar_id VARCHAR(255),
                    outlook_calendar_name VARCHAR(255),
                    events_synced INTEGER,
                    events_skipped INTEGER,
                    events_failed INTEGER,
                    events_deleted INTEGER,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Create table for individual event logs
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS event_logs (
                    id SERIAL PRIMARY KEY,
                    event_sync_id INTEGER REFERENCES event_syncs(id),
                    event_title VARCHAR(255),
                    event_start_time TIMESTAMP,
                    event_end_time TIMESTAMP,
                    action VARCHAR(50),
                    status VARCHAR(50),
                    error_message TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            self.conn.commit()
            print("Database tables initialized")
        except Exception as e:
            self.conn.rollback()
            print(f"Error initializing tables: {e}")
    
    def log_sync_start(self, config_name: str) -> int:
        """Log the start of a sync process and return the sync ID."""
        try:
            self.cursor.execute(
                """
                INSERT INTO sync_logs (config_name, status, message)
                VALUES (%s, %s, %s)
                RETURNING id
                """,
                (config_name, "started", f"Sync started for {config_name}")
            )
            sync_id = self.cursor.fetchone()[0]
            self.conn.commit()
            return sync_id
        except Exception as e:
            self.conn.rollback()
            print(f"Error logging sync start: {e}")
            return -1
    
    def log_sync_end(self, sync_id: int, success: bool, message: str = "") -> bool:
        """Log the end of a sync process."""
        try:
            self.cursor.execute(
                """
                UPDATE sync_logs
                SET status = %s, message = %s
                WHERE id = %s
                """,
                ("success" if success else "failed", message, sync_id)
            )
            self.conn.commit()
            return True
        except Exception as e:
            self.conn.rollback()
            print(f"Error logging sync end: {e}")
            return False
    
    def log_calendar_sync(self, 
                          sync_id: int, 
                          calendar_pair: str,
                          feishu_calendar_id: str,
                          feishu_calendar_name: str,
                          outlook_calendar_id: str,
                          outlook_calendar_name: str,
                          events_synced: int,
                          events_skipped: int,
                          events_failed: int,
                          events_deleted: int) -> int:
        """Log statistics for a calendar pair sync."""
        try:
            self.cursor.execute(
                """
                INSERT INTO event_syncs (
                    sync_id, calendar_pair, feishu_calendar_id, feishu_calendar_name,
                    outlook_calendar_id, outlook_calendar_name, events_synced,
                    events_skipped, events_failed, events_deleted
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
                """,
                (
                    sync_id, calendar_pair, feishu_calendar_id, feishu_calendar_name,
                    outlook_calendar_id, outlook_calendar_name, events_synced,
                    events_skipped, events_failed, events_deleted
                )
            )
            event_sync_id = self.cursor.fetchone()[0]
            self.conn.commit()
            return event_sync_id
        except Exception as e:
            self.conn.rollback()
            print(f"Error logging calendar sync: {e}")
            return -1
    
    def log_event_action(self, 
                         event_sync_id: int,
                         event_title: str,
                         event_start_time: datetime,
                         event_end_time: datetime,
                         action: str,
                         status: str,
                         error_message: str = "") -> bool:
        """Log an action performed on an individual event."""
        try:
            self.cursor.execute(
                """
                INSERT INTO event_logs (
                    event_sync_id, event_title, event_start_time, event_end_time,
                    action, status, error_message
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    event_sync_id, event_title, event_start_time, event_end_time,
                    action, status, error_message
                )
            )
            self.conn.commit()
            return True
        except Exception as e:
            self.conn.rollback()
            print(f"Error logging event action: {e}")
            return False
    
    def get_sync_statistics(self, limit: int = 10) -> List[Dict[str, Any]]:
        """Get statistics for recent syncs."""
        try:
            self.cursor.execute(
                """
                SELECT 
                    sl.id, sl.timestamp, sl.config_name, sl.status,
                    SUM(es.events_synced) as total_synced,
                    SUM(es.events_skipped) as total_skipped,
                    SUM(es.events_failed) as total_failed,
                    SUM(es.events_deleted) as total_deleted
                FROM 
                    sync_logs sl
                LEFT JOIN 
                    event_syncs es ON sl.id = es.sync_id
                GROUP BY 
                    sl.id, sl.timestamp, sl.config_name, sl.status
                ORDER BY 
                    sl.timestamp DESC
                LIMIT %s
                """,
                (limit,)
            )
            columns = [desc[0] for desc in self.cursor.description]
            results = []
            for row in self.cursor.fetchall():
                results.append(dict(zip(columns, row)))
            return results
        except Exception as e:
            print(f"Error getting sync statistics: {e}")
            return []
    
    def close(self) -> None:
        """Close database connection."""
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close() 