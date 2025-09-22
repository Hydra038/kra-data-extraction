#!/usr/bin/env python3
"""
KRA Excel Database Utilities
Handles persistent storage of all extraction results in a master Excel file
"""

import pandas as pd
import os
from datetime import datetime
import logging
from typing import Tuple, Optional
from deduplication_utils import deduplicate_dataframe

# Configure logging
logger = logging.getLogger(__name__)

# Database configuration
DATABASE_FILE = "kra_master_database.xlsx"
BACKUP_FILE = "kra_master_database_backup.xlsx"

def get_database_path() -> str:
    """Get the full path to the database file"""
    # Store database in the application directory (works for both local and production)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # For production deployment, ensure we have write permissions
    if os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('RENDER'):
        # Production environment - use app directory
        db_path = os.path.join(script_dir, DATABASE_FILE)
    else:
        # Local development
        db_path = os.path.join(script_dir, DATABASE_FILE)
    
    return db_path

def get_backup_path() -> str:
    """Get the full path to the backup database file"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, BACKUP_FILE)

def create_backup() -> bool:
    """Create a backup of the current database"""
    try:
        db_path = get_database_path()
        backup_path = get_backup_path()
        
        if os.path.exists(db_path):
            # Read and save as backup
            df = pd.read_excel(db_path)
            df.to_excel(backup_path, index=False)
            logger.info(f"Database backup created: {backup_path}")
            return True
        return False
    except Exception as e:
        logger.error(f"Failed to create backup: {str(e)}")
        return False

def load_existing_database() -> pd.DataFrame:
    """
    Load existing database or return empty DataFrame if not found
    
    Returns:
        pd.DataFrame: Existing database records or empty DataFrame
    """
    db_path = get_database_path()
    
    try:
        if os.path.exists(db_path):
            # Explicitly load the correct sheet
            df = pd.read_excel(db_path, sheet_name='KRA_Database')
            logger.info(f"Loaded existing database with {len(df)} records")
            return df
        else:
            logger.info("No existing database found, will create new one")
            return pd.DataFrame()
    except Exception as e:
        logger.error(f"Error loading database: {str(e)}")
        return pd.DataFrame()
# --- Add missing utility functions for stats and export ---
import openpyxl
from typing import Optional
from datetime import datetime

def get_database_stats() -> tuple:
    """
    Returns basic statistics from the database.
    Returns:
        Tuple: (total_records, last_updated_date, unique_taxpayers, unique_stations)
    """
    try:
        df = load_existing_database()
        if df.empty:
            return 0, None, 0, 0
        total_records = len(df)
        if 'date_extracted' in df.columns:
            df['date_extracted'] = pd.to_datetime(df['date_extracted'], errors='coerce')
            last_updated = df['date_extracted'].max()
        else:
            last_updated = None
        unique_taxpayers = df['pin'].nunique() if 'pin' in df.columns else 0
        unique_stations = df.get('station', pd.Series([])).nunique()
        return total_records, last_updated, unique_taxpayers, unique_stations
    except Exception as e:
        logger.error(f"Error getting database stats: {str(e)}")
        return 0, None, 0, 0

def export_database_to_excel() -> Optional[bytes]:
    """
    Loads the entire database file content as bytes for Streamlit download.
    """
    try:
        db_path = get_database_path()
        if not os.path.exists(db_path):
            logger.warning("Attempted to export but database file does not exist.")
            return None
        with open(db_path, "rb") as f:
            excel_bytes = f.read()
        return excel_bytes
    except Exception as e:
        logger.error(f"Error exporting database to Excel: {str(e)}")
        return None

def save_to_database(new_data_df: pd.DataFrame, source_app: str = "unknown") -> Tuple[int, int, int]:
    """
    Save new extraction results to the master database with automatic deduplication
    
    Args:
        new_data_df: DataFrame with new extraction results
        source_app: Name of the app that generated the data
        
    Returns:
        Tuple[int, int, int]: (total_records, new_records, duplicates_removed)
    """
    try:
        if new_data_df.empty:
            logger.warning("No data to save to database")
            return 0, 0, 0
        
        db_path = get_database_path()
        
        # Create backup before making changes
        create_backup()
        
        # Load existing database
        existing_df = load_existing_database()
        
        # Add metadata to new data
        new_data_with_meta = new_data_df.copy()
        new_data_with_meta['date_extracted'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        new_data_with_meta['source_app'] = source_app
        new_data_with_meta['record_id'] = range(
            len(existing_df) + 1, 
            len(existing_df) + len(new_data_with_meta) + 1
        )
        
        # Combine with existing data
        if not existing_df.empty:
            # Ensure columns match
            for col in new_data_with_meta.columns:
                if col not in existing_df.columns:
                    existing_df[col] = ''
            for col in existing_df.columns:
                if col not in new_data_with_meta.columns:
                    new_data_with_meta[col] = ''
            
            combined_df = pd.concat([existing_df, new_data_with_meta], ignore_index=True)
        else:
            combined_df = new_data_with_meta
        
        # Apply deduplication
        original_count = len(combined_df)
        deduplicated_df = deduplicate_dataframe(combined_df)
        final_count = len(deduplicated_df)
        duplicates_removed = original_count - final_count
        
        # Save to Excel with proper formatting
        with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
            # Main data sheet
            deduplicated_df.to_excel(writer, sheet_name='KRA_Database', index=False)
            
            # Summary sheet
            summary_data = {
                'Metric': [
                    'Total Records',
                    'Last Updated',
                    'New Records Added (This Session)',
                    'Duplicates Removed (This Session)',
                    'Date Range (Earliest)',
                    'Date Range (Latest)',
                    'Unique Taxpayers',
                    'Unique Stations'
                ],
                'Value': [
                    len(deduplicated_df),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    len(new_data_df),
                    duplicates_removed,
                    deduplicated_df['date'].min() if not deduplicated_df.empty and 'date' in deduplicated_df.columns else 'N/A',
                    deduplicated_df['date'].max() if not deduplicated_df.empty and 'date' in deduplicated_df.columns else 'N/A',
                    deduplicated_df['taxpayerName'].nunique() if not deduplicated_df.empty and 'taxpayerName' in deduplicated_df.columns else 0,
                    deduplicated_df['station'].nunique() if not deduplicated_df.empty and 'station' in deduplicated_df.columns else 0
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Database_Summary', index=False)
        
        logger.info(f"Database updated: {final_count} total records, {len(new_data_df)} new, {duplicates_removed} duplicates removed")
        
        return final_count, len(new_data_df), duplicates_removed
        
    except Exception as e:
        logger.error(f"Error saving to database: {str(e)}")
        return 0, 0, 0

def get_full_database() -> pd.DataFrame:
    """
    Get complete database for download
    
    Returns:
        pd.DataFrame: Complete database records
    """
    try:
        return load_existing_database()
    except Exception as e:
        logger.error(f"Error retrieving full database: {str(e)}")
        return pd.DataFrame()

def get_database_stats() -> dict:
    """
    Get database statistics for display
    
    Returns:
        dict: Database statistics
    """
    try:
        df = load_existing_database()
        
        if df.empty:
            return {
                'total_records': 0,
                'last_updated': 'Never',
                'unique_taxpayers': 0,
                'unique_stations': 0,
                'date_range': 'No data'
            }
        
        stats = {
            'total_records': len(df),
            'last_updated': df['date_extracted'].max() if 'date_extracted' in df.columns else 'Unknown',
            'unique_taxpayers': df['taxpayerName'].nunique() if 'taxpayerName' in df.columns else 0,
            'unique_stations': df['station'].nunique() if 'station' in df.columns else 0,
            'date_range': f"{df['date'].min()} to {df['date'].max()}" if 'date' in df.columns else 'Unknown'
        }
        
        return stats
        
    except Exception as e:
        logger.error(f"Error getting database stats: {str(e)}")
        return {
            'total_records': 0,
            'last_updated': 'Error',
            'unique_taxpayers': 0,
            'unique_stations': 0,
            'date_range': 'Error'
        }

def export_database_to_excel() -> Optional[bytes]:
    """
    Export the complete database as Excel bytes for download
    
    Returns:
        bytes: Excel file content or None if error
    """
    try:
        df = get_full_database()
        
        if df.empty:
            return None
        
        # Create Excel file in memory
        from io import BytesIO
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Main data
            df.to_excel(writer, sheet_name='KRA_Complete_Database', index=False)
            
            # Summary statistics
            stats = get_database_stats()
            summary_data = {
                'Metric': list(stats.keys()),
                'Value': list(stats.values())
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Database_Summary', index=False)
        
        return output.getvalue()
        
    except Exception as e:
        logger.error(f"Error exporting database: {str(e)}")
        return None

def clear_database() -> bool:
    """
    Clear the database (for testing purposes)
    
    Returns:
        bool: True if successful
    """
    try:
        db_path = get_database_path()
        if os.path.exists(db_path):
            create_backup()  # Backup before clearing
            os.remove(db_path)
            logger.info("Database cleared successfully")
            return True
        return True  # Already cleared
    except Exception as e:
        logger.error(f"Error clearing database: {str(e)}")
        return False

if __name__ == "__main__":
    # Test the database functions
    print("Testing KRA Excel Database...")
    
    # Test data
    test_data = pd.DataFrame({
        'date': ['2024-01-15', '2024-01-16'],
        'pin': ['A123456789X', 'B987654321Y'],
        'taxpayerName': ['John Doe', 'Jane Smith'],
        'preAmount': ['14,769.50', '25,432.75'],
        'finalAmount': ['', ''],
        'year': ['2023', '2023'],
        'officerName': ['Officer A', 'Officer B'],
        'station': ['Station 1', 'Station 2']
    })
    
    # Test save
    total, new, dupes = save_to_database(test_data, "test_app")
    print(f"Saved: {total} total, {new} new, {dupes} duplicates removed")
    
    # Test stats
    stats = get_database_stats()
    print(f"Database stats: {stats}")
    
    # Test export
    excel_data = export_database_to_excel()
    print(f"Export successful: {excel_data is not None}")