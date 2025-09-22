#!/usr/bin/env python3
"""
Database Migration Script
Converts old column format to new camelCase format and removes duplicates
"""

import pandas as pd
import os
import shutil
from datetime import datetime

def migrate_database():
    """Migrate database from old format to new camelCase format"""
    
    db_path = 'kra_master_database.xlsx'
    backup_path = f'kra_master_database_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    if not os.path.exists(db_path):
        print("❌ Database file not found")
        return False
    
    try:
        # Create backup
        shutil.copy2(db_path, backup_path)
        print(f"✅ Backup created: {backup_path}")
        
        # Read existing database
        df = pd.read_excel(db_path)
        print(f"📊 Original database: {len(df)} records, {len(df.columns)} columns")
        print(f"🔍 Original columns: {list(df.columns)}")
        
        # Create new DataFrame with camelCase columns only
        new_df = pd.DataFrame()
        
        # Map old columns to new columns and merge data
        column_mapping = {
            'date': ['Date', 'date'],
            'pin': ['PIN', 'pin'],
            'taxpayerName': ['Taxpayer_Name', 'taxpayerName'],
            'preAmount': ['Pre-Amount', 'preAmount'],
            'finalAmount': ['finalAmount'],  # New field
            'year': ['Year', 'year'],
            'officerName': ['Officer_Name', 'officerName'],
            'station': ['Station', 'station']
        }
        
        # Process each target column
        for new_col, old_cols in column_mapping.items():
            # Try to get data from any of the old column names
            for old_col in old_cols:
                if old_col in df.columns:
                    existing_data = df[old_col].dropna()
                    if not existing_data.empty:
                        if new_col in new_df.columns:
                            # Merge with existing data
                            new_df[new_col] = new_df[new_col].fillna(existing_data)
                        else:
                            new_df[new_col] = existing_data
                        break
            
            # If no data found, create empty column
            if new_col not in new_df.columns:
                new_df[new_col] = ''
        
        # Add metadata columns
        if 'date_extracted' in df.columns:
            new_df['date_extracted'] = df['date_extracted']
        else:
            new_df['date_extracted'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
        if 'source_app' in df.columns:
            new_df['source_app'] = df['source_app']
        else:
            new_df['source_app'] = 'migration_script'
            
        if 'record_id' in df.columns:
            new_df['record_id'] = df['record_id']
        else:
            new_df['record_id'] = range(1, len(new_df) + 1)
        
        # Remove completely empty rows
        new_df = new_df.dropna(how='all')
        
        # Ensure finalAmount is properly initialized
        new_df['finalAmount'] = new_df['finalAmount'].fillna('')
        
        print(f"🔄 Migrated database: {len(new_df)} records, {len(new_df.columns)} columns")
        print(f"✅ New columns: {list(new_df.columns)}")
        
        # Save migrated database
        new_df.to_excel(db_path, index=False)
        print(f"💾 Database migrated successfully!")
        
        # Show sample
        if not new_df.empty:
            print("\n📋 Sample migrated data:")
            print(new_df.head(2))
        
        return True
        
    except Exception as e:
        print(f"❌ Migration failed: {e}")
        # Restore backup if migration failed
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, db_path)
            print("🔄 Database restored from backup")
        return False

if __name__ == "__main__":
    migrate_database()