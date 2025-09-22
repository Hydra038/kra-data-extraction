#!/usr/bin/env python3
"""
KRA Data Deduplication Utilities
Handles duplicate detection and intelligent merging of extraction results
"""

import pandas as pd
import hashlib
from typing import List, Dict, Any
import logging

def create_record_hash(record: Dict[str, Any]) -> str:
    """
    Create a unique hash for a record based on key identifying fields.
    
    Args:
        record: Dictionary containing extraction results
        
    Returns:
        str: MD5 hash of key identifying fields
    """
    # Use key fields that should be unique per document (core fields only)
    key_fields = ['pin', 'date', 'taxpayerName']
    
    # Create a string from key fields (handle missing values)
    key_string = ""
    for field in key_fields:
        value = str(record.get(field, '')).strip().upper()
        key_string += f"{field}:{value}|"
    
    # Create MD5 hash
    return hashlib.md5(key_string.encode()).hexdigest()

def calculate_record_score(record: Dict[str, Any]) -> float:
    """
    Calculate a quality score for a record based on completeness and accuracy.
    
    Args:
        record: Dictionary containing extraction results
        
    Returns:
        float: Quality score (0-100)
    """
    score = 0.0
    max_score = 100.0
    
    # Core fields and their weights (8 fields)
    field_weights = {
        'date': 16,
        'pin': 20,
        'taxpayerName': 18,
        'preAmount': 12,
        'finalAmount': 8,
        'year': 12,
        'officerName': 8,
        'station': 6
    }
    
    for field, weight in field_weights.items():
        value = record.get(field, '')
        if value and str(value).strip():
            # Additional scoring based on data quality
            if field == 'pin' and len(str(value)) == 11:  # Valid PIN format
                score += weight * 1.2  # Bonus for valid format
            elif field == 'year' and str(value).isdigit() and 2020 <= int(value) <= 2030:
                score += weight * 1.1  # Bonus for valid year
            elif field in ['date', 'taxpayerName', 'officerName', 'station'] and len(str(value)) > 3:
                score += weight  # Standard score for meaningful content
            else:
                score += weight * 0.8  # Reduced score for questionable data
    
    return min(score, max_score)

def find_duplicates(df: pd.DataFrame) -> Dict[str, List[int]]:
    """
    Find duplicate records in a DataFrame based on record hashes.
    
    Args:
        df: DataFrame containing extraction results
        
    Returns:
        Dict mapping hash to list of row indices that share that hash
    """
    hash_groups = {}
    
    for idx, row in df.iterrows():
        record_hash = create_record_hash(row.to_dict())
        
        if record_hash not in hash_groups:
            hash_groups[record_hash] = []
        hash_groups[record_hash].append(idx)
    
    # Return only groups with multiple records (duplicates)
    return {h: indices for h, indices in hash_groups.items() if len(indices) > 1}

def merge_duplicate_records(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Intelligently merge duplicate records, keeping the best data from each.
    
    Args:
        records: List of duplicate records to merge
        
    Returns:
        Dict: Merged record with best available data
    """
    if not records:
        return {}
    
    if len(records) == 1:
        return records[0]
    
    # Find the record with the highest quality score
    scored_records = [(calculate_record_score(record), record) for record in records]
    scored_records.sort(reverse=True, key=lambda x: x[0])
    
    best_record = scored_records[0][1].copy()
    
    # Merge additional data from other records if better
    merge_fields = ['date', 'pin', 'taxpayerName', 'preAmount', 'finalAmount', 'year', 'officerName', 'station']
    
    for _, record in scored_records[1:]:
        for field in merge_fields:
            best_value = best_record.get(field, '')
            candidate_value = record.get(field, '')
            
            # Use candidate value if it's better (longer, more complete, etc.)
            if not best_value and candidate_value:
                best_record[field] = candidate_value
            elif candidate_value and len(str(candidate_value)) > len(str(best_value)):
                # Keep longer values if they seem more complete
                if field in ['Taxpayer_Name', 'Officer_Name']:
                    best_record[field] = candidate_value
    
    # Add merge metadata
    best_record['Merged_From_Count'] = len(records)
    best_record['Merge_Sources'] = ', '.join([r.get('Extraction_Method', 'Unknown') for r in records])
    best_record['Best_Score'] = scored_records[0][0]
    
    return best_record

def deduplicate_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicates from a DataFrame using intelligent merging.
    
    Args:
        df: DataFrame containing extraction results
        
    Returns:
        DataFrame with duplicates removed and merged
    """
    if df.empty:
        return df
    
    # Find duplicate groups
    duplicate_groups = find_duplicates(df)
    
    if not duplicate_groups:
        logging.info("No duplicates found")
        return df
    
    logging.info(f"Found {len(duplicate_groups)} duplicate groups affecting {sum(len(indices) for indices in duplicate_groups.values())} records")
    
    # Create new DataFrame without duplicates
    deduplicated_records = []
    processed_indices = set()
    
    # Process duplicate groups
    for hash_value, indices in duplicate_groups.items():
        if any(idx in processed_indices for idx in indices):
            continue
            
        # Get records for this group
        duplicate_records = [df.iloc[idx].to_dict() for idx in indices]
        
        # Merge the duplicates
        merged_record = merge_duplicate_records(duplicate_records)
        deduplicated_records.append(merged_record)
        
        # Mark these indices as processed
        processed_indices.update(indices)
    
    # Add non-duplicate records
    for idx, row in df.iterrows():
        if idx not in processed_indices:
            deduplicated_records.append(row.to_dict())
    
    # Create new DataFrame
    result_df = pd.DataFrame(deduplicated_records)
    
    logging.info(f"Deduplication complete: {len(df)} -> {len(result_df)} records")
    
    return result_df

def compare_extraction_methods(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Compare the performance of different extraction methods.
    
    Args:
        df: DataFrame containing extraction results
        
    Returns:
        Dict containing comparison statistics
    """
    if 'Extraction_Method' not in df.columns:
        return {"error": "No extraction method information available"}
    
    methods = df['Extraction_Method'].unique()
    comparison = {}
    
    for method in methods:
        method_df = df[df['Extraction_Method'] == method]
        
        total_records = len(method_df)
        successful_records = len(method_df[method_df['Processing_Status'] == 'Success'])
        
        if total_records > 0:
            success_rate = (successful_records / total_records) * 100
            avg_score = method_df.apply(lambda row: calculate_record_score(row.to_dict()), axis=1).mean()
            avg_fields = method_df['Fields_Found'].mean() if 'Fields_Found' in method_df.columns else 0
        else:
            success_rate = 0
            avg_score = 0
            avg_fields = 0
        
        comparison[method] = {
            'total_records': total_records,
            'successful_records': successful_records,
            'success_rate': success_rate,
            'avg_quality_score': avg_score,
            'avg_fields_found': avg_fields
        }
    
    return comparison

if __name__ == "__main__":
    # Test the deduplication system
    sample_data = [
        {
            'File_Name': 'test.pdf',
            'date': '29TH AUGUST, 2025',
            'pin': 'A001126762Z',
            'taxpayerName': 'Peter Kimutai Telengech',
            'preAmount': '14,769.50',
            'finalAmount': '',
            'year': '2024',
            'officerName': 'Franciscar Nyangweta',
            'station': 'ELDORET',
            'Processing_Status': 'Success',
            'Fields_Found': 8,
            'Extraction_Method': 'multi_format_extractor'
        },
        {
            'File_Name': 'test.pdf',
            'date': '29TH AUGUST, 2025',
            'pin': 'A001126762Z',
            'taxpayerName': 'Peter Kimutai Telengech',
            'preAmount': '14,769.50',
            'finalAmount': '',
            'year': '2024',
            'officerName': 'Franciscar Nyangweta',
            'station': 'ELDORET',
            'Processing_Status': 'Success',
            'Fields_Found': 8,
            'Extraction_Method': 'app_extractor'
        }
    ]
    
    df = pd.DataFrame(sample_data)
    print("Original records:", len(df))
    
    deduplicated_df = deduplicate_dataframe(df)
    print("After deduplication:", len(deduplicated_df))
    
    print("\nMerged record:")
    print(deduplicated_df.iloc[0].to_dict())