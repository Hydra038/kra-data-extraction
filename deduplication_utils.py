import pandas as pd
import logging

def deduplicate_dataframe(df: pd.DataFrame, subset_cols: list = None) -> pd.DataFrame:
    """
    Deduplicates the DataFrame based on a combination of key fields.
    If multiple records are exact matches, the most recently extracted one is kept.
    """
    if df.empty:
        return df

    if subset_cols is None:
        subset_cols = ['pin', 'date', 'preAmount', 'taxpayerName']

    initial_count = len(df)
    # Sort by 'date_extracted' descending to keep the LATEST record when duplicates are found
    if 'date_extracted' in df.columns:
        df_sorted = df.sort_values(by='date_extracted', ascending=False)
    else:
        df_sorted = df.copy()

    # Drop duplicates based on key fields, keeping the 'first' (latest)
    deduplicated_df = df_sorted.drop_duplicates(
        subset=subset_cols,
        keep='first'
    )

    duplicates_removed = initial_count - len(deduplicated_df)
    if duplicates_removed > 0:
        logging.getLogger(__name__).info(f"Deduplication completed. {duplicates_removed} duplicates removed.")
    return deduplicated_df

def compare_extraction_methods(results1, results2, method1_name, method2_name):
    return {}
