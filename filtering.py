"""
Contains functions to filter data based on user-defined criteria. 
It is used to narrow down the dataset to relevant subsets, which is essential for the generation of accurate and specific labels. 
Filtering is performed on various attributes like box values and criteria columns.
"""

# filtering.py

import logging

def filter_df(selected_criteria, full_df, criteria_columns):
    '''Filter a DataFrame based on selected criteria (series or box).'''
    try:
        filtered_rows = full_df[full_df[criteria_columns].isin(selected_criteria).any(axis=1)]
        return filtered_rows.drop_duplicates()
    except Exception as e:
        logging.error(f"Error filtering DataFrame: {str(e)}")
        return None

def filter_df_by_box_values(df, box_values, column_name='BOX', add_prefix=False):
    if add_prefix:
        # Adjust the box_values to match the "Box " prefix format
        adjusted_box_values = [f"Box {value}" for value in box_values]
    else:
        # Handle cases like '10A' or '10 (Oversize)' by ensuring we compare strings
        adjusted_box_values = [str(value) for value in box_values]
    return df[df[column_name].isin(adjusted_box_values)]