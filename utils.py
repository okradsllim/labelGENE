# utils.py

"""
Module containing utility functions.

This module includes various utility functions used throughout the script, such as sorting functions,
file management operations, and data transformation helpers. These functions provide reusable functionality
to support the main processes of the script.
"""

import os
import shutil
from datetime import datetime, timedelta
import logging
import re
import pandas as pd


def box_sort_order(box):
        # because there might be alphanumeric box number for e.g. '10A'
        match = re.search(r'\d+', box)
        # because we want (the often few) alphanumerics to appear before their immediate numeric counterparts, for e.g. '10A' , '10'
        # also ensures that math operations can be easily performed on box number ranges without worrying about a possible alphanumberic outer range
        if match:
            return (1, int(match.group())) 
        else:
            return (0, box)
        
def custom_sort_key(option):
    ''' Custom sorting function for box selection display to sort by number first, then text. '''
    matches = re.match(r'(\d+)(.*)', option)
    if matches:
        number_part = int(matches.group(1))
        text_part = matches.group(2)
        return (number_part, text_part)
    return (0, option)  # Default for items without a leading number

def move_recent_ead_files(working_directory):
    '''this copies recently downloaded EAD files from ASpace so that user skips the step of having to go to Downloads folder'''
    downloads_folder = os.path.join("C:\\", "Users", os.getlogin(), "Downloads")
    file_extension = ".xml"
    one_day_ago = datetime.now() - timedelta(days=1)

    try:
        # First, filter out only recent files
        recent_files = [f for f in os.listdir(downloads_folder) if f.endswith(file_extension) and
                        datetime.fromtimestamp(os.path.getmtime(os.path.join(downloads_folder, f))) > one_day_ago]

        # Log the count of recent files
        logging.info(f"Total recent XML files in Downloads folder: {len(recent_files)}")

        # Process only the recent files
        for filename in recent_files:
            filepath = os.path.join(downloads_folder, filename)
            destination = os.path.join(working_directory, filename)
            shutil.copy(filepath, destination)
            logging.info(f"Copied {filename} to {working_directory}")

    except Exception as e:
        logging.error(f"Error copying files to current directory: {str(e)}")
              
def prepend_or_fill(column_name, x, idx):
    prefix = "Box " if column_name == 'BOX' else "Folder "
    if pd.notnull(x):  # If the cell has a value, it must be INTEGER. If Box, 
        return prefix + str(x)
    else:  # If the cell is empty
        return prefix + str(idx + 1)