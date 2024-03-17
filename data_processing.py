"""
Provides the logic for transforming and organizing data extracted from XML files. 
This includes determining whether a node in the XML is a terminal node, handling explicit and implicit folder numbering, and processing series and box selections. 
This module is key in preparing the data for the final output, such as mail merge operations and label generation.
"""

import os
import logging
from lxml import etree as ET
import re
import datetime 

from data_extraction import extract_box_number, extract_folder_date, extract_base_folder_title, extract_ancestor_data
from user_interaction import display_options, parse_user_input
from filtering import filter_df, filter_df_by_box_values
from utils import custom_sort_key

def is_terminal_node(node):
    """Determines if a node is a terminal node by checking its children."""
    for child in node:
        tag = ET.QName(child.tag).localname
        if re.match(r'c\d{0,2}$|^c$', tag): 
            return False
    return True

def has_explicit_folder_numbering(did_element, containers, ancestor_data, folder_df, namespaces):
    """populates df when folders are explicitly numbered
    This function supplies folder numbers as string/text
    ancester_data is set to None because some terminal c nodes representing file level description are no series yet have no ancestor c nodes"""
    
    folder_container = next(elem for elem in containers if elem.attrib.get('type', '').lower() == 'folder')
    folder_text = folder_container.text.lower()
    box_number = extract_box_number(did_element, namespaces)
    container_element = did_element.find('./ns:container', namespaces=namespaces)  # find the container element
    container_type = container_element.attrib.get('altrender', None) if container_element is not None else None  # get altrender attribute, if it exists
    base_title = extract_base_folder_title(did_element, namespaces)
    date = extract_folder_date(did_element, namespaces)

    ancestor_values = ancestor_data
    ancestor_values += [None] * (5 - len(ancestor_values)) # '5' for the 5 <cxx> ancestor columns in folder_df
    
    start, end = None, None
    # if a range of folders
    if '-' in folder_text:
        # Split and filter non-numeric characters
        start, end = folder_text.split('-')
        start = re.sub(r'\D', '', start)
        end = re.sub(r'\D', '', end)

        # Convert to integers
        start, end = int(start), int(end)
        
        for i in range(start, end + 1):
            folder_title = f"{base_title} [{i - start + 1} of {end - start + 1}]"
            df_row = [folder_df['COLLECTION'][0], folder_df['CALL_NO.'][0], box_number, str(i), container_type] + ancestor_values + [folder_title, date]
            folder_df.loc[len(folder_df)] = df_row
    else:
        folder_number = folder_text
        df_row = [folder_df['COLLECTION'][0], folder_df['CALL_NO.'][0], box_number, folder_number, container_type] + ancestor_values + [base_title, date]
        folder_df.loc[len(folder_df)] = df_row

def has_implicit_folder_numbering(did_element, ancestor_data, folder_df, namespaces):
    """ populates df row when either folders are not numbered or 'folder(s)' is not mentioned at all.
    The function does not supply folder numbers, hence "None" at idx 3 in df_row population
    I've seen a situation where there's more than 2 <physdesc> inside one terminal node "Hello Henri Chopin!"
    But anyways, that would rarely be a problem because it'll most likely be because it wouldn't be about physical folders, perhaps intangible discrete items
    ancester_data is set to None because some terminal c nodes representing file level description are no series yet have no ancestor c nodes"""
    
    box_number = extract_box_number(did_element, namespaces)
    container_element = did_element.find('./ns:container', namespaces=namespaces)  # find the container element
    container_type = container_element.attrib.get('altrender', None) if container_element is not None else None  # get altrender attribute, if it exists
    base_title = extract_base_folder_title(did_element, namespaces)
    date = extract_folder_date(did_element, namespaces)

    ancestor_values = ancestor_data
    ancestor_values += [None] * (5 - len(ancestor_values)) # '5' for the 5 <cxx> ancestor columns in folder_df
    
    physdesc_elements = did_element.findall('./ns:physdesc/ns:extent', namespaces=namespaces)
    folder_count = None
    for extent in physdesc_elements:
        extent_text = extent.text
        integer_match = re.search(r'\b[1-9]\d*\b', extent_text) # Modified regex to exclude zero
        if integer_match:
            folder_count = int(integer_match.group())
            break  # Stop after finding the first valid integer

    # Use folder_count to populate df_row
    if folder_count is not None:
        if folder_count != 1:
            for i in range(1, folder_count + 1):
                folder_title = f"{base_title} [{i} of {folder_count}]"
                df_row = [folder_df['COLLECTION'][0], folder_df['CALL_NO.'][0], box_number, None, container_type] + ancestor_values + [folder_title, date]
                folder_df.loc[len(folder_df)] = df_row
        else:
            df_row = [folder_df['COLLECTION'][0], folder_df['CALL_NO.'][0], box_number, None, container_type] + ancestor_values + [base_title, date]
            folder_df.loc[len(folder_df)] = df_row
    else:
        # Handle the case where no valid folder count is found
        df_row = [folder_df['COLLECTION'][0], folder_df['CALL_NO.'][0], box_number, None, container_type] + ancestor_values + [base_title, date]
        folder_df.loc[len(folder_df)] = df_row

def process_series_selection(folder_df, box_df, working_directory, collection_name, call_number):
    series_data = folder_df['C01_ANCESTOR'].unique()

    # Check if NaN values are present
    unknown_series_present = folder_df['C01_ANCESTOR'].isna().any()

    # Custom sort function for series data
    def sort_key(series_name):
        
        if series_name is None:# Handle None values first
            return (2, None)  # 2 ensures None values are sorted to the end

        # Check if the series name matches the date pattern
        date_match = re.search(r'(\w+)\s+(\d{4})\s+acquisition', series_name)
        if date_match:
            # Extract and convert date to datetime object for sorting
            month, year = date_match.groups()
            date = datetime.datetime.strptime(f'{month} {year}', '%B %Y')
            return (1, date)  # 1 ensures date values are sorted after standard values
        else:
            # Standard series name (no date), sort alphabetically
            return (0, series_name.lower())  # 0 ensures standard values are sorted first

    # Sort series data using the custom sort function
    ordered_series = sorted(series_data, key=sort_key)

    # Append 'Unknown series' only if NaN values were present
    if unknown_series_present:
        ordered_series.append('Unknown series (CAUTION: choosing this might cause unexpected behavior in the program)')

    display_options(ordered_series, "series")

    while True:
        try:
            user_input_for_series = input("Select series by individual numbers, range, or a combination (e.g., '1', '2-3', '4, 5-6') or type 'q' to quit: \n\n")
            logging.info(f"user selects series: {user_input_for_series}")
            if user_input_for_series.lower() == 'q':
                print("\nExiting series selection...")
                return None, None

            selected_series_names = parse_user_input(user_input_for_series, ordered_series)
            
            print(f"\nYou selected: \n")
            for series in selected_series_names:
                        print(series)

            if selected_series_names:
                filtered_folder_df_by_series = filter_df(selected_series_names, folder_df, ['C01_ANCESTOR'])
                filtered_box_df_by_series = filter_df(selected_series_names, box_df, ["FIRST_C01_SERIES", "SECOND_C01_SERIES", "THIRD_C01_SERIES", "FOURTH_C01_SERIES", "FIFTH_C01_SERIES"])

                filtered_folder_df_by_series_path = os.path.join(working_directory, f"{collection_name}_{call_number}_folders_by_series_specified.xlsx")
                filtered_box_df_by_series_path = os.path.join(working_directory, f"{collection_name}_{call_number}_boxes_by_series_specified.xlsx")
                filtered_folder_df_by_series.to_excel(filtered_folder_df_by_series_path, index=False)
                filtered_box_df_by_series.to_excel(filtered_box_df_by_series_path, index=False)

                return filtered_folder_df_by_series_path, filtered_box_df_by_series_path
            else:
                print("\nNo valid series were selected or invalid input.")
        except Exception as e:
            logging.error(f"An error occurred during series selection: {str(e)}")
            return None, None

def process_box_selection(box_df, folder_df, working_directory, collection_name, call_number):
    box_list = sorted(box_df['BOX'].tolist(), key=custom_sort_key)
    
    # Display options in columns
    num_columns = (len(box_list) + 29) // 30  # Calculate the number of columns needed
    for i in range(0, len(box_list), 30):
        column_options = box_list[i:i+30]
        print("\n".join("{:2d}. {}".format(idx+1, option) for idx, option in enumerate(column_options, start=i)))
        if num_columns > 1:
            print("\t\t", end="")  # Add tab space between columns

    while True:
        try:
            user_input_for_boxes = input("\n\nType in individual numbers, range, or a combination (e.g., '1', '2-3', '4, 5-6') or 'q' to quit: \n\n")
            logging.info(f"user selects Box(es): {user_input_for_boxes}")
            if user_input_for_boxes.lower() == 'q':
                print("\nExiting box selection...\n")
                return None, None

            selected_boxes = parse_user_input(user_input_for_boxes, box_list)

            if selected_boxes is not None:
                selected_boxes_sorted = sorted(selected_boxes)
                print(f"\nYou selected: {', '.join(selected_boxes_sorted)}")

                # Filter folder_df and box_df based on extracted box values
                filtered_folder_df_by_box = filter_df_by_box_values(folder_df, selected_boxes, add_prefix=True)
                filtered_box_df_by_box = filter_df_by_box_values(box_df, selected_boxes, add_prefix=False)

                # Save to Excel and return file paths
                filtered_folder_df_by_box_path = os.path.join(working_directory, f"{collection_name}_{call_number}_folders_by_box_specified.xlsx")
                filtered_box_df_by_box_path = os.path.join(working_directory, f"{collection_name}_{call_number}_boxes_by_box_specified.xlsx")
                filtered_folder_df_by_box.to_excel(filtered_folder_df_by_box_path, index=False)
                filtered_box_df_by_box.to_excel(filtered_box_df_by_box_path, index=False)

                return filtered_folder_df_by_box_path, filtered_box_df_by_box_path
            else:
                print("No valid boxes were selected or invalid input.")
        except Exception as e:
            logging.error(f"An error occurred during box selection: {str(e)}")
            return None, None