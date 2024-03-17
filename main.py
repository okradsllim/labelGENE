"""
This is the main driver script for the label generation application, labelGENE. 
It integrates the functionalities of other modules to process Extensible Markup Language (XML) Encoded Archival Description (EAD) files, interact with the user for input, and generate labels for archival boxes and folders. 
It sets up the working directory, initializes data processing, and manages the workflow from XML processing to mail merging.
"""

import logging
import os
import sys
import pandas as pd
import re
from lxml import etree as ET
import win32com.client

from xml_processing import process_ead_files, is_terminal_node
from user_interaction import user_select_collection
from data_processing import process_series_selection, process_box_selection, has_explicit_folder_numbering, has_implicit_folder_numbering
from mail_merge import label_selection_menu
from data_extraction import extract_ancestor_data
from utils import box_sort_order, prepend_or_fill


# Constants
NAMESPACES = {'ns': 'urn:isbn:1-931666-22-9'}

# Logging configuration
logging.basicConfig(filename='program_log.txt', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')


def main():
    print("\nHello! Thanks for testing this program: enhancements will be coming soon, so stay tuned!")

    # Set the correct working directory based on the execution context
    working_directory = get_working_directory()

    collection_info = process_ead_files(working_directory, NAMESPACES)

    if collection_info is not None:
        # Initialize folder and box dataframes with preset headers
        folder_df = initialize_folder_dataframe()
        box_df = initialize_box_dataframe()

        # Extract general relevant data
        collection_name, call_number, repository_name, finding_aid_author = extract_collection_info(collection_info)

        # Process and populate dataframes
        folder_df, box_df = process_collection(collection_info, collection_name, call_number, repository_name, folder_df, box_df, NAMESPACES)

        # Prompt user for folder numbering preference
        folder_numbering_preference, folders_already_numbered = prompt_folder_numbering_preference(folder_df)

        # Finalize dataframes based on folder numbering preference
        folder_df, box_df = finalize_dataframes(folder_df, box_df, collection_name, call_number, repository_name, folder_numbering_preference, folders_already_numbered, NAMESPACES)

        # Generate Excel files for mail merge
        excel_file_for_folders, excel_file_for_boxes = generate_excel_files(folder_df, box_df, collection_name, call_number, working_directory)

        # Prompt user for label selection
        process_label_selection(excel_file_for_folders, excel_file_for_boxes, working_directory, folder_numbering_preference, folders_already_numbered, collection_name)

        logging.info('Program finished.')

        # Warning to user before they print flagged label
        check_flagged_labels(folder_df, box_df)

        input(f"\nPress any key and 'Enter' to exit...")
        
    else:
        
        check_flagged_labels()


def get_working_directory():
    # Determine the correct working directory based on the execution context
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle (as an executable)
        return os.path.dirname(sys.executable)
    else:
        # Running as a normal Python script
        return os.path.dirname(os.path.abspath(__file__))

def initialize_folder_dataframe():
    # Initialize folder dataframe with preset headers
    return pd.DataFrame(columns=['COLLECTION', 'CALL_NO.', 'BOX', 'FOLDER', 'CONTAINER_TYPE',
                                 'C01_ANCESTOR', 'C02_ANCESTOR', 'C03_ANCESTOR', 'C04_ANCESTOR', 'C05_ANCESTOR', 
                                 'FOLDER TITLE', 'FOLDER DATES']).astype('object')

def initialize_box_dataframe():
    # Initialize box dataframe with preset headers
    return pd.DataFrame(columns=['REPOSITORY', 'COLLECTION', 'CALL_NO.', 'BOX', 'FOLDER_COUNT', 'FIRST_FOLDER', 'LAST_FOLDER', 'CONTAINER_TYPE', 
                                 'FIRST_C01_SERIES', 'SECOND_C01_SERIES', 'THIRD_C01_SERIES', 'FOURTH_C01_SERIES', 'FIFTH_C01_SERIES']).astype('object')

def extract_collection_info(collection_info):
    # Extract general relevant data from the collection info
    tree = ET.parse(collection_info["path"])
    root = tree.getroot()
    collection_name = collection_info["name"]
    call_number = collection_info["number"]
    repository_name = collection_info["repository"]
    finding_aid_author = collection_info["author"]
    return collection_name, call_number, repository_name, finding_aid_author

def process_collection(collection_info, collection_name, call_number, repository_name, folder_df, box_df, namespaces):
    # Extract relevant data from the collection info
    tree = ET.parse(collection_info["path"])
    root = tree.getroot()
    dsc_element = root.find('.//ns:dsc', namespaces=namespaces)

    print(f"\nProcessing {collection_name} : {call_number}")

    has_explicit_folder_numbering_count = 0
    has_implicit_folder_numbering_count = 0

    if dsc_element is not None:
        all_c_elements = [elem for elem in dsc_element.iterdescendants() if re.match(r'c\d{0,2}$|^c$', ET.QName(elem.tag).localname)]

        for elem in all_c_elements:
            try:
                if is_terminal_node(elem):
                    did_element = elem.find('.//ns:did', namespaces=namespaces)
                    ancestor_data = extract_ancestor_data(did_element, namespaces)

                    if did_element is not None:
                        containers = [elem for elem in did_element.iterchildren() if ET.QName(elem.tag).localname == 'container']
                        container_count = len(containers)

                        has_folder = any(elem.attrib.get('type', '').lower() == 'folder' for elem in containers)
                        has_box = any(elem.attrib.get('type', '').lower() == 'box' for elem in containers)

                        if container_count >= 2 and has_folder:
                            has_explicit_folder_numbering(did_element, containers, ancestor_data, folder_df, namespaces)
                            has_explicit_folder_numbering_count += 1
                        elif container_count == 1 and has_box:
                            has_implicit_folder_numbering(did_element, ancestor_data, folder_df, namespaces)
                            has_implicit_folder_numbering_count += 1
                        else:
                            has_implicit_folder_numbering(did_element, ancestor_data, folder_df, namespaces)
                            has_implicit_folder_numbering_count += 1

            except Exception as e:
                title_text = ""
                date_text = ""
                try:
                    did_element = elem.find('.//ns:did', namespaces=namespaces)
                    if did_element is not None:
                        unittitle_elem = did_element.find('ns:unittitle', namespaces=namespaces)
                        if unittitle_elem is not None:
                            title_text = " ".join(unittitle_elem.itertext()).strip()
                        unitdate_elem = did_element.find('ns:unitdate', namespaces=namespaces)
                        if unitdate_elem is not None:
                            date_text = unitdate_elem.text or ""

                    if not title_text:
                        title_text = "Title unavailable"
                    if not date_text:
                        date_text = "Date unavailable"

                except Exception:
                    title_text = "Error extracting title"
                    date_text = "Error extracting date"

                print(f"Ran into a hiccup with a component titled '{title_text}' from {date_text}: {str(e)}\nI'll keep working though")

    if has_implicit_folder_numbering_count > has_explicit_folder_numbering_count:
        print(f"\nOh boy! The folders have not been numbered; maybe I can help ;)\n")

    return folder_df, box_df

def prompt_folder_numbering_preference(folder_df):
    folder_numbering_preference = None
    folders_already_numbered = folder_df['FOLDER'].notna().sum() > folder_df['FOLDER'].isna().sum()
    if not folders_already_numbered:
        while True:
            folder_numbering_preference = input("If you want the folders numbered, choose numbering preference or press '3' to exit... \n"
                                                "\n1. Continuous (box labels show FIRST - LAST folder number range) "
                                                "\n2. Non-Continuous (box labels show total FOLDER COUNT per box) "
                                                "\n3. Exit program\n\n")
            if folder_numbering_preference in ["1", "2"]:
                break
            elif folder_numbering_preference == "3":
                print("\nExiting program...Thanks, and have a great day!")
                sys.exit()
            else:
                print("Invalid input. Please enter '1', '2', or '3'.")

    return folder_numbering_preference, folders_already_numbered

def finalize_dataframes(folder_df, box_df, collection_name, call_number, repository_name, folder_numbering_preference, folders_already_numbered, namespaces):
    # Finalizing base folder_df
    folder_df['sort_order'] = folder_df['BOX'].apply(box_sort_order)
    folder_df['Folder_temp'] = folder_df['FOLDER'].apply(lambda x: int(re.search(r'\d+', x).group()) if x and re.search(r'\d+', x) else 0)
    folder_df.sort_values(by=['sort_order', 'Folder_temp'], inplace=True)

    # Finalizing base box_df based on folder numbering/user preferences
    if folders_already_numbered or folder_numbering_preference == "1":
        folder_df['BOX'] = [prepend_or_fill('BOX', val, idx) for idx, val in enumerate(folder_df['BOX'])]
        folder_df['FOLDER'] = [prepend_or_fill('FOLDER', val, idx) for idx, val in enumerate(folder_df['FOLDER'])]
        logging.info(f"Preparing dataFrame for {collection_name} boxes")

        # Finalizing continuous box_df
        c01_series_columns = ['FIRST_C01_SERIES', 'SECOND_C01_SERIES', 'THIRD_C01_SERIES', 'FOURTH_C01_SERIES', 'FIFTH_C01_SERIES']

        unique_boxes = folder_df['BOX'].unique()
        for box in unique_boxes:
            box_rows = folder_df[folder_df['BOX'] == box]

            folder_per_box_count = box_rows['FOLDER'].count()
            folder_string = "folder" if folder_per_box_count == 1 else "folders"
            folder_count = f"{folder_per_box_count} {folder_string}"

            first_folder = min([int(re.search(r'(\d+)', folder).group(1)) for folder in box_rows['FOLDER']])
            last_folder = max([int(re.search(r'(\d+)', folder).group(1)) for folder in box_rows['FOLDER']])           
            if first_folder == last_folder:
                last_folder = None

            box_df.loc[len(box_df), ['BOX', 'FIRST_FOLDER', 'LAST_FOLDER', 'FOLDER_COUNT']] = [box, first_folder, last_folder, folder_count]

            unique_ancestors = box_rows['C01_ANCESTOR'].unique()
            for i, ancestor in enumerate(unique_ancestors):
                if i >= len(c01_series_columns):
                    break
                box_df.at[len(box_df)-1, c01_series_columns[i]] = ancestor

            unique_container_types = box_rows['CONTAINER_TYPE'].unique()
            for container_type in unique_container_types:
                box_df.at[len(box_df)-1, 'CONTAINER_TYPE'] = container_type

            box_df.at[len(box_df)-1, 'REPOSITORY'] = repository_name
            box_df.at[len(box_df)-1, 'COLLECTION'] = collection_name
            box_df.at[len(box_df)-1, 'CALL_NO.'] = call_number

    if folder_numbering_preference == "2":
        folder_df['BOX'] = "Box " + folder_df['BOX'].astype(str)

        empty_folder_indices = folder_df[folder_df['FOLDER'].isna()].index
        if not empty_folder_indices.empty:  
            folder_counter = 1
            current_box = None
            for index in empty_folder_indices:
                row = folder_df.loc[index]
                if current_box != row['BOX']:
                    current_box = row['BOX']
                    folder_counter = 1
                folder_df.at[index, 'FOLDER'] = f"{folder_counter}"
                folder_counter += 1

        folder_df['FOLDER'] = "Folder " + folder_df['FOLDER'].astype(str)

        # Finalizing non-continuous box_df
        c01_series_columns = ['FIRST_C01_SERIES', 'SECOND_C01_SERIES', 'THIRD_C01_SERIES', 'FOURTH_C01_SERIES', 'FIFTH_C01_SERIES']

        unique_boxes = folder_df['BOX'].unique()
        for box in unique_boxes:
            box_rows = folder_df[folder_df['BOX'] == box]

            folder_per_box_count = box_rows['FOLDER'].count()
            folder_string = "folder" if folder_per_box_count == 1 else "folders"
            folder_count = f"{folder_per_box_count} {folder_string}"

            box_df.loc[len(box_df), ['BOX', 'FOLDER_COUNT']] = [box, folder_count]

            unique_ancestors = box_rows['C01_ANCESTOR'].unique()
            for i, ancestor in enumerate(unique_ancestors):
                if i >= len(c01_series_columns):
                    break
                box_df.at[len(box_df)-1, c01_series_columns[i]] = ancestor

            unique_container_types = box_rows['CONTAINER_TYPE'].unique()
            for container_type in unique_container_types:
                box_df.at[len(box_df)-1, 'CONTAINER_TYPE'] = container_type

            box_df.at[len(box_df)-1, 'REPOSITORY'] = repository_name
            box_df.at[len(box_df)-1, 'COLLECTION'] = collection_name
            box_df.at[len(box_df)-1, 'CALL_NO.'] = call_number

    # Drop temporary columns before finalizing and strip col 'BOX' of "Box"
    folder_df.drop(columns=['sort_order', 'Folder_temp'], inplace=True) 
    box_df['BOX'] = box_df['BOX'].apply(lambda x: x.replace('Box', '').strip())
    print(f"\nCounted a total of {len(folder_df)} folder{'s' if len(folder_df) != 1 else ''} in {len(box_df)} box{'es' if len(box_df) != 1 else ''}")

    return folder_df, box_df

def generate_excel_files(folder_df, box_df, collection_name, call_number, working_directory):
    # Generate Excel files for mail merge
    logging.info(f"Prepping Excel files for mail merge operation")
    folder_dataFrame_path = os.path.join(working_directory, f"{collection_name}_{call_number}_folder.xlsx")
    box_dataFrame_path = os.path.join(working_directory, f"{collection_name}_{call_number}_box.xlsx")

    folder_df.to_excel(folder_dataFrame_path, index=False)
    box_df.to_excel(box_dataFrame_path, index=False)

    return folder_dataFrame_path, box_dataFrame_path

def process_label_selection(excel_file_for_folders, excel_file_for_boxes, working_directory, folder_numbering_preference, folders_already_numbered, collection_name):
    wordApp = win32com.client.Dispatch('Word.Application')
    label_selection_menu(wordApp, excel_file_for_folders, excel_file_for_boxes, working_directory, folder_numbering_preference, folders_already_numbered, collection_name)

def check_flagged_labels(folder_df=None, box_df=None):
    if folder_df is not None and box_df is not None:
        if "10001" in folder_df['BOX'].values or "10001" in box_df['BOX'].values:
            print("\nNote before you leave: '10001' was used as a flag for non-standard box numbering in this collection. \nPlease verify and update box data before printing labels.\n")
            print(f"Goodbye!")

if __name__ == "__main__":
    main()