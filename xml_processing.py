# xml_processing.py

"""
Module for processing XML files.

This module contains functions for identifying EAD files, preprocessing and sanitizing XML,
and extracting relevant metadata from the parsed XML files. It provides functionality for
handling character encoding issues, parsing XML, and extracting collection-level information.
"""

import glob
import os
from lxml import etree as ET
import logging

from utils import move_recent_ead_files
from data_processing import is_terminal_node  
from user_interaction import user_select_collection

def is_ead_file(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
        first_lines = ''.join(file.readline() for _ in range(10))
        return '<ead>' in first_lines or '<ead ' in first_lines  # Simple check for EAD root element

def preprocess_ead_file(file_path):
    sanitized_file_path = file_path.replace(".xml", "_sanitized.xml")
    if not try_parse(file_path):
        print(f"\nSanitizing EAD file due to character encoding issues: {file_path}\n")
        sanitize_xml(file_path, sanitized_file_path)
        if try_parse(sanitized_file_path):
            return sanitized_file_path
        else:
            print(f"Failed to parse EAD file even after sanitizing: {file_path}")
            return None
    else:
        return file_path
    
def try_parse(input_file):
    """Attempt to parse EAD(.xml) and return boolean result."""
    try:
        ET.parse(input_file)
        return True
    except ET.XMLSyntaxError:
        return False
    
def sanitize_xml(input_file_path, output_file_path):
    """Sanitize EAD by replacing characters that are not allowed in .xml"""

    def is_valid_xml_char(ch):
        codepoint = ord(ch)
        return (
            codepoint == 0x9 or 
            codepoint == 0xA or 
            codepoint == 0xD or 
            (0x20 <= codepoint <= 0xD7FF) or 
            (0xE000 <= codepoint <= 0xFFFD) or 
            (0x10000 <= codepoint <= 0x10FFFF)
        )

    replaced_data = {}

    with open(input_file_path, 'r', encoding='utf-8') as infile, open(output_file_path, 'w', encoding='utf-8') as outfile:
        for line_num, line in enumerate(infile, start=1):  # enumerate will provide a counter starting from 1
            replaced_chars = []
            sanitized_line_list = []
            
            for ch in line:
                if is_valid_xml_char(ch):
                    sanitized_line_list.append(ch)
                else:
                    sanitized_line_list.append('?')
                    replaced_chars.append(ch)
            
            # If we found invalid chars in this line, store them in the dictionary
            if replaced_chars:
                replaced_data[line_num] = replaced_chars
                print(f"Found invalid characters on line {line_num}: {' '.join(replaced_chars)}")

            outfile.write(''.join(sanitized_line_list))

    if not replaced_data:
        print("No invalid characters found!")
    else:
        total_lines = len(replaced_data)
        total_chars = sum(len(chars) for chars in replaced_data.values())
        print(f"Done checking! Found and replaced {total_chars} invalid characters on {total_lines} lines.")
        
    return replaced_data

def process_ead_files(working_directory, namespaces):
    try:
        move_recent_ead_files(working_directory)
        
        # Fetch all XML files in the working directory
        xml_files = glob.glob(os.path.join(working_directory, '*.xml'))
        logging.info(f"Total XML files found in working directory: {len(xml_files)}")

        # Sort files by modification time, with most recent first
        xml_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Filter for EAD files
        ead_files = [file for file in xml_files if is_ead_file(file)]
        logging.info(f"Total EAD files after filtering: {len(ead_files)}")

        if len(ead_files) == 0:
            logging.info("No EAD files found after filtering.")
            print("No EAD files found in the directory.\n")
            print("Please make sure to bring over the EAD finding aid file into this directory.\n")
            print("Until then...thank you, and goodbye!\n")
            return None

        collections = []
        
        # this part extracts the generic data from EAD that'll go on label/printed to console
        for file_path in ead_files:
            processed_file = preprocess_ead_file(file_path)
            if processed_file is not None:
                try:
                    tree = ET.parse(processed_file)
                    root = tree.getroot()
                    
                    repository_element = root.find('./ns:archdesc/ns:did/ns:repository/ns:corpname', namespaces=namespaces)
                    collection_name_element = root.find('./ns:archdesc/ns:did/ns:unittitle', namespaces=namespaces)
                    call_num_element = root.find('./ns:archdesc/ns:did/ns:unitid', namespaces=namespaces)
                    finding_aid_author_element = root.find('./ns:eadheader/ns:filesdesc/ns:titlestmt/ns:author', namespaces=namespaces)
                    
                    repository_name = repository_element.text if repository_element is not None else "Unknown Repository"
                    collection_name = collection_name_element.text if collection_name_element is not None else "Unknown Collection"
                    call_number = call_num_element.text if call_num_element is not None else "Unknown Call Number"
                    finding_aid_author = finding_aid_author_element.text if finding_aid_author_element is not None else "by Unknown Author"

                    collections.append({"path": processed_file, "name": collection_name, "number": call_number, "repository": repository_name, "author": finding_aid_author})

                except Exception as e:
                    logging.error(f"Error processing file {file_path}: {str(e)}")
                    print(f"Encountered an error with file {file_path}, but continuing with processing.\n")

        if len(collections) == 1:
            return collections[0]

        elif len(collections) > 1:
            return user_select_collection(collections)

        else:
            print("No suitable EAD files found for processing.\n")
            return None

    except Exception as e:
        logging.error(f"Error in process_ead_files: {str(e)}")
        return None
    try:
        move_recent_ead_files(working_directory)
        
        # Fetch all XML files in the working directory
        xml_files = glob.glob(os.path.join(working_directory, '*.xml'))
        logging.info(f"Total XML files found in working directory: {len(xml_files)}")

        # Sort files by modification time, with most recent first
        xml_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Filter for EAD files
        ead_files = [file for file in xml_files if is_ead_file(file)]
        logging.info(f"Total EAD files after filtering: {len(ead_files)}")

    except Exception as e:
        logging.error(f"Error in process_ead_files: {str(e)}")
    
    if len(ead_files) == 0:
        logging.info("No EAD files found after filtering.")
        print("No EAD files found in the directory.\n")
        print("Please make sure to bring over the EAD finding aid file into this directory.\n")
        print("Until then...thank you, and goodbye!\n")
        return None

    collections = []
    
    # this part extracts the generic data from EAD that'll go on label/printed to console
    for file_path in ead_files:
        processed_file = preprocess_ead_file(file_path)
        if processed_file is not None:
            try:
                tree = ET.parse(processed_file)
                root = tree.getroot()
                
                repository_element = root.find('./ns:archdesc/ns:did/ns:repository/ns:corpname', namespaces=namespaces)
                collection_name_element = root.find('./ns:archdesc/ns:did/ns:unittitle', namespaces=namespaces)
                call_num_element = root.find('./ns:archdesc/ns:did/ns:unitid', namespaces=namespaces)
                finding_aid_author_element = root.find('./ns:eadheader/ns:filesdesc/ns:titlestmt/ns:author', namespaces=namespaces)
                
                repository_name = repository_element.text if repository_element is not None else "Unknown Repository"
                collection_name = collection_name_element.text if collection_name_element is not None else "Unknown Collection"
                call_number = call_num_element.text if call_num_element is not None else "Unknown Call Number"
                finding_aid_author = finding_aid_author_element.text if finding_aid_author_element is not None else "by Unknown Author"

                collections.append({"path": processed_file, "name": collection_name, "number": call_number, "repository": repository_name, "author": finding_aid_author})

            except Exception as e:
                logging.error(f"Error processing file {file_path}: {str(e)}")
                print(f"Encountered an error with file {file_path}, but continuing with processing.\n")

    if len(collections) == 1:
        return collections[0]

    elif len(collections) > 1:
        return user_select_collection(collections)

    else:
        print("No suitable EAD files found for processing.\n")
        return None