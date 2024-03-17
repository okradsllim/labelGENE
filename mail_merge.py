# mail_merge.py

"""
Module for generating output files using mail merge.

This module provides functions for performing mail merge operations to generate box and folder label files.
It interacts with Microsoft Word using the win32com library to populate predefined templates with the
extracted and processed data, creating the final output files.
"""

# mail_merge.py

import os
import sys
import time
import logging
import pandas as pd
import re

def perform_mail_merge(wordApp, excel_files, template_name, working_directory):
    time.sleep(1)
    # Determine if running as a script or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
        
    for excel_file in excel_files:
        try:
            template_path = os.path.join(application_path, template_name)
            if not os.path.exists(template_path):
                logging.error(f"Template file not found: {template_path}")
                continue

            logging.info(f"Opening template: {template_path}")
            doc = wordApp.Documents.Open(template_path)

            if "folder" in template_name:
                wordApp.Run("MergeForFolders", excel_file)
            else:
                wordApp.Run("MergeForBoxes", excel_file)

            newDoc = wordApp.ActiveDocument
            
            # Check if 'left' is in the template name and adjust the resulting doc's filename
            if "left" in template_name:
                label_part = '_left_labels'
            else:
                label_part = '_labels'
            
            resulting_doc = os.path.join(working_directory, f"{os.path.basename(excel_file).replace('.xlsx', label_part + '.docx')}")
            logging.info(f"Saving merged document: {resulting_doc}")

            newDoc.SaveAs2(FileName=resulting_doc, FileFormat=16)
            newDoc.Close(SaveChanges=0)

            doc.Saved = True
            doc.Close()
        except Exception as e:
            logging.error(f"An error occurred during mail merge: {str(e)}")
            # Close the current document if it's open
            if 'newDoc' in locals() and newDoc is not None:
                newDoc.Close(SaveChanges=0)
            if 'doc' in locals() and doc is not None:
                doc.Saved = True
                doc.Close()
            # wordApp.Quit()

    logging.info("Mail merge process completed.")

def label_selection_menu(wordApp, folder_excel_path, box_excel_path, working_directory, folder_numbering_preference, folders_already_numbered, collection_name):
    while True:
        try:
            select_label_type = input("\nPlease choose a number for the type of labels you want, or quit program...\n"
                                            "\n1. DEFAULT folder/box "
                                            "\n2. LEFT label (folder) and DEFAULT box "
                                            "\n3. LEFT label (folder) and CUSTOM box"
                                            "\n4. DEFAULT folder and CUSTOM box "
                                            "\n5. DEFAULT folder only "
                                            "\n6. LEFT label (folder) only "
                                            "\n7. DEFAULT box only "
                                            "\n8. CUSTOM box only "
                                            "\n9. Exit program...\n\n")
            
            logging.info(f"User enters: {select_label_type}")
                            
            if select_label_type == '1': # DEFAULT folder and box labels
                logging.info(f"Option 1 selected: # DEFAULT folder and box labels")
                try:
                    perform_mail_merge(wordApp, [folder_excel_path], "default_folder_template.docm", working_directory)
                    logging.info(f"Mail merge for default folder labels completed.")
                    box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                    perform_mail_merge(wordApp, [box_excel_path], box_template, working_directory)
                    logging.info(f"Mail merge for default box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break
                except Exception as e:
                    logging.error(f"An error occurred in option 1: # DEFAULT folder and box labels {str(e)}")

            elif select_label_type == '2': # LEFT labels (FOLDER) and DEFAULT box labels
                logging.info("Option 2 selected: Left labels for folders and default box labels.")
                try:
                    perform_mail_merge(wordApp, [folder_excel_path], "left_labels_folder_template.docm", working_directory)
                    logging.info("Mail merge for left labels (folder) completed.")
                    box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                    perform_mail_merge(wordApp, [box_excel_path], box_template, working_directory)
                    logging.info("Mail merge for default box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break
                except Exception as e:
                    logging.error(f"An error occurred in option 2: Left labels for folders and default box labels. {str(e)}")        

            elif select_label_type == '3': # LEFT labels (FOLDER) and CUSTOM box labels
                logging.info("Option 3 selected: Left labels for folders and CUSTOM box labels.")
                try:
                    perform_mail_merge(wordApp, [folder_excel_path], "left_labels_folder_template.docm", working_directory)
                    logging.info("Mail merge for left labels (folder) completed.")
                    
                    # Read the Excel file into a DataFrame for processing
                    custom_df_box = pd.read_excel(box_excel_path)
                    
                    # Create a copy of the custom_df_box for the default_box_df
                    default_box_df = custom_df_box.copy()

                    def check_flat_box_condition(row):
                        # Convert row to string to ensure .split() can be called
                        row = str(row)
                        # Check if 'flat box' is in the string and proceed with extraction
                        if row.startswith('flat box'):
                            # Find all parts that contain 'h' which indicates height measurement
                            height_parts = [part.replace('h', '') for part in row.split() if 'h' in part]
                            for part in height_parts:
                                try:
                                    # Check if any part that contains 'h' has a number greater than 2
                                    if float(part) > 2:
                                        return True
                                except ValueError as e:
                                    # Log the error and ignore this part if it's not a valid number
                                    logging.error(f"Error converting part to float: {part}, Error: {e}")
                        return False

                    # Group 1: Archive Half Legal and Archive Half Letter Boxes
                    archive_half_df = custom_df_box[custom_df_box['CONTAINER_TYPE'].isin(['archive half legal', 'archive half letter'])]
                    logging.info(f"Number of 'archive half legal' and 'archive half letter' containers: {len(archive_half_df)}")


                    if not archive_half_df.empty:
                        # Remove rows from default_box_df corresponding to archive_half_df
                        default_box_df = default_box_df.drop(archive_half_df.index)
                        # Perform mail merge for archive_half_df
                        archive_half_legal_path = os.path.join(working_directory, f"{collection_name}_half_hollinger.xlsx")
                        archive_half_df.to_excel(archive_half_legal_path, index=False)
                        box_template = "vertical_half_holl_continuous_numbering.docm" if folders_already_numbered else ("vertical_half_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "vertical_half_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [archive_half_legal_path], box_template, working_directory)
                        logging.info("Mail merge for half Hollinger custom box labels completed.")

                    # Group 2: Special Flat Boxes
                    custom_df_box['CONTAINER_TYPE'] = custom_df_box['CONTAINER_TYPE'].astype(str) # so that "NaN"s don't throw off df manipulations with .str
                    flat_box_df = custom_df_box[
                        custom_df_box['CONTAINER_TYPE'].str.startswith('flat box') & 
                        custom_df_box['CONTAINER_TYPE'].apply(check_flat_box_condition)
                    ]
                    logging.info(f"Number of 'flat box' containers with height > 2: {len(flat_box_df)}")

                    if not flat_box_df.empty:
                        # Remove rows from default_box_df corresponding to flat_box_df
                        default_box_df = default_box_df.drop(flat_box_df.index)
                        # Perform mail merge for flat_box_df
                        flat_box_path = os.path.join(working_directory, f"{collection_name}_flat_box_tall.xlsx")
                        flat_box_df.to_excel(flat_box_path, index=False)
                        box_template = "half_horizontal_holl_continuous_numbering.docm" if folders_already_numbered else ("half_horizontal_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "half_horizontal_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [flat_box_path], box_template, working_directory)
                        logging.info("Mail merge for flat box where height is more than '2' custom box labels completed.")

                    # Group 3: Default Boxes
                    if not default_box_df.empty:
                        # Perform mail merge for default_box_df
                        default_box_path = os.path.join(working_directory, f"{collection_name}_default_hollinger.xlsx")
                        default_box_df.to_excel(default_box_path, index=False)
                        box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [default_box_path], box_template, working_directory)
                        logging.info("Mail merge for non-half Hollinger and/or flat box less than 2 in height custom box labels completed.")

                    logging.info("Mail merge for all custom box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break

                except Exception as e:
                    logging.error(f"An error occurred in option 3: # CUSTOM box labels. {str(e)}")    
            
            elif select_label_type == '4': # DEFAULT folder and CUSTOM box labels
                logging.info("Option 4 selected: # DEFAULT folder and CUSTOM box labels.")
                try:
                    # Default folder mail merge
                    perform_mail_merge(wordApp, [folder_excel_path], "default_folder_template.docm", working_directory)
                    logging.info("Mail merge for default folder labels completed.")
                    
                    # Read the Excel file into a DataFrame for processing
                    custom_df_box = pd.read_excel(box_excel_path)
                    # Create a copy of the custom_df_box for the default_box_df
                    default_box_df = custom_df_box.copy()

                    def check_flat_box_condition(row):
                        # Convert row to string to ensure .split() can be called
                        row = str(row)
                        # Check if 'flat box' is in the string and proceed with extraction
                        if row.startswith('flat box'):
                            # Find all parts that contain 'h' which indicates height measurement
                            height_parts = [part.replace('h', '') for part in row.split() if 'h' in part]
                            for part in height_parts:
                                try:
                                    # Check if any part that contains 'h' has a number greater than 2
                                    if float(part) > 2:
                                        return True
                                except ValueError as e:
                                    # Log the error and ignore this part if it's not a valid number
                                    logging.error(f"Error converting part to float: {part}, Error: {e}")
                        return False

                    # Group 1: Archive Half Legal and Archive Half Letter Boxes
                    archive_half_df = custom_df_box[custom_df_box['CONTAINER_TYPE'].isin(['archive half legal', 'archive half letter'])]
                    logging.info(f"Number of 'archive half legal' and 'archive half letter' containers: {len(archive_half_df)}")


                    if not archive_half_df.empty:
                        # Remove rows from default_box_df corresponding to archive_half_df
                        default_box_df = default_box_df.drop(archive_half_df.index)
                        # Perform mail merge for archive_half_df
                        archive_half_legal_path = os.path.join(working_directory, f"{collection_name}_half_hollinger.xlsx")
                        archive_half_df.to_excel(archive_half_legal_path, index=False)
                        box_template = "vertical_half_holl_continuous_numbering.docm" if folders_already_numbered else ("vertical_half_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "vertical_half_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [archive_half_legal_path], box_template, working_directory)
                        logging.info("Mail merge for half Hollinger custom box labels completed.")

                    # Group 2: Special Flat Boxes
                    custom_df_box['CONTAINER_TYPE'] = custom_df_box['CONTAINER_TYPE'].astype(str) # so that "NaN"s don't throw off df manipulations with .str
                    flat_box_df = custom_df_box[
                        custom_df_box['CONTAINER_TYPE'].str.startswith('flat box') & 
                        custom_df_box['CONTAINER_TYPE'].apply(check_flat_box_condition)
                    ]
                    logging.info(f"Number of 'flat box' containers with height > 2: {len(flat_box_df)}")

                    if not flat_box_df.empty:
                        # Remove rows from default_box_df corresponding to flat_box_df
                        default_box_df = default_box_df.drop(flat_box_df.index)
                        # Perform mail merge for flat_box_df
                        flat_box_path = os.path.join(working_directory, f"{collection_name}_flat_box_tall.xlsx")
                        flat_box_df.to_excel(flat_box_path, index=False)
                        box_template = "half_horizontal_holl_continuous_numbering.docm" if folders_already_numbered else ("half_horizontal_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "half_horizontal_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [flat_box_path], box_template, working_directory)
                        logging.info("Mail merge for flat box ~ < 2 in 'h' custom box labels completed.")

                    # Group 3: Default Boxes
                    if not default_box_df.empty:
                        # Perform mail merge for default_box_df
                        default_box_path = os.path.join(working_directory, f"{collection_name}_default_hollinger.xlsx")
                        default_box_df.to_excel(default_box_path, index=False)
                        box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [default_box_path], box_template, working_directory)
                        logging.info("Mail merge for non-half Hollinger and/or flat box ~ < 2 in h custom box labels completed.")

                    logging.info("Mail merge for all custom box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break

                except Exception as e:
                    logging.error(f"An error occurred in option 4: # CUSTOM box labels. {str(e)}")    
                            
            elif select_label_type == '5': # Default folders only
                logging.info("Option 5 selected: # DEFAULT folder labels")
                try:
                    perform_mail_merge(wordApp, [folder_excel_path], "default_folder_template.docm", working_directory)
                    logging.info("Mail merge for default folder labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break
                except Exception as e:
                    logging.error(f"An error occurred in option 5: # DEFAULT folder labels {str(e)}")
            
            elif select_label_type == '6': # LEFT labels (FOLDER) and DEFAULT box labels
                logging.info("Option 6 selected: Left labels for folders.")
                try:
                    perform_mail_merge(wordApp, [folder_excel_path], "left_labels_folder_template.docm", working_directory)
                    logging.info("Mail merge for left labels (folder) completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break
                except Exception as e:
                    logging.error(f"An error occurred in option 6: Left labels for folders. {str(e)}") 
                
            elif select_label_type == '7': # DEFAULT box labels only
                logging.info("Option 7 selected: DEFAULT box labels only.")
                try:
                    box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                    perform_mail_merge(wordApp, [box_excel_path], box_template, working_directory)
                    logging.info("Mail merge for default box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break
                except Exception as e:
                    logging.error(f"An error occurred in option 7: DEFAULT box labels only. {str(e)}")

            elif select_label_type == '8': # CUSTOM box labels
                logging.info("Option 8 selected: # CUSTOM box labels.")
                try:
                    # Read the Excel file into a DataFrame for processing
                    custom_df_box = pd.read_excel(box_excel_path)
                    # Create a copy of the custom_df_box for the default_box_df
                    default_box_df = custom_df_box.copy()

                    def check_flat_box_condition(row):
                        # Convert row to string to ensure .split() can be called
                        row = str(row)
                        # Check if 'flat box' is in the string and proceed with extraction
                        if row.startswith('flat box'):
                            # Find all parts that contain 'h' which indicates height measurement
                            height_parts = [part.replace('h', '') for part in row.split() if 'h' in part]
                            for part in height_parts:
                                try:
                                    # Check if any part that contains 'h' has a number greater than 2
                                    if float(part) > 2:
                                        return True
                                except ValueError as e:
                                    # Log the error and ignore this part if it's not a valid number
                                    logging.error(f"Error converting part to float: {part}, Error: {e}")
                        return False

                    # Group 1: Archive Half Legal and Archive Half Letter Boxes
                    archive_half_df = custom_df_box[custom_df_box['CONTAINER_TYPE'].isin(['archive half legal', 'archive half letter'])]
                    logging.info(f"Number of 'archive half legal' and 'archive half letter' containers: {len(archive_half_df)}")

                    if not archive_half_df.empty:
                        # Remove rows from default_box_df corresponding to archive_half_df
                        default_box_df = default_box_df.drop(archive_half_df.index)
                        # Perform mail merge for archive_half_df
                        archive_half_legal_path = os.path.join(working_directory, f"{collection_name}_half_hollinger.xlsx")
                        archive_half_df.to_excel(archive_half_legal_path, index=False)
                        box_template = "vertical_half_holl_continuous_numbering.docm" if folders_already_numbered else ("vertical_half_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "vertical_half_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [archive_half_legal_path], box_template, working_directory)
                        logging.info("Mail merge for half Hollinger custom box labels completed.")

                    # Group 2: Special Flat Boxes
                    custom_df_box['CONTAINER_TYPE'] = custom_df_box['CONTAINER_TYPE'].astype(str)
                    flat_box_df = custom_df_box[
                        custom_df_box['CONTAINER_TYPE'].str.startswith('flat box') & 
                        custom_df_box['CONTAINER_TYPE'].apply(check_flat_box_condition)
                    ]
                    logging.info(f"Number of 'flat box' containers with height > 2: {len(flat_box_df)}")

                    if not flat_box_df.empty:
                        # Remove rows from default_box_df corresponding to flat_box_df
                        default_box_df = default_box_df.drop(flat_box_df.index)
                        # Perform mail merge for flat_box_df
                        flat_box_path = os.path.join(working_directory, f"{collection_name}_flat_box_tall.xlsx")
                        flat_box_df.to_excel(flat_box_path, index=False)
                        box_template = "half_horizontal_holl_continuous_numbering.docm" if folders_already_numbered else ("half_horizontal_holl_continuous_numbering.docm" if folder_numbering_preference == "1" else "half_horizontal_holl_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [flat_box_path], box_template, working_directory)
                        logging.info("Mail merge for flat box ~ < 2 in 'h' custom box labels completed.")

                    # Group 3: Default Boxes
                    if not default_box_df.empty:
                        # Perform mail merge for default_box_df
                        default_box_path = os.path.join(working_directory, f"{collection_name}_default_hollinger.xlsx")
                        default_box_df.to_excel(default_box_path, index=False)
                        box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
                        perform_mail_merge(wordApp, [default_box_path], box_template, working_directory)
                        logging.info("Mail merge for non-half Hollinger and/or flat box ~ < 2 in h custom box labels completed.")

                    logging.info("Mail merge for all custom box labels completed.")
                    print(f"\nSuccess! Check directory for the output files...")
                    break

                except Exception as e:
                    logging.error(f"An error occurred in option 8: # CUSTOM box labels. {str(e)}")
                
            elif select_label_type == '9': # Exit
                print("\nExiting program...Thanks, and have a great day!")
                sys.exit()
            
            else:
                print(f"\nWrong input: please make a valid selection...")
        except Exception as e:
            logging.error(f"An error occurred during filtering: {str(e)}")