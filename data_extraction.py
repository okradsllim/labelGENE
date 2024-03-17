# data_extraction.py

"""
Module for extracting data from EAD elements.

This module contains functions for extracting specific data elements from EAD XML, such as box numbers,
folder dates, base folder titles, and ancestor data. These functions are used to retrieve relevant
metadata from the parsed XML files.
"""

import re
from lxml import etree as ET

def extract_box_number(did_element, namespaces):
    """Extracts the box number from the given element."""
    all_box_components = did_element.findall(".//ns:container", namespaces=namespaces)
    for box_component in all_box_components:
        if box_component.attrib.get('type', '').lower() == 'box':
            return box_component.text # Beinecke EADs almost always have terminal c components which if has only one container element, that element is of @type 'box'
    # but what if only one container element but @type 'folder'? 
    # if not @type box get whatever is in text node if not @type folder
    # helps catch container elem with direct text like 123 (Art) (e.g. Henri Chopin ead(encoding problem fixed))
    for box_component in all_box_components:
        if box_component.attrib.get('type', '').lower() != 'folder': 
            return box_component.text
    return "10001"  # arbitrary num string flag for unusual/unavailable box number info/num cos mathematically useful than using "Box unavailable"

def extract_folder_date(did_element, namespaces):
    """Extracts the folder date from the given element."""
    unitdate_element = did_element.find(".//ns:unitdate", namespaces=namespaces)
    return unitdate_element.text if unitdate_element is not None else "Date unavailable" # arbitrary to ensure return type str 

def extract_base_folder_title(did_element, namespaces):
    """Extracts base FOLDER_TITLE from the given element."""
    unittitle_element = did_element.find(".//ns:unittitle", namespaces=namespaces)
    return " ".join(unittitle_element.itertext()) if unittitle_element is not None else "Title unavailable" # arbitrary to ensure return type str 

def extract_ancestor_data(node, namespaces):
    """extracts ancestor data from each terminal <c>/<cxx> node"""
    ancestors_data = []
    ancestor_count = 0 # for limits to column number count for c ancestors

    # Exclude descendant from returned list: ancestors only
    #ancestors = node.iterancestors()
    #ancestors = ancestors[::-1]
    
    ancestors = list(reversed(list(node.iterancestors())))
    ancestors.pop()

    for ancestor in ancestors:
        ancestor_tag = ET.QName(ancestor.tag).localname

        # Match both unnumbered/numbered c tags
        if re.match(r'c\d{0,2}$|^c$', ancestor_tag): 
            is_first_gen_c = ET.QName(ancestor.getparent().tag).localname == 'dsc' # because all 1st gen 'c'/cxx' are direct children of <dsc>
            #is_series = ancestor.attrib.get('level') == 'series' # because not all first_gen_c's are an archival 'series'

            did_element = ancestor.find("./ns:did", namespaces=namespaces)
            if did_element is not None:
                unittitle_element = did_element.find("./ns:unittitle", namespaces=namespaces)
                unittitle = " ".join(unittitle_element.itertext()).strip() if unittitle_element is not None else "X" # arbitrary to ensure return type str 

                unitid_element = did_element.find("./ns:unitid", namespaces=namespaces)
                if is_first_gen_c and unitid_element is not None:  # if unitid in first_gen_c, it MUST be @level="series"
                    unitid_text = unitid_element.text
                    try:
                        unitid = int(unitid_text)
                        if unitid <= 40:  # Convert to Roman only if unitid is an integer and the integer is up to 40
                            roman_numeral = convert_to_roman(unitid)
                            ancestors_data.append(f"Series {roman_numeral}. {unittitle}") 
                        else:
                            ancestors_data.append(unittitle) 
                    except ValueError:  # For non-integer unitids
                        ancestors_data.append(f"Series {unitid_text}. {unittitle}")  
                else:
                        ancestors_data.append(unittitle)
                    
                ancestor_count += 1

                if ancestor_count >= 5: # '5' for the 5 <cxx> ancestor columns in folder_df
                    break
                
    return ancestors_data

def convert_to_roman(num):
    # Purpose: Converts integer series numbers to Roman numerals in keeping with the traditional presentation of archival series information.
    # I can't seem to get Roman module to work for me. This will do for now because I haven't yet worked with a collection with more than 12 series
    # Return Type: str - The corresponding Roman numeral as a string, or the number itself if not mapped.
    roman_dict = { 1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI', 7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X', 11: 'XI', 12: 'XII', 13: 'XIII', 14: 'XIV', 15: 'XV',  16:'XVI', 17: 'XVII', 18: 'XVIII', 19: 'XIX', 20: 'XX',
                  21: 'XXI', 22: 'XXII', 23: 'XXIII', 24: 'XXIV', 25: 'XXV', 26: 'XXVI', 27: 'XXVII', 28: 'XXVIII', 29: 'XXIX', 30: 'XXX', 31: 'XXXI', 32: 'XXXII', 33: 'XXXIII', 34: 'XXXIV', 35: 'XXXV',  36:'XXXVI', 37: 'XXXVII', 38: 'XXXVIII', 39: 'XXXIX', 40: 'XL'
    }
    return roman_dict.get(num, str(num))