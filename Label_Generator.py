# Last updated: May 29, 2024


# June 7, 2024 Update: Modified for EAD3 support.
# May 29, 2024 Update: Improved commentaries.
# March 14, 2024 Update: Modified box selection display to show all boxes so that Ruth can see all boxes and select the ones she wants to print labels for.

# This Python script is for processing EAD XML files to produce box and folder labels. It focuses on parsing, data extraction, and metadata management. 
# It starts by identifying EAD files in a specified directory, and then sanitizes these XML files, 
# addressing common encoding issues, and parses them using the lxml library for efficient XML processing. 
# Key metadata elements such as repository details, collection names, call numbers, and specific hierarchical data are systematically extracted. 
# It traverses and stores data from every terminal 'c' node (only recursive in spirit lol), effectively capturing item-level descriptions across a variety of well-formed, even non-valid EAD2002 files. 
# It can handle both explicit and implicit folder numbering, accommodating the varied structuring of EADs. 
# Explicit folder numbering is managed by extracting folder numbers directly from the XML, while implicit numbering involves a nuanced approach to deduce folder counts from available metadata.
# It is complemented by a suite of .docm template Word files, each embedded with VBA code tailored to work in tandem with the script's mail_merge operation function. 
# This integration facilitates the generation of folder and box label files, aligning with the user’s specific labeling preferences.

# November 2023 update notes:
# Enhanced terminal 'c' node Handling:
# - Improved logic to process 'c' nodes with multiple <extent> elements in <physdesc>.
# - Ignores <extent> elements starting with '0' to accurately capture relevant folder data.
# Adjusted to handle '0' quantities in <extent>, common in Ackroyd and other non-standard EADs.
# Refined Folder Numbering: Updated 'has_implicit_folder_numbering' function to accurately extract non-zero integers from all <physdesc> subelements.

# December update notes:
# Script description above refined. Will highlight anything missing later


# IMPORTS AND GLOBAL VARIABLES

import re
from lxml import etree as ET
import pandas as pd
import time
import glob
import sys
import os
import win32com.client
import logging
from datetime import datetime, timedelta
import shutil


# XML PARSING AND EAD FILE PROCESSING FUNCTIONS

def set_namespace(root):
    if 'urn:isbn:1-931666-22-9' in root.nsmap.values():
        namespaces['ns'] = 'urn:isbn:1-931666-22-9'
    elif 'http://ead3.archivists.org/schema/' in root.nsmap.values():
        namespaces['ns'] = 'http://ead3.archivists.org/schema/'
    else:
        raise ValueError("Unsupported EAD version.")
    return namespaces['ns']

def is_ead_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            content = file.read(1024)  # Read first 1024 bytes to accommodate headers
            return '<ead' in content
    except Exception as e:
        logging.error(f"Error reading {file_path}: {e}")
        return False

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

def process_ead_files(working_directory):
    try:
        move_recent_ead_files(working_directory)
        
        # Fetch all XML files in the working directory
        xml_files = glob.glob(os.path.join(working_directory, '*.xml'))
        logging.info(f"Total XML files found in working directory: {len(xml_files)}")

        # Sort files by modification time, with most recent first
        xml_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Filter for EAD files using the enhanced is_ead_file
        ead_files = [file for file in xml_files if is_ead_file(file)]
        logging.info(f"Total EAD files after filtering: {len(ead_files)}")

    except Exception as e:
        logging.error(f"Error in process_ead_files: {str(e)}")
        return None

    if not ead_files:
        logging.info("No EAD files found after filtering.")
        print("\nNo EAD files found in the directory. Please ensure you have EAD files in the directory to process.")
        return None

    collections = []
    
    # Process each EAD file to extract necessary data
    for file_path in ead_files:
        processed_file = preprocess_ead_file(file_path)
        if processed_file is not None:
            try:
                tree = ET.parse(processed_file)
                root = tree.getroot()
                # Dynamically set namespace based on the EAD root element
                namespace_uri = set_namespace(root)
                namespaces['ns'] = namespace_uri

                ead_version = 'ead2002' if namespace_uri == 'urn:isbn:1-931666-22-9' else 'ead3'

                if ead_version == 'ead2002':
                    repository_element = root.find('./ns:archdesc/ns:did/ns:repository/ns:corpname', namespaces=namespaces)
                    finding_aid_author_element = root.find('./ns:eadheader/ns:filedesc/ns:titlestmt/ns:author', namespaces=namespaces)
                else:  # ead3
                    repository_element = root.find('./ns:control/ns:filedesc/ns:publicationstmt/ns:publisher', namespaces=namespaces)
                    finding_aid_author_element = root.find('./ns:control/ns:filedesc/ns:titlestmt/ns:author', namespaces=namespaces)

                collection_name_element = root.find('./ns:archdesc/ns:did/ns:unittitle', namespaces=namespaces)
                call_num_element = root.find('./ns:archdesc/ns:did/ns:unitid', namespaces=namespaces)
                
                repository_name = repository_element.text if repository_element is not None else "Unknown Repository"
                collection_name = collection_name_element.text if collection_name_element is not None else "Unknown Collection"
                call_number = call_num_element.text if call_num_element is not None else "Unknown Call Number"
                finding_aid_author = finding_aid_author_element.text if finding_aid_author_element is not None else "Unknown Author"

                collections.append({
                    "path": processed_file, 
                    "name": collection_name,
                    "number": call_number, 
                    "repository": repository_name, 
                    "author": finding_aid_author,
                    "ead_version": ead_version, 
                    "namespaces": namespaces
                })

            except Exception as e:
                logging.error(f"Error processing file {file_path}: {str(e)}")
                print(f"Encountered an error with file {file_path}, but continuing with processing.\n")

    # Decision-making based on the number of collections processed
    if len(collections) == 1: # If only one collection is found, return it directly and proceed with processing that collecition
        return collections[0]
    elif len(collections) > 1:
        return user_select_collection(collections) # If multiple collections are found, prompt the user to select one
    else:
        print("No suitable EAD files found for processing.\n")
        return None

def try_parse(input_file):
    """Attempt to parse EAD(.xml) and return boolean result."""
    try:
        ET.parse(input_file)
        return True
    except ET.XMLSyntaxError:
        return False
    
def sanitize_xml(input_file_path, output_file_path):
    """Sanitize EAD by replacing characters that are not allowed in .xml"""
    
    # https://stackoverflow.com/questions/8733233/filtering-out-certain-bytes-in-python
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

def is_terminal_node(node):
    """Determines if a node is a terminal node by checking its children."""
    # 06-06-2024: Modified regex to back to original to exclude ns:
    for child in node:
        tag = ET.QName(child.tag).localname
        if re.match(r'c\d{0,2}$|^c$', tag): 
            return False
    return True


# DATA EXTRACTION AND MANIPULATION FUNCTIONS

def extract_box_number(did_element, namespaces, ead_version):
    container_xpath = ".//ns:container"
    type_attr = 'type' if ead_version == 'ead2002' else 'localtype'
    all_box_components = did_element.findall(container_xpath, namespaces=namespaces)
    for box_component in all_box_components:
        if box_component.attrib.get(type_attr, '').lower() == 'box':
            return box_component.text
    for box_component in all_box_components:
        if box_component.attrib.get(type_attr, '').lower() != 'folder':
            return box_component.text
    #print("No box number found")
    return "10001"  # arbitrary num string flag for unusual/unavailable box number info/num cos mathematically useful than using "Box unavailable"

def extract_folder_date(did_element, namespaces, ead_version):
    if ead_version == 'ead2002':
        date_xpath = ".//ns:unitdate"
        unitdate_element = did_element.find(date_xpath, namespaces=namespaces)
        if unitdate_element is not None:
            #print(f"Extracted folder date (EAD2002): {unitdate_element.text}")
            return unitdate_element.text
        else:
            #print("No folder date found (EAD2002)")
            return "Date unavailable"
       
    else:
        date_xpath = ".//ns:unitdatestructured"
        unitdatestructured_element = did_element.find(date_xpath, namespaces=namespaces)
        if unitdatestructured_element is not None:
            if 'altrender' in unitdatestructured_element.attrib:
                return unitdatestructured_element.attrib['altrender']
            else:
                return "Date unavailable"
        else: #Adding this because I've seen some EAD3 where either unitdatestructured or unitdate is used but not both (e.g. John Holker) When <unitdate> is used though, it is always expressing "undated" in Holker; I don't know about others.
            date_xpath = ".//ns:unitdate"
            unitdate_element = did_element.find(date_xpath, namespaces=namespaces)
            if unitdate_element is not None:
                return unitdate_element.text
            else:
                return "Date unavailable"
            
        
        '''if date_range_element is not None:
            from_date = date_range_element.find('ns:fromdate', namespaces=namespaces)
            to_date = date_range_element.find('ns:todate', namespaces=namespaces)
            if from_date is not None and to_date is not None:
                #print(f"Extracted folder date (EAD3): {from_date.text}–{to_date.text}")
                return f"{from_date.text}–{to_date.text}"
            else:
                #print("Incomplete date range found (EAD3)")
                return "Date unavailable"
        else:
            single_date_xpath = ".//ns:unitdatestructured/ns:datesingle"
            single_date_element = did_element.find(single_date_xpath, namespaces=namespaces)
            if single_date_element is not None:
                #print(f"Extracted single folder date (EAD3): {single_date_element.text}")
                return single_date_element.text
            else:
                #print("No folder date found (EAD3)")
                return "Date unavailable"'''

def extract_base_folder_title(did_element, namespaces):
    unittitle_xpath = ".//ns:unittitle"
    unittitle_element = did_element.find(unittitle_xpath, namespaces=namespaces)
    if unittitle_element is not None:
        base_title = " ".join(unittitle_element.itertext())
        #print(f"Extracted base folder title: {base_title}")
        return base_title
    else:
        #print("No base folder title found")
        return "Title unavailable"

def extract_ancestor_data(node, namespaces):
    """extracts ancestor data from each terminal <c>/<cxx> node"""
    ancestors_data = []
    ancestor_count = 0 # for limits to column number count for c ancestors in the spreadsheet

    # Exclude descendant from returned list: ancestors only
    #ancestors = node.iterancestors()
    #ancestors = ancestors[::-1]
    
    ancestors = list(reversed(list(node.iterancestors())))
    ancestors.pop()

    for ancestor in ancestors:
        ancestor_tag = ET.QName(ancestor.tag).localname

        # Match both unnumbered/numbered c tags
        # 06-06-2024: Modified regex back to exclude ns:
        if re.match(r'c\d{0,2}$|^c$', ancestor_tag): 
            is_first_gen_c = ET.QName(ancestor.getparent().tag).localname == 'dsc' # because all first gen 'c'/cxx' are defined here as direct children of <dsc>
            #is_series = ancestor.attrib.get('level') == 'series' # because not all first_gen_c's are an archival/ASpace 'Series'

            did_element = ancestor.find("./ns:did", namespaces=namespaces)
            if did_element is not None:
                unittitle_element = did_element.find("./ns:unittitle", namespaces=namespaces)
                unittitle = " ".join(unittitle_element.itertext()).strip() if unittitle_element is not None else "X" # arbitrary to ensure return type str 

                unitid_element = did_element.find("./ns:unitid", namespaces=namespaces)
                if is_first_gen_c and unitid_element is not None:  # if unitid in first_gen_c, then this c MUST be @level="Series" in ASpace EADs
                    unitid_text = unitid_element.text
                    try:
                        unitid = int(unitid_text)
                        if unitid <= 40:  # Convert to Roman only if unitid is an integer and the integer is up to 40 (this is arbitrary for now because I've not seen a collection with more than 40 series)
                            roman_numeral = convert_to_roman(unitid)
                            ancestors_data.append(f"Series {roman_numeral}. {unittitle}") 
                        else:
                            ancestors_data.append(unittitle) 
                    except ValueError:  # For non-integer unitids, don't convert to Roman; just append the title
                        ancestors_data.append(f"Series {unitid_text}. {unittitle}")  
                else:
                        ancestors_data.append(unittitle)
                    
                ancestor_count += 1

                if ancestor_count >= 5: # '5' for the 5 <cxx> ancestor columns in folder_df
                    break
                
    return ancestors_data

def convert_to_roman(num):
    # Purpose: Converts integer series numbers to Roman numerals in keeping with the traditional presentation of archival series information.
    # I can't seem to get Roman module to work for me. This will do for now because I haven't yet worked with a collection with more than 40 series
    # Return Type: str - The corresponding Roman numeral as a string, or the number itself if not mapped.
    roman_dict = { 1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI', 7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X', 11: 'XI', 12: 'XII', 13: 'XIII', 14: 'XIV', 15: 'XV',  16:'XVI', 17: 'XVII', 18: 'XVIII', 19: 'XIX', 20: 'XX',
                  21: 'XXI', 22: 'XXII', 23: 'XXIII', 24: 'XXIV', 25: 'XXV', 26: 'XXVI', 27: 'XXVII', 28: 'XXVIII', 29: 'XXIX', 30: 'XXX', 31: 'XXXI', 32: 'XXXII', 33: 'XXXIII', 34: 'XXXIV', 35: 'XXXV',  36:'XXXVI', 37: 'XXXVII', 38: 'XXXVIII', 39: 'XXXIX', 40: 'XL'
    }
    return roman_dict.get(num, str(num))


# DATA PROCESSING FOR FOLDER AND BOX MANAGEMENT

def has_explicit_folder_numbering(did_element, containers, ancestor_data, ead_version):
    type_attr = 'type' if ead_version == 'ead2002' else 'localtype'
    container_type_attr = 'altrender' if ead_version == 'ead2002' else 'encodinganalog'
    
    folder_container = next(elem for elem in containers if elem.attrib.get(type_attr, '').lower() == 'folder')
    folder_text = folder_container.text.lower()
    box_number = extract_box_number(did_element, namespaces, ead_version)
    
    # Find the box container or any container with 'Box' as its localtype
    box_container = next((elem for elem in containers if elem.attrib.get(type_attr, '').lower() == 'box' or
                          elem.attrib.get('localtype', '').lower() == 'box'), None)
    
    container_type = box_container.attrib.get(container_type_attr, None) if box_container is not None else None
    base_title = extract_base_folder_title(did_element, namespaces)
    date = extract_folder_date(did_element, namespaces, ead_version)
    ancestor_values = ancestor_data
    ancestor_values += [None] * (5 - len(ancestor_values))
    
    
    start, end = None, None
    if '-' in folder_text:
        start, end = folder_text.split('-')
        start = re.sub(r'\D', '', start)
        end = re.sub(r'\D', '', end)
        start, end = int(start), int(end)
        for i in range(start, end + 1):
            folder_title = f"{base_title} [{i - start + 1} of {end - start + 1}]"
            df_row = [collection_name, call_number, box_number, str(i), container_type] + ancestor_values + [folder_title, date]
            folder_df.loc[len(folder_df)] = df_row
    else:
        folder_number = folder_text
        df_row = [collection_name, call_number, box_number, folder_number, container_type] + ancestor_values + [base_title, date]
        folder_df.loc[len(folder_df)] = df_row

def has_implicit_folder_numbering(did_element, ancestor_data, ead_version):
    type_attr = 'type' if ead_version == 'ead2002' else 'localtype'
    container_type_attr = 'altrender' if ead_version == 'ead2002' else 'encodinganalog'
    
    box_number = extract_box_number(did_element, namespaces, ead_version)
    container_element = did_element.find('./ns:container', namespaces=namespaces)
    container_type = container_element.attrib.get(container_type_attr, None) if container_element is not None else None
    base_title = extract_base_folder_title(did_element, namespaces)
    date = extract_folder_date(did_element, namespaces, ead_version)
    ancestor_values = ancestor_data
    ancestor_values += [None] * (5 - len(ancestor_values))
    
    physdesc_xpath = './ns:physdesc/ns:extent' if ead_version == 'ead2002' else './ns:physdescstructured'
    physdesc_elements = did_element.findall(physdesc_xpath, namespaces=namespaces)
    folder_count = None
    for physdesc in physdesc_elements:
        if ead_version == 'ead2002':
            extent_text = physdesc.text
            integer_match = re.search(r'\b[1-9]\d*\b', extent_text)
            if integer_match:
                folder_count = int(integer_match.group())
                break
        else:
            quantity_elem = physdesc.find('./ns:quantity', namespaces=namespaces)
            if quantity_elem is not None:
                folder_count = int(quantity_elem.text)
                break

    if folder_count is not None:
        if folder_count != 1:
            for i in range(1, folder_count + 1):
                folder_title = f"{base_title} [{i} of {folder_count}]"
                df_row = [collection_name, call_number, box_number, None, container_type] + ancestor_values + [folder_title, date]
                folder_df.loc[len(folder_df)] = df_row
        else:
            df_row = [collection_name, call_number, box_number, None, container_type] + ancestor_values + [base_title, date]
            folder_df.loc[len(folder_df)] = df_row
    else:
        df_row = [collection_name, call_number, box_number, None, container_type] + ancestor_values + [base_title, date]
        folder_df.loc[len(folder_df)] = df_row
      
def prepend_or_fill(column_name, x, idx):
    prefix = "Box " if column_name == 'BOX' else "Folder "
    if pd.notnull(x):  # If the cell has a value, it must be INTEGER. If Box, 
        return prefix + str(x)
    else:  # If the cell is empty
        return prefix + str(idx + 1)
    

# USER INTERACTION AND SELECTION FUNCTIONS

def user_select_collection(collections):
    retry_count = 0
    max_retries = 10

    while retry_count < max_retries:
        try:
            print("\nPlease select a collection to process (or type 'q' to quit): \n")
            for i, collection in enumerate(collections, start=1):
                print(f"{i}. {collection['name']} - {collection['number']}")

            user_input = input("\nChoose number (or type 'q' to exit): \n\n")

            if user_input.lower() == 'q':
                print("\nExiting...Thanks for using the program, and have a great day!\n")
                sys.exit()

            selected_index = int(user_input) - 1
            logging.info(f"User selected index: {selected_index}")

            if 0 <= selected_index < len(collections):
                selected_collection = collections[selected_index]
                return selected_collection

            print("Invalid selection. Please enter a number from the list or 'q' to quit.")
            retry_count += 1

        except ValueError:
            print("Invalid input. Please enter a valid number or type 'q' to quit.\n")
            logging.warning("Invalid input encountered.\n")
            retry_count += 1

    logging.error("Maximum retries reached. Exiting program.")
    print("Too many incorrect attempts. Exiting program. Thanks for your using the program, and have a great day!\n")
    sys.exit()

def display_options(options_list, title):
    ''' Display options in order. This generic function is for both series and box display of options), with a custom display for lists over 30 items. '''
    try:
        logging.info(f"Displaying available {title} options:")
        
        # Determine the prefix and header based on the title
        if title.lower() == "box":
            prefix = "Box "
            header = "\nSelect box(es):"
        else:  # Assume it's series if not box
            prefix = ""
            header = "\nSelect series:"

        print(header)
        print()

        if len(options_list) > 30:
            # Display the first 10 options
            for i in range(10):
                print(f"{i + 1}. {prefix}{options_list[i]}")
            # Display ellipsis lines
            for _ in range(3):
                print("... ...")
            # Display the last 3 options
            for i in range(-3, 0, 1):
                box_number = options_list[i]
                print(f"{len(options_list) + i + 1}. {prefix}{box_number}")
                if box_number == '10001':
                    print("Note: Box 10001 is used as a flag. Please verify box number data before printing labels.")
        else:
            # If 30 or fewer, just display all options
            for i, option in enumerate(options_list, 1):
                print(f"{i}. {prefix}{option}")
                if option == '10001':
                    print("Note: Box 10001 is used as a flag. Please verify box number data before printing labels.")
                    
        print()  # Final newline for formatting
    except Exception as e:
        logging.error(f"Error displaying {title} options: {str(e)}")

def parse_user_input(input_str, options_list):
    ''' The function takes two parameters: input_str and options_list. input_str is a string of comma-separated indices or index ranges, and options_list is a list of options that the user can select from.

# We start by initializing an empty set called selected_options_set. We use a set here because it automatically removes any duplicates. 
# This is particularly useful in scenarios where the user might accidentally specify the same index or range multiple times. 
# For example, if the user inputs "1, 2, 1-3, 2", the resulting selected_options_set will only contain the unique options "option 1", "option 2", and "option 3". Also, in the program if user for instance, 'Would like to specify by SERIES,' if they chose 'option 3' when program asks "Please choose a labeling option,", we give them the flexibility to "Select series by individual numbers, range, or a combination (e.g., '1', '2-3', '4, 5-6') or type 'q' to quit:".
# The set automatically removes the duplicate "option 2" that was specified twice.
# We then split the input_str into individual items using the comma as the delimiter. Each item represents either an index or an index range.
# We iterate over each input_item in the inputs list. For each input_item, we remove any leading or trailing whitespace.
# If the input_item contains a hyphen ("-"), it means it represents an index range. We split the input_item using the hyphen as the delimiter and assign the start and end values to the variables start and end, respectively.
# And then check if both the start and end values are within the bounds of the options_list. If they are, we add the corresponding options from the options_list to the selected_options_set. 
# We adjust the indices by subtracting 1 because Python uses 0-based indexing. For example, if the user specifies "1-3", the actual indices are 0, 1, and 2.
# If the input_item does not contain a hyphen, it means it represents a single index. We convert the input_item to an integer and assign it to the variable index.

# We then check if the index is within the bounds of the options_list. If it is, we add the corresponding option from the options_list to the selected_options_set. Again, we adjust the index by subtracting 1.

# If the index or range selection is out of bounds, we raise a ValueError with an appropriate error message. This is to ensure that the user is aware of the mistake and can correct it.

# Finally, we convert the selected_options_set to a list and return it.
If there is an error during the parsing process, we catch the ValueError, log the error message, print an error message to the user, and return None. This allows the calling function to handle the error gracefully and take appropriate action.
'''
    selected_options_set = set() 
    try:
        inputs = input_str.split(',')
        for input_item in inputs:
            input_item = input_item.strip()
            if '-' in input_item:
                start, end = map(int, input_item.split('-'))
                if 0 < start <= len(options_list) and 0 < end <= len(options_list):
                    selected_options_set.update(options_list[start - 1:end])
                else:
                    raise ValueError("Range selection out of bounds.")
            else:
                index = int(input_item)
                if 0 < index <= len(options_list):
                    selected_options_set.add(options_list[index - 1])
                else:
                    raise ValueError("Selection out of bounds.")
        return list(selected_options_set)
    except ValueError as e:
        logging.error(f"Error parsing user input: {str(e)}")
        print(f"Error: {str(e)}")
        return None


# FILTERING AND DATAFRAME MANAGEMENT

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
            date = datetime.strptime(f'{month} {year}', '%B %Y')
            return (1, date)  # 1 ensures date values are sorted after standard values
        else:
            # Standard series name (no date), sort alphabetically
            return (0, series_name.lower())  # 0 ensures standard values are sorted first

    # Sort series data using the custom sort function
    ordered_series = sorted(series_data, key=sort_key)

    # Append 'Unknown series' only if NaN values were present
    if unknown_series_present:
        ordered_series.append('Unknown series (CAUTION: choosing this might cause unexpected behavior in the program)')

    # Display the series options to the user
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
    num_columns = (len(box_list) + 49) // 50  # Calculate the number of columns needed. Each column will have 50 box options. I can change this to 30 if I want to display 30 boxes per column.
    column_data = [[] for _ in range(num_columns)]  # Create a list of empty lists for each column
    for i, box in enumerate(box_list):  # Iterate over the box_list with index and value
        column_index = i // 50  # Calculate the column index for the current box
        column_data[column_index].append(f"({i+1:{2 if i >= 9 else 1}d})  Box {box}")  # Add the formatted box option to the corresponding column
    
    for i in range(50):  # Iterate over the rows
        row_data = []  # Create an empty list for the current row
        for j in range(num_columns):  # Iterate over the columns
            if i < len(column_data[j]):  # Check if the current row index is within the range of the column data
                row_data.append(column_data[j][i])  # Add the box option to the current row
            else:
                row_data.append(" " * len(column_data[j-1][i]))  # Add empty spaces to align the columns
        print("  ".join(row_data))  # Print the row data with column spacing

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


# MAIL MERGE AND DOCUMENT GENERATION FUNCTIONS

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


# UTILITY AND SORTING FUNCTIONS

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


###################### MAIN ################### MAIN ##################### MAIN #################### MAIN #################### MAIN ####################### MAIN ####################

# Logging config
logging.basicConfig(filename='program_log.txt', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

print("\nHello! Thanks for testing this program: enhancements will be coming soon, so stay tuned!")

# Set the correct working directory based on the execution context
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle (as an executable)
    working_directory = os.path.dirname(sys.executable)
else:
    # Running as a normal Python script
    working_directory = os.path.dirname(os.path.abspath(__file__))

# Initialize namespaces dictionary here to be dynamically set later
namespaces = {}

# Processing EAD files from the working directory
collection_info = process_ead_files(working_directory)

# Give user option to (1) proceed in cases where len(collection) > 1, (2) go back to user_select_collection function menu of collection options to choose a different collection, or (3) exit the program.

if collection_info is not None:
    # Initialize DataFrame structures
    folder_df = pd.DataFrame(columns=['COLLECTION', 'CALL_NO.', 'BOX', 'FOLDER', 'CONTAINER_TYPE',
                                      'C01_ANCESTOR', 'C02_ANCESTOR', 'C03_ANCESTOR', 'C04_ANCESTOR', 'C05_ANCESTOR', 'FOLDER TITLE', 'FOLDER DATES']).astype('object')
    box_df = pd.DataFrame(columns=['REPOSITORY', 'COLLECTION', 'CALL_NO.', 'BOX', 'FOLDER_COUNT', 'FIRST_FOLDER', 'LAST_FOLDER', 'CONTAINER_TYPE', 
                                   'FIRST_C01_SERIES', 'SECOND_C01_SERIES', 'THIRD_C01_SERIES', 'FOURTH_C01_SERIES', 'FIFTH_C01_SERIES']).astype('object')

    # Extract general relevant data
    tree = ET.parse(collection_info["path"])
    root = tree.getroot()
    set_namespace(root)  # Ensure the namespace is set based on this document
    
    
    # Display general collection information
    ead_version = collection_info["ead_version"]
    print(f"\nEAD Version: {ead_version}")
    collection_name = collection_info["name"]
    print(f"\nCollection: {collection_name}")
    repository_name = collection_info["repository"]
    print(f"\nRepository: {repository_name}")
    call_number = collection_info["number"]
    print(f"\nCall Number: {call_number}")
    finding_aid_author = collection_info["author"]
    print(f"\nFinding aid author (ed): {finding_aid_author}")
    
    
    # Use the correct namespace in the find method
    dsc_element = root.find('.//ns:dsc', namespaces=namespaces)
    print(f"\nProcessing {collection_name} : {call_number}")
    
    #print("\nDebug: Namespaces used in dsc_element search:", namespaces)

    # To tell which function was used more (folders numbered vs folder unnumbered)
    has_explicit_folder_numbering_count = 0
    has_implicit_folder_numbering_count = 0
    
    if dsc_element is not None:
        # regex to cater for both numbered/unnumbered c elements
        all_c_elements = [elem for elem in dsc_element.iterdescendants() if re.match(r'c\d{0,2}$|^c$', ET.QName(elem.tag).localname)]

    for elem in all_c_elements:
        try:
            if is_terminal_node(elem):
                did_element = elem.find('.//ns:did', namespaces=namespaces)
                ancestor_data = extract_ancestor_data(did_element, namespaces)
                
                if did_element is not None:
                    
                    containers = [elem for elem in did_element.iterchildren() if ET.QName(elem.tag).localname == 'container']
                    container_count = len(containers)

                    # 06-04-2024: Added check for 'localtype' attribute for EAD3 compatibility
                    has_folder = any(elem.attrib.get('type', '').lower() == 'folder' for elem in containers)  # EAD2002
                    has_folder |= any(elem.attrib.get('localtype', '').lower() == 'folder' for elem in containers)  # EAD3

                    has_box = any(elem.attrib.get('type', '').lower() == 'box' for elem in containers)  # EAD2002
                    has_box |= any(elem.attrib.get('localtype', '').lower() == 'box' for elem in containers)  # EAD3

                    if container_count >= 2 and has_folder:
                        has_explicit_folder_numbering(did_element, containers, ancestor_data, ead_version)
                        has_explicit_folder_numbering_count += 1
                        
                    elif container_count == 1 and has_box:
                        has_implicit_folder_numbering(did_element, ancestor_data, ead_version)
                        has_implicit_folder_numbering_count += 1
    
                    else:
                        has_implicit_folder_numbering(did_element, ancestor_data, ead_version)
                        has_implicit_folder_numbering_count += 1
                        
        except Exception as e:
            # Grab title and date to help identify which <c> element(s) gave trouble while parsing
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

    # if has_explicit_folder_numbering_count > has_implicit_folder_numbering_count:
    #    print(f"Love it when all the folders are numbered!\n")
    
    # Let's store if has_implicit_folder_numbering_count > has_explicit_folder_numbering_count in a variable to use later
    if has_implicit_folder_numbering_count > has_explicit_folder_numbering_count:
        print(f"\nOh boy! The folders have not been numbered; maybe I can help ;)\n")

    # let's ask user if they want implicit folders numbered or not
    # then we decide how to finalize the dfs
    folder_numbering_preference = None
    folders_already_numbered = has_explicit_folder_numbering_count > has_implicit_folder_numbering_count
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
    
    logging.info(f"Preparing dataFrame for {collection_name} folders")
    
    # Finalizing base folder_df:
    
    folder_df['sort_order'] = folder_df['BOX'].apply(box_sort_order)
    folder_df['Folder_temp'] = folder_df['FOLDER'].apply(lambda x: int(re.search(r'\d+', x).group()) if x and re.search(r'\d+', x) else 0)
    folder_df.sort_values(by=['sort_order', 'Folder_temp'], inplace=True)
    
    
    # Finalizing base box_df based on folder numbering/user preferences:
    
    # '1' for continuous, 
    # '2' for non-continuous:
    if folders_already_numbered or folder_numbering_preference == "1":
        folder_df['BOX'] = [prepend_or_fill('BOX', val, idx) for idx, val in enumerate(folder_df['BOX'])]
        folder_df['FOLDER'] = [prepend_or_fill('FOLDER', val, idx) for idx, val in enumerate(folder_df['FOLDER'])]
        logging.info(f"Preparing dataFrame for {collection_name} boxes")
        
        # Finalizing continuous box_df:
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
        # Add "Box " prefix to box numbers
        folder_df['BOX'] = "Box " + folder_df['BOX'].astype(str)
        
        # Assign folder numbers to empty folders
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
        
        # Add "Folder " prefix to folder numbers
        folder_df['FOLDER'] = "Folder " + folder_df['FOLDER'].astype(str)
        
        # Finalizing non-continuous box_df:
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
    
    logging.info(f"Prepping Excel files for mail merge operation")
    folder_dataFrame_path = os.path.join(working_directory, f"{collection_name}_{call_number}_folder.xlsx")
    box_dataFrame_path = os.path.join(working_directory, f"{collection_name}_{call_number}_box.xlsx")

    folder_df.to_excel(folder_dataFrame_path, index=False)
    box_df.to_excel(box_dataFrame_path, index=False)

    excel_file_for_folders = folder_dataFrame_path
    excel_file_for_boxes = box_dataFrame_path
    
else:
    sys.exit()
    
logging.info('Master Excel files for folders and boxes are ready.')

# print(f"Excel files for all {collection_name} folder and box components are ready...\n")

# Check the size of folder_df to determine the next steps
# large size means probably large collection, recommend specifying to user
if len(folder_df) < 1234: # '1234' just cos.
    logging.info('Total folder count fewer than 1,000; proceeding to label generation menu...')

else:
    logging.info('Large collection detected with more than 1000 folders.')
    print("\nWoah! This collection's enormous! I almost lost my mind counting up the folders! Hahaha :D\n")
    print("Consider SPECIFYing your needs for faster processing; otherwise, processing might take longer...")


while True:
    try: # If implicit folder numbering is more than explicit folder numbering, let's give the user option to go back and choose a different numbering preference. If explicit folder numbering is more than implicit folder numbering, let's give the user option to proceed to the label generation menu.
        main_label_menu_choice = input(
            "\nPlease choose a labeling option, or press 'q', and 'Enter' to exit: \n\n"
            "1. Default folder, default box \n"
            "2. Default folder, custom box \n"
            "3. SPECIFY (specify by series/box number(s), choose left labels, custom box labels, combo, default folder/box labels etc?)\n\n"
            # Insert option to go back
            "Note: \n"
            "'Default folder' means left- and right-handed label pairs \n"
            "'Default box' means Paige- or Full Hollinger-size type labels \n"
            "'Custom box' means customized or tailored-to-box types, if available; otherwise, 'Default box' \n\n"  
        )

        has_series_data = folder_df['C01_ANCESTOR'].notna().any()
        
        if main_label_menu_choice == "1": # DEFAULT folder and box label generation.
            wordApp = win32com.client.Dispatch('Word.Application')
            perform_mail_merge(wordApp, [excel_file_for_folders], "default_folder_template.docm", working_directory)
            box_template = "box_template_continuous_numbering.docm" if folders_already_numbered else ("box_template_continuous_numbering.docm" if folder_numbering_preference == "1" else "box_template_non_continuous_numbering.docm")
            perform_mail_merge(wordApp, [excel_file_for_boxes], box_template, working_directory)
            print(f"\nSuccess! Check directory for the output files...")
            break
        
        elif main_label_menu_choice == "2": # DEFAULT folder and CUSTOM box labels
            logging.info("Option 3 selected: # DEFAULT folder and CUSTOM box labels.")
            try:
                wordApp = win32com.client.Dispatch('Word.Application')
                # Default folder mail merge
                perform_mail_merge(wordApp, [excel_file_for_folders], "default_folder_template.docm", working_directory)
                logging.info("Mail merge for default folder labels completed.")
                
                # Read the Excel file into a DataFrame for processing
                custom_df_box = pd.read_excel(excel_file_for_boxes)
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
                                # Check if any part that contains 'h' has a number greater than 2. 2 because if more than 2, it's a tall box
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
                    flat_box_path = os.path.join(working_directory, f"{collection_name}_flat_box.xlsx")
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
                logging.error(f"An error occurred in option 3: # DEFAULT folder and CUSTOM box labels. {str(e)}")    
                   
        elif main_label_menu_choice == "3" and has_series_data: # SPECIFY folder and box label generation
            while True:
                try:
                    wordApp = win32com.client.Dispatch('Word.Application')
                    specify_menu_choice = input("\nWould you like to specify by SERIES or by BOX number?\n\n"
                                                "1. Specify by series\n"
                                                "2. Specify by box number\n"
                                                "3. Exit\n\n")
                    
                    if specify_menu_choice == "1": # by SERIES
                        folder_excel_path, box_excel_path = process_series_selection(folder_df, box_df, working_directory, collection_name, call_number)
                        if folder_excel_path is not None and box_excel_path is not None:
                            label_selection_menu(wordApp, folder_excel_path, box_excel_path, working_directory, folder_numbering_preference, folders_already_numbered, collection_name)
                        else:
                            print("\nSeries selection was exited or invalid...")
                        break
                                
                    elif specify_menu_choice == "2": # by BOX
                        folder_excel_path, box_excel_path = process_box_selection(box_df, folder_df, working_directory, collection_name, call_number)
                        if folder_excel_path is not None and box_excel_path is not None:
                            label_selection_menu(wordApp, folder_excel_path, box_excel_path, working_directory, folder_numbering_preference, folders_already_numbered, collection_name)
                        else:
                            print("\nBox selection was exited or invalid...")
                        break
                        
                    elif specify_menu_choice == "3": # Exit
                        print("\nExiting...Thanks, and have a great day!")
                        sys.exit()
                                
                    else:
                        print("\nInvalid input. Please try again.")
                        
                except Exception as e:
                    logging.error(f"An error occurred during specification choice: {str(e)}")
                
        elif main_label_menu_choice == "3":  # has no series data/not categorized according to series
            print("\nThis finding aid hasn't been categorized by SERIES: you may specify by BOX and/or LABEL type only.")
            wordApp = win32com.client.Dispatch('Word.Application')
            folder_excel_path, box_excel_path = process_box_selection(box_df, folder_df, working_directory, collection_name, call_number)
            if folder_excel_path is not None and box_excel_path is not None:
                label_selection_menu(wordApp, folder_excel_path, box_excel_path, working_directory, folder_numbering_preference, folders_already_numbered, collection_name)
            else:
                print("Box selection was exited or invalid.")
            break
                
        elif main_label_menu_choice == "q":
            print("\nExiting program...Thanks, and have a great day!")
            sys.exit()
            
        else:
            print("\nWrong input. Please enter a valid choice.\n")
            
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")

logging.info('Program finished.')

# Warning to user before they print flagged label
# Check if "10001" is in the 'BOX' column of either folder_df or box_df
if "10001" in folder_df['BOX'].values or "10001" in box_df['BOX'].values:
    print("\nNote before you leave: '10001' was used as a flag for non-standard box numbering in this collection. \nPlease verify and update box data before printing labels.\n")
    print(f"Goodbye!")

input(f"\nPress any key and 'Enter' to exit...")