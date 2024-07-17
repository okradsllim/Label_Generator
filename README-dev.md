# Label Generator - Developer Documentation

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Project Structure](#project-structure)
4. [Key Components](#key-components)
5. [Workflow](#workflow)
6. [Functions Overview](#functions-overview)
7. [Data Processing](#data-processing)
8. [User Interaction](#user-interaction)
9. [Mail Merge and Document Generation](#mail-merge-and-document-generation)
10. [Logging](#logging)
11. [Future Enhancements](#future-enhancements)
12. [Troubleshooting](#troubleshooting)

## Introduction

The Label Generator is a Python-based tool designed to create archival box and folder labels from ArchivesSpace EAD2002 XML files. It processes EAD files, extracts relevant data, and generates Microsoft Word documents with labels using mail merge functionality.

## System Requirements

- Python 3.x
- Required Python libraries: lxml, pandas, win32com
- Microsoft Word (with macros enabled)

## Project Structure

- LABEL_GENERATOR/
  - Label_Generator.py
  - default_folder_template.docm
  - box_template_continuous_numbering.docm
  - box_template_non_continuous_numbering.docm
  - left_labels_folder_template.docm
  - vertical_half_holl_continuous_numbering.docm
  - vertical_half_holl_non_continuous_numbering.docm
  - half_horizontal_holl_continuous_numbering.docm
  - half_horizontal_holl_non_continuous_numbering.docm
  - README.md
  - README-dev.md

## Key Components

1. **EAD File Processing**: Parses EAD XML files to extract collection metadata.
2. **Data Extraction**: Extracts box and folder information from the EAD structure.
3. **DataFrame Creation**: Organizes extracted data into pandas DataFrames.
4. **User Interaction**: Provides a command-line interface for user input and selection.
5. **Mail Merge**: Utilizes Microsoft Word's mail merge functionality to generate label documents.

## Workflow

1. Load and process EAD files
2. Extract collection metadata and component information
3. Create and populate DataFrames for folders and boxes
4. Present user with options for label generation
5. Process user selections
6. Generate Excel files for mail merge
7. Perform mail merge to create label documents

## Functions Overview

### EAD Processing

```python
def process_ead_files(working_directory, namespaces):
    # Processes EAD files in the working directory
    # Returns selected collection information

def preprocess_ead_file(file_path):
    # Preprocesses EAD file to handle encoding issues
    # Returns sanitized file path if necessary

def sanitize_xml(input_file_path, output_file_path):
    # Sanitizes XML content by replacing invalid characters
    # Returns dictionary of replaced characters
```

### Data Extraction

```python
def extract_box_number(did_element, namespaces):
    # Extracts box number from EAD element
    # Returns box number as string

def extract_folder_date(did_element, namespaces):
    # Extracts folder date from EAD element
    # Returns date as string

def extract_base_folder_title(did_element, namespaces):
    # Extracts folder title from EAD element
    # Returns title as string

def extract_ancestor_data(node, namespaces):
    # Extracts ancestor data for each terminal node
    # Returns list of ancestor data
```

### Data Processing

```python
def has_explicit_folder_numbering(did_element, containers, ancestor_data=None):
    # Processes components with explicit folder numbering
    # Populates folder_df DataFrame

def has_implicit_folder_numbering(did_element, ancestor_data=None):
    # Processes components with implicit folder numbering
    # Populates folder_df DataFrame

def prepend_or_fill(column_name, x, idx):
    # Prepends 'Box' or 'Folder' to values or fills with incremented values
    # Returns formatted string
```

### User Interaction

```python
def user_select_collection(collections):
    # Presents user with collection selection options
    # Returns selected collection information

def display_options(options_list, title):
    # Displays options for user selection
    # Prints formatted list of options

def parse_user_input(input_str, options_list):
    # Parses user input for series or box selection
    # Returns list of selected options
```

### Mail Merge and Document Generation

```python
def perform_mail_merge(wordApp, excel_files, template_name, working_directory):
    # Performs mail merge operation using Word templates and Excel data
    # Generates label documents

def label_selection_menu(wordApp, folder_excel_path, box_excel_path, working_directory, folder_numbering_preference, folders_already_numbered, collection_name):
    # Presents label selection menu to user
    # Calls appropriate mail merge functions based on user selection
```

## Data Processing

The program uses pandas DataFrames to organize and process data:

- `folder_df`: Stores folder-level information
- `box_df`: Stores box-level information

DataFrames are populated during EAD processing and later used for mail merge operations.

## User Interaction

The program provides a command-line interface for user interaction. Key interaction points include:

1. Collection selection
2. Folder numbering preference
3. Label type selection (default/custom, folder/box)
4. Series or box number specification

## Mail Merge and Document Generation

Mail merge is performed using Microsoft Word templates and Excel files generated from the DataFrames. The `perform_mail_merge` function handles this process, utilizing Word's COM interface through the `win32com` library.

## Logging

The program uses Python's `logging` module to track operations and errors. Log files are saved as `program_log.txt` in the working directory.

## Future Enhancements

1. Implement a graphical user interface (GUI)
2. Add support for EAD3 standard
3. Develop an ArchivesSpace plugin for web interface integration
4. Refactor code to follow object-oriented programming (OOP) paradigm

## Troubleshooting

Common issues and their solutions:

1. **Encoding errors**: Use the `sanitize_xml` function to handle special characters in EAD files.
2. **Mail merge failures**: Ensure Microsoft Word is properly configured with macros enabled.
3. **Unexpected box/folder numbering**: Check the EAD structure and adjust extraction logic if necessary.

For further assistance or to report bugs, please contact [Will](mailto:william.nyarko@yale.edu).
```
