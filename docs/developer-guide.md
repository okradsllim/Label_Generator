# Developer Guide

## Code structure overview

- main.py: Entry point of the application
- ead_processing.py: Functions for EAD file handling and parsing
- data_extraction.py: Metadata extraction from EAD components
- label_generation.py: Label document creation and mail merge operations
- utils.py: Utility functions and helper methods

## Key functions and their purposes

- process_ead_files(): Identifies and processes EAD files in the working directory
- extract_metadata(): Extracts relevant information from EAD components
- generate_labels(): Creates label documents based on extracted metadata
- perform_mail_merge(): Executes mail merge operation with Word templates

## Libraries used and their roles

- lxml: XML parsing and manipulation
- pandas: Data organization and Excel file creation
- win32com: Interaction with Microsoft Word for mail merge
- os, sys: File and system operations
- logging: Error tracking and debugging

## Extending or modifying the program

- Adding new label types: Create new Word templates and update label_generation.py
- Supporting additional metadata: Modify data_extraction.py to capture new fields
- Enhancing user interface: Improve command-line interactions in main.py

## Contributing to the project

1. Fork the repository on GitHub
2. Create a new branch for your feature or bug fix
3. Make your changes and commit them with clear, descriptive messages
4. Push your changes to your fork
5. Submit a pull request to the main repository
6. Ensure your code follows PEP 8 style guidelines and includes appropriate documentation
