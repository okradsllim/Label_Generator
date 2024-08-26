# Technical Description

## System architecture

The Label Generator is a Python-based application that processes EAD XML files, extracts relevant metadata, and generates label documents using Microsoft Word templates. It employs a modular structure with separate functions for XML parsing, data extraction, and label generation.

## XML processing methodology

### EAD2002 support
- Uses lxml library for efficient XML parsing
- Navigates EAD structure using XPath expressions
- Extracts metadata from <archdesc> and <dsc> sections

### EAD3 support
- Detects EAD3 namespace and adjusts XPath queries accordingly
- Handles differences in element names and attribute structures between EAD2002 and EAD3

## Integration with ArchivesSpace

- Processes EAD XML files exported directly from ArchivesSpace
- Supports both published and unpublished versions of finding aids
- Handles various levels of description and component structures

## Metadata extraction process

- Recursively traverses <dsc> section to process all component levels
- Extracts box numbers, folder numbers, titles, dates, and hierarchical information
- Handles both explicit and implicit folder numbering scenarios

## Label generation workflow

1. Organizes extracted metadata into pandas DataFrames
2. Generates Excel spreadsheets as data sources for mail merge
3. Uses win32com library to interact with Microsoft Word
4. Applies appropriate Word templates based on user-selected options
5. Performs mail merge to create final label documents

## Performance considerations

- Implements error handling and logging for robust operation
- Uses efficient XML parsing techniques to handle large EAD files
- Provides options for processing specific series or box ranges to manage resource usage
