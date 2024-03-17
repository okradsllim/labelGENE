Introduction
============

This project is a Python script for processing EAD XML files to produce box and folder labels. It focuses on parsing, data extraction, and metadata management. The script identifies EAD files in a specified directory, sanitizes the XML files to address common encoding issues, and parses them using the lxml library for efficient XML processing.

Key metadata elements such as repository details, collection names, call numbers, and specific hierarchical data are systematically extracted. The script traverses and stores data from every terminal 'c' node, effectively capturing item-level descriptions across a variety of valid EAD2002 files from Beinecke ArchivesSpace.

The script can handle both explicit and implicit folder numbering, accommodating the varied structuring of EADs. Explicit folder numbering is managed by extracting folder numbers directly from the XML, while implicit numbering involves a nuanced approach to deduce folder counts from available metadata.

The script is complemented by a suite of .docm template Word files, each embedded with VBA code tailored to work in tandem with the script's mail_merge operation function. This integration facilitates the generation of folder and box label files, aligning with the user's specific labeling preferences.