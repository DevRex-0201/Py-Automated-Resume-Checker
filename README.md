# Automated Resume and Database Matching System

## Overview

The Automated Resume and Database Matching System is an advanced tool designed to streamline the recruitment process by automatically identifying and scoring potential matches between candidates' resumes and job requirements stored in a database. This Python-based application supports a wide array of document formats, including `.docx`, `.pdf`, and `.doc` for resumes, as well as `.xlsx` and `.csv` for database entries. It leverages several powerful libraries to manipulate data, process documents, and provide a user-friendly graphical interface for easy operation. Ideal for HR departments and recruitment agencies, this tool significantly reduces manual workload by automating the match-making process.

## Key Features

- **Comprehensive File Format Support**: Seamlessly processes resumes in `.docx`, `.pdf`, and `.doc` formats, and interacts with databases in `.xlsx` or `.csv` formats.
- **Intelligent Automated Matching**: Employs a robust algorithm to match resumes with database entries based on key identifiers such as email, phone number, mobile number, and name.
- **Adaptive Scoring Mechanism**: Incorporates a dynamic scoring system to evaluate and prioritize matches, ensuring the most relevant candidates are highlighted.
- **Graphical User Interface**: Features an intuitive GUI built with Tkinter, enabling straightforward navigation and operation without the need for command-line interaction.
- **Detailed Logging and Error Tracking**: Maintains comprehensive logs for monitoring the system's operations and troubleshooting errors.
- **Database Synchronization**: Automatically updates the database with match results, including the identification of relevant documents and the calculation of match scores.

## Installation Requirements

Ensure your system has Python 3.6 or later installed. The system also requires the installation of specific dependencies to function properly.

### Required Libraries

- pandas
- python-docx
- PyMuPDF (fitz)
- pypiwin32 (only for Windows users, for win32com support)
- Tkinter (typically included with Python)

### Dependency Installation

Execute the following command to install the necessary libraries:

```shell
pip install pandas python-docx PyMuPDF pypiwin32
```

## How to Use

1. **Starting the Application**: Execute the script to open the graphical user interface.
2. **Folder Selection**: Click "Browse" to choose the directory containing the candidate resumes.
3. **Database File Selection**: Click "Browse" to select the database file (.xlsx or .csv) against which the resumes will be matched.
4. **Initiate Processing**: Press "Start Processing" to begin the matching process. Upon completion, the system will update the database with the results and display a notification.

### Graphical User Interface (GUI) Components

- **Folder Selection**: Allows users to specify the location of the resumes.
- **Database File Selection**: Enables users to choose the database file for matching.
- **Processing Button**: Triggers the automated matching and updating process.

## System Mechanics

The system reads each resume from the chosen folder and extracts text to compare against database entries. Matching is performed based on specified criteria, with each match receiving a score reflecting its relevance. The database is subsequently updated to include these scores along with the associated resume files and any identified Contracts of Service.

## Troubleshooting and Support

If you encounter any issues:

- Consult the `app.log` file for error details.
- Ensure all dependencies are properly installed.
- Confirm that the documents are in supported formats and correctly structured.

## Contributions

We welcome contributions to enhance the system's functionality or address bugs. Please adhere to established coding practices and submit pull requests for any proposed changes.

## License

[Specify the project's license here, usually an open-source license for GitHub projects]

This README aims to provide a detailed introduction to the Automated Resume and Database Matching System, outlining its capabilities, installation procedure, usage guidelines, and support framework. Adjust the license information to align with your project's licensing terms.