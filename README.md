# Email Processor (OAuth2)
Email processing script for data extraction and loading. 
These functions are designed to perform various tasks related to email processing, SQL querying, data extraction, and error handling. They utilize a variety of libraries for different purposes. Here's a summary of the functions, the libraries used, and the reasons for using them.

## Libraries Used:
- logging: Used for logging information about the script's execution.
- email, imaplib: Used to handle email communication and parsing.
- subprocess: Provides various functions to create and manage additional processes, although it's not used in the provided code.
- MyLogins: Custom module containing login information, presumably for database connections.
- re: Used for regular expression operations.
- os, datetime, csv, zipfile: Used for various file and date/time related operations.
- typing: Provides support for type hints in function signatures.
- BeautifulSoup from bs4: Used for parsing and navigating HTML content.
- ErrorHandler: A custom error handling class defined within the script.
- msal: Microsoft Authentication Library, used for OAuth2 authentication.
- pyodbc: Used for connecting to SQL Server databases.

## Reasons for Using these Libraries:
- Libraries like email, imaplib, BeautifulSoup, pyodbc, and zipfile provide specialized functions that simplify tasks like email handling, HTML parsing, database querying, and file compression.
- The msal library is used for OAuth2 authentication with Microsoft services.
- Custom modules like MyLogins are likely used to store sensitive login information securely.
- The typing library is used for type hints, making the code more readable and understandable.
- The logging library is used for generating logs, which is crucial for debugging and tracking script execution.


## Functions and Their Purpose:
- debug: Print debugging information based on a debug flag.
- generate_auth_string: Generate an authentication string for API requests.
- log: Write a string to the log file and print it to the console.
- check_sender_and_filename_V2: Check if sender and filename are valid by executing an SQL query.
- getfilename4emailsubject: Get filename details based on email subject.
- getheader4table: Get the first column header from the email subject.
- get_allowed_sender: Validate and authorize sender's email.
- insert_datetime: Insert current date and time into the filename.
- extract_email_body_and_preprocess: Extract HTML content from an email, preprocess it, and extract table data.
- convert_and_zip_relevant_data: Convert relevant data to CSV, zip it, and return the zip file path.
- ErrorHandler: A class to manage error messages and error email sending status.
- inform_sender_about_errors: Handle errors by logging them and sending an error email to the sender.
- save_string_to_file: Write a string to a file.
- save_bytes_to_file: Write bytes to a file.
- delete_email_from_server: Delete an email from an IMAP server.
- authenticate_imap: Authenticate to a Microsoft 365 IMAP server using OAuth2.

