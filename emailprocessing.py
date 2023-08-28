import logging
import email
import imaplib
import subprocess
from MyLogins import *
import re
import os
import datetime
import csv
import zipfile
from typing import List
from bs4 import BeautifulSoup
from email.utils import parseaddr
from bs4 import Comment
import sys
import msal

# 2023-04-30 Bruna: Approach with pyodbc and parameterized queries to avoid issues with quotes and apostrophes
import pyodbc

# 2023-03-16 Bruna: Code debugging
DEBUGGING_ON = True

def debug(value):
    """
    Print the value as debugging information, if debugging is turned on.

    Args:
        value (Any): The value to print as debugging information.

    Returns:
        None
    """
    if DEBUGGING_ON == True:
        print(f"Debugging info: {value}")


# from https://stackoverflow.com/questions/73368376/use-imap-with-xoauth2-and-ms365-application-permissions-to-authenticate-imap-for
def generate_auth_string(user: str, token: str) -> str:
    """
    Returns an authentication string for an API request.

    Args:
        user (str): The username for the API request.
        token (str): The access token for the API request.

    Returns:
        str: The authentication string for the API request, in the format "user={user}\x01auth=Bearer {token}\x01\x01".
    """
    return f"user={user}\x01auth=Bearer {token}\x01\x01"


def log(string: str) -> None:
    """
    Writes a string to the log file and prints it to the console.

    Args:
        string (str): The string to be logged.

    Returns:
        None
    """
    with open(logFullPath, "a") as myfile:
        myfile.write(str(datetime.datetime.now()) + " " + string + "\n")
    print(string)

logger = logging.getLogger(__name__)

def check_sender_and_filename_V2(filename, sender):
    """
    Check if the sender and filename are valid by executing an SQL query.

    Parameters:
        filename (str): The name of the file to be checked.
        sender (str): The email address of the sender.

    Returns:
        bool: True if the check succeeded, False otherwise.
    """
    try:
        # Connect to the SQL Server database
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=ZEUWNLDODH01;DATABASE={MyLogins.sql_db};UID={MyLogins.sql_username};PWD={MyLogins.sql_password}"
        conn = pyodbc.connect(conn_str)

        # Create a cursor to execute the query
        cursor = conn.cursor()

        # Use parameterized queries to avoid issues with quotes and apostrophes
        # Execute the SQL function and fetch the result
        cursor.execute(
            """
        SELECT COUNT(*) FROM dbo.InputFilesMatchingRules WHERE (Begining = LEFT(?, LEN(Begining))) AND (Ending = RIGHT(?, LEN(Ending))) AND (AllowedEmailSenders LIKE N'%' + ? + N'%')""",
            (filename, filename, sender),
        )
        result = cursor.fetchone()

        debug(f"Result from function is: {bool(result[0])}")

        # Unpack the result tuple and return it as a boolean value
        return bool(result[0])

    except pyodbc.Error as e:
        debug("Error executing SQL query: " + str(e))
        return False

    finally:
        # Close the database connection in the finally block
        if conn:
            conn.close()
            debug("Database connection closed.")

def getfilename4emailsubject(subject):
    """
    Get the filename details for the given email subject so that the created file can be stored under the expected name.
    Function to be used when there are tables present in the email that need to be extracted and processed.

    Parameters:
        subject (str): The subject of the email.

    Returns:
        str: The filename details (beginning and ending) as determined by the SQL table InputFilesMatchingRules.
    """
    try:
        # Connect to the SQL Server database
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=ZEUWNLDODH01;DATABASE={MyLogins.sql_db};UID={MyLogins.sql_username};PWD={MyLogins.sql_password}"
        conn = pyodbc.connect(conn_str)

        # Create a cursor to execute the query
        cursor = conn.cursor()

        # Use parameterized queries to avoid issues with quotes and apostrophes
        # Execute the SQL function and fetch the result
        cursor.execute(
            """
        SELECT Begining, Ending FROM dbo.InputFilesMatchingRules WHERE RIGHT(?, len(EmailSubject)) = EmailSubject""",
            (subject),
        )
        result = cursor.fetchone()

        debug("Fetching filename from email subject...")
        debug(f"Result from getfilename4emailsubject function is: {result}")

        if result:
            return result[0]  # Return the first column of the result (FileNameDetails)
        else:
            return None  # Return None if no result is found

    except pyodbc.Error as e:
        debug("Error executing SQL query: " + str(e))
        return None
    finally:
        # Close the database connection in the finally block
        if conn:
            conn.close()
            debug("Connection to the database closed.")


# 2023-08-24 Bruna: Get first cell from table header 
def getheader4table(subject):
    """
    Get the beginning column of the first column header for the given email subject to easily identify where the relevant data starts.
    Function to be used when there are tables present in the email that need to be extracted and processed.

    Parameters:
        subject (str): The subject of the email.

    Returns:
        str: The first column header as determined by the SQL table InputFilesMatchingRules.
    """
    try:
        # Connect to the SQL Server database
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=ZEUWNLDODH01;DATABASE={MyLogins.sql_db};UID={MyLogins.sql_username};PWD={MyLogins.sql_password}"
        conn = pyodbc.connect(conn_str)

        # Create a cursor to execute the query
        cursor = conn.cursor()

        # Use parameterized queries to avoid issues with quotes and apostrophes
        # Execute the SQL function and fetch the result
        cursor.execute(
            """
        SELECT EmailFirstColumnHeader FROM dbo.InputFilesMatchingRules WHERE RIGHT(?, len(EmailSubject)) = EmailSubject""",
            (subject),
        )

        debug("Fetching First Column Header...")
        result = cursor.fetchone()

        debug(f"Result from getheader4table function is: {result}")

        if result:
            debug(f"First Column Header = {result[0]}")
            return result[0]  # Return the first column of the result (EmailFirstColumnHeader)
        else:
            return None  # Return None if no result is found

    except pyodbc.Error as e:
        debug("Error executing SQL query: " + str(e))
        return None
    finally:
        # Close the database connection in the finally block
        if conn:
            conn.close()
            debug("Connection to the database closed.")


# 2023-04-06 Bruna: Check email sender
def get_allowed_sender(
    filename: str, frm: str, emailid: int, subject: str, M: imaplib.IMAP4_SSL
) -> str:
    """
    Returns the email address of the sender if it is valid and authorized to send the file.

    Parameters:
    filename (str): The name of the file to be checked.
    frm (str): The 'From' header field of the email message.
    emailid (int): The unique identifier of the email message in the mailbox.
    subject (str): The subject line of the email.
    M (imaplib.IMAP4_SSL): The IMAP client object used to access the mailbox.

    Returns:
    str: The email address of the sender if it is valid and authorized to send the file.
    """

    # Extract sender email address from the 'From' header
    if not frm:
        return ""

    if frm.find("<") > 0 and frm.find(">") > 0:
        # debug("From with <>")
        sender = frm[frm.find("<") + 1 : frm.find(">")]
    else:
        # There seems to be no <> in the frm
        # debug("From without <>")
        if frm.find("[") > 0 and frm.find("]") > 0:
            # debug("From with []")
            sender = frm[frm.find("[") + 1 : frm.find("]")]
        else:
            # There seems to be also no [] in the frm
            # We must assume that the from field is a valid email address
            # debug("From without []")
            sender = frm

    # Check if the sender email address is valid
    if not re.match(r"[^@]+@[^@]+\.[^@]+", sender):
        # debug('Email format not valid.')
        delete_flag = delete_email_from_server(emailid, authenticate_imap(conf))
        if delete_flag:
            log(f"Invalid email from invalid email address {sender} deleted.")
            sys.exit()
        else:
            # debug(f"Email from {sender} already deleted.")
            sys.exit()

    # Check if the sender email address ends with the allowed domains
    elif not sender.endswith("@nokia.com") and not sender.endswith("@fieldglass.net"):
        debug("Domain from the sender email address is not allowed.")
        ErrorHandler.add_error_message(
            f"Domain from the {sender} email address is not an allowed domain."
        )
        mailbody_content = ErrorHandler.get_mailbody_content()
        debug(
            f"Inform sender function value before calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}"
        )

        # Call inform_sender_about_errors with the updated mailbody_content
        if inform_sender_about_errors(
            filename, sender, subject, mailbody_content, error_handler=error_handler
        ):
            log(
                f"Error in email ({filename}), sent from {sender}. This email is not from an allowed domain."
            )
            debug(
                f"Error in files sent from {sender}. This email is not an allowed domain."
            )
            debug(
                f"Inform sender function value after calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}"
            )

            # Delete the email from the server after the error message has been sent successfully
            delete_flag = delete_email_from_server(emailid, authenticate_imap(conf))
            if delete_flag:
                log(f"Invalid email from unauthorized {sender} deleted.")
                sys.exit()
            else:
                debug(f"Email from {sender} already deleted.")
                sys.exit()

    return sender

# 2023-04-12 Bruna: Insert DATETIME into FILENAME
def insert_datetime(filename):
    """
    Inserts the current date and time into the filename before the file extension, except for csv files.

    Args:
    - filename (str): The original filename to be modified.

    Returns:
    - new_filename (str): The new filename with the current date and time inserted before the file extension.
    """

    debug("Insert datetime function working.")
    debug(f"Filename before datetime function: {filename}")

    # Get the current datetime
    now = datetime.datetime.now()

    # Split the filename into base name and file extension
    base_name, file_extension = os.path.splitext(filename)

    debug(file_extension)

    # Klaus 2023-04-19
    # Create a new filename with datetime inserted before the file extension, except for csv files
    if file_extension != ".csv":
        new_filename = f"{base_name}_{now.strftime('%Y%m%d_%H%M%S')}{file_extension}"
        filename = new_filename
        debug(f"New filename: {filename}")

    return filename

def extract_email_body_and_preprocess(mail: email.message.Message) -> str or None:
    """
    Extracts the HTML content from an email, preprocesses it, and extracts relevant table data.

    Parameters:
        mail (email.message.Message): The email message object.

    Returns:
        str or None: A string containing the preprocessed email content if relevant table data is found,
        or None if no relevant table data is found or no HTML content is present.
    """
    debug("Extract email body and preprocess function initiated...")
    email_body = None
    
    # Iterate through the email parts to find the HTML part
    for part in mail.walk():
        content_type = part.get_content_type()
        debug(f"Email content type = {content_type}")
        if content_type == "text/html":
            payload = part.get_payload(decode=True)
            charset = part.get_content_charset() or 'utf-8'
            email_body = payload.decode(charset)
            break
    
    if email_body:
        debug("Preprocessing HTML content...")
        soup = BeautifulSoup(email_body, 'lxml')  # Use the 'lxml' parser
        
        # Remove the <head>...</head> content from the email
        head_tag = soup.find("head")
        if head_tag:
            head_tag.decompose()
        
        # Get the starting point based on the header obtained from SQL
        start_string = getheader4table(mail["Subject"])
        debug(f"Start string = {start_string}")
        
        if start_string:
            relevant_data = []

            # Find the tag containing the specified string
            start_tag = soup.find(string=start_string)
            debug(f"Start tag = {start_tag}")
            
            if start_tag:
                # Traverse upwards to find the relevant row
                tr_tag = start_tag.find_parent("tr")
                
                # Regular expression to remove non-alphanumeric characters, except possibly relevant characters 
                regex = re.compile(r'[^a-zA-Z0-9\s\.\-|:()%&+]+')

                # Process the header row
                relevant_data.append([regex.sub('', cell.text).replace(',', '').strip() for cell in tr_tag.find_all(["td", "th"])])

                # Process the remaining rows
                for row in tr_tag.find_next_siblings("tr"):
                    cells = row.find_all(["td", "th"])
                    row_data = [regex.sub('', cell.text).replace(',', '').strip() for cell in cells]
                    relevant_data.append(row_data)

                if relevant_data:
                    debug("Relevant Data Extracted:")
                    debug(relevant_data)
                    # Return the formatted data as a string
                    return "\n".join(",".join(row) for row in relevant_data)
                else:
                    debug("No relevant table data found in the email")
                    return None
            else:
                debug("Starting point not found in the email")
                return None
        else:
            debug("No relevant header found for the email subject")
            return None
    else:
        debug("No HTML content found in the email")
        return None


def convert_and_zip_relevant_data(relevant_data: str, detach_dir: str, subject: str) -> str or None:
    """
    Converts relevant data to a CSV file and zips it.

    Parameters:
        relevant_data (str): Relevant data extracted from the email.
        detach_dir (str): The path to the directory where the zip file will be saved.
        subject (str): The subject of the email.

    Returns:
        str or None: The path to the generated ZIP file if successful,
        or None if no relevant data is found.
    """
    debug("Converting extracted data to CSV file...")

    # Split the relevant data into rows and cells
    debug(f"Relevant Data being converted into CSV format = {relevant_data}")
    rows = relevant_data.split("\n")
    table_data = [row.split(",") for row in rows]

    # Generate a CSV filename based on the email subject
    csv_filename = insert_datetime(getfilename4emailsubject(subject)) + ".csv"

    # Create the full CSV file path
    csv_filepath = os.path.join(detach_dir, csv_filename)

    # Write the table data to the CSV file
    with open(csv_filepath, "w", newline="", encoding="utf-8") as csv_file:
        csv_writer = csv.writer(csv_file)

        # Write each row to the CSV file
        for row in table_data:
            csv_writer.writerow(row)

    # Generate a ZIP filename based on the email subject
    zip_filename = insert_datetime(getfilename4emailsubject(subject)) + ".zip"

    # Create the full ZIP file path
    zip_filepath = os.path.join(detach_dir, zip_filename)

    debug(f"Zipping generated CSV file... ZIP Filepath = {zip_filepath}...")
    
    # Create a ZIP archive and add the CSV file to it
    with zipfile.ZipFile(zip_filepath, "w") as zip_file:
        
        # Add the CSV file to the zip archive
        zip_file.write(csv_filepath, os.path.basename(csv_filepath))

    # Delete the CSV file after it's zipped
    os.remove(csv_filepath)

    # Return the path to the generated ZIP file
    return zip_filepath


# 2023-04-18 Bruna: Add the mailbody_content parameter for error clarification in email sent to sender
# 2023-04-30 Bruna: Ensure the sender is being informed only once
# 2023-05-01 Bruna: Add ErrorHandler class to store the error messages as an attribute
# 2023-05-02 Implement class method
class ErrorHandler:
    error_messages = []
    error_email_sent = False

    @classmethod
    def add_error_message(cls, error_message: str) -> None:
        """
        Adds an error message to the list of error messages.

        Parameters:
        error_message (str): The error message to add.
        """
        cls.error_messages.append(error_message)

    @classmethod
    def get_mailbody_content(cls) -> str:
        """
        Concatenates all error messages together into the email body content.

        Returns:
        str: The email body content with all error messages included.
        """
        if not cls.error_messages:
            return ""

        return (
            "<b>ERROR: The email sent to reporting.rso@nokia.com reported an issue.</b><br><br>"
            + "<br>".join(cls.error_messages)
            + "<br><br>If you still get this email, please contact: <b>bruna.duarte@nokia.com</b> "
            + "for further assistance."
        )

    @classmethod
    def set_error_email_sent(cls, sent: bool) -> None:
        """
        Sets the error email sent flag.

        Parameters:
        sent (bool): The value to set the flag to.
        """
        cls.error_email_sent = sent

    @classmethod
    def get_error_email_sent(cls) -> bool:
        """
        Gets the error email sent flag.

        Returns:
        bool: The value of the flag.
        """
        return cls.error_email_sent


def inform_sender_about_errors(
    filename: str,
    sender: str,
    subject: str,
    mailbody_content: str,
    error_handler: ErrorHandler,
) -> bool:
    """
    Handles an error by logging it and sending an error email to the sender.

    Parameters:
    filename (str): The name of the file that caused the error.
    sender (str): The email address of the sender.
    subject (str): The subject of the email.
    mailbody_content (str): The email body content with all error messages included.
    error_handler (ErrorHandler): The instance of the ErrorHandler class to retrieve the error messages.

    Returns:
    bool: True if the function executes successfully, False otherwise.
    """
    try:
        debug("Inform sender about errors function initiated...")
        # debug(f"error_handler.get_error_email_sent() value when executing the inform_sender_about_errors function is {error_handler.get_error_email_sent()}")
        # debug(f"mailbody_content value when executing the inform_sender_about_errors function is {mailbody_content}")
        # debug(f"ErrorHandler.get_mailbody_content() value when executing the inform_sender_about_errors function is {ErrorHandler.get_mailbody_content()}")

        if not mailbody_content or not ErrorHandler.get_mailbody_content():
            debug("Mail body content for sender error informing is initially empty.")
            return False

        # If the email has already been sent, return True
        if error_handler.get_error_email_sent():
            return True

        # Use the SQL procedure to send the email to the sender
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=ZEUWNLDODH01;DATABASE={MyLogins.sql_db};UID={MyLogins.sql_username};PWD={MyLogins.sql_password}"
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Use parameterized queries to avoid issues with quotes and apostrophes
        cursor.execute(
            "EXECUTE RSOEUR_Master.dbo.[RP_Check_EmailHandling_Errors] @SubjectSender=?, @EmailSender=?, @mailbody_content=?",
            subject,
            sender,
            mailbody_content,
        )
        conn.commit()

        # Set the error email sent flag
        error_handler.set_error_email_sent(True)

        debug("Error handled successfully!")
        return True

    except Exception as e:
        debug(f"Error occurred in inform_sender_about_errors function: {e}")
        return False
    
    finally:
        # Close the database connection in the finally block
        if conn:
            conn.close()
            debug("Connection to the database closed.")



def save_string_to_file(string: str, filepath: str) -> None:
    """
    Writes a string to a file.

    Parameters:
    string (str): The string to write to the file.
    filepath (str): The path to the file.

    Returns:
    None
    """
    fp = open(filepath, "w")
    fp.write(string)
    fp.close()


def save_bytes_to_file(bytes: bytes, filepath: str) -> None:
    """
    Writes bytes to a file.

    Parameters:
    bytes (bytes): The bytes to write to the file.
    filepath (str): The path to the file.

    Returns:
    None
    """
    fp = open(filepath, "wb")
    fp.write(bytes)
    fp.close()


# 2023-04-28 Bruna: Function to permanently delete email from the server
def delete_email_from_server(emailid, M):
    """
    Delete an email from an IMAP server given the email ID.

    Args:
    - emailid (str): the ID of the email to delete
    - M (imaplib.IMAP4_SSL): the IMAP server instance

    Returns:
    - delete_flag (bool): True if the email was deleted from the server, False otherwise
    """
    debug("Running delete_email_from_server function...")
    try:
        # Select the mailbox
        M.select()

        # Set the delete flag for the specified email ID
        M.store(emailid, "+FLAGS", "\\Deleted")

        # Permanently remove any emails with the delete flag
        M.expunge()

        # Check if the email with the specified ID was deleted from the server
        response = M.search(None, "UID", emailid)
        # if emailid is not found, it was already deleted
        delete_flag = False if response[0].split() else True

        # Print a debug message indicating if the email was deleted
        if delete_flag:
            debug(f"Email {emailid} successfully deleted from server.")
        else:
            debug(f"Email {emailid} not found on server, already deleted.")

        return delete_flag

    except Exception as e:
        debug(f"Error occurred in delete_email_from_server function: {e}")
        return log(f"Error occurred while deleting email {emailid}: {e}")


def authenticate_imap(conf):
    """
    Authenticates to a Microsoft 365 IMAP server using OAuth2 and returns an authenticated IMAP4_SSL object.

    Parameters:
        conf (dict): a dictionary with the following keys:
            - 'authority': the URL of the Azure AD authority for your tenant.
            - 'client_id': the client ID of your application registered in Azure AD.
            - 'scope': a list of scopes to request access to (usually ['https://outlook.office.com/.default']).
            - 'secret': the client secret of your application in Azure AD.

    Returns:
        An authenticated IMAP4_SSL object connected to the Microsoft 365 IMAP server.

    Raises:
        Any exceptions raised by the underlying MSAL or imaplib libraries.
    """

    app = msal.ConfidentialClientApplication(
        conf["client_id"], authority=conf["authority"], client_credential=conf["secret"]
    )
    result = app.acquire_token_silent(conf["scope"], account=None)

    result = None

    result = app.acquire_token_silent(conf["scope"], account=None)

    if not result:
        result = app.acquire_token_for_client(scopes=conf["scope"])

    if "access_token" in result:
        print("Token successfully acquired")
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))

    M = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    print("Connected")

    # from https://stackoverflow.com/questions/73368376/use-imap-with-xoauth2-and-ms365-application-permissions-to-authenticate-imap-for
    M.authenticate(
        "XOAUTH2",
        lambda x: generate_auth_string(
            "reporting.rso@nokia.com", result["access_token"]
        ),
    )

    debug("Logged in")

    M.select()
    return M


# Set the log file path
logFullPath = "C:/RSOData/Data/RSOReportingEmailProcessor.log"

# Set the directory where attachments will be saved
detach_dir = "c:/RSOData/Data/Input/"

# Logging start of the script
log("--- Running: Reporting Platform IMAP Emails Extractor ---")

# Configure the authentication credentials
# Source: https://stackoverflow.com/questions/73368376/use-imap-with-xoauth2-and-ms365-application-permissions-to-authenticate-imap-for
conf = {
    "authority": "https://login.microsoftonline.com/5d471751-9675-428d-917b-70f44f9630b0",
    "client_id": "2e41c712-2411-4a01-9955-e45bcc026ce4",
    "scope": ["https://outlook.office.com/.default"],
    "secret": "UH-8Q~Y98tUeTZ.fPSxQPSQ6SjudFDsG8qsYGbiH",
    "secret-id": "26a62841-08ff-4418-953e-28a8a9604b20",  # for documentation only
    "tenant-id": "5d471751-9675-428d-917b-70f44f9630b0",
}

# Authenticate using IMAP and OAuth2
M = authenticate_imap(conf)

# Print enabled debugging
debug("Debug function is switched on. Debugging info will be shown as the script runs.")

# Fetch email IDs based on IMAP search criteria
resp, items = M.search(None, "ALL")
items = items[0].split()  # getting the mails id

# 2023-03-03 Bruna: Implementing Error Handling
# Initialize variables for error handling
mailbody_content = ""
error_handler = ErrorHandler()
filename = None
relevant_data = None
attachment_found = False
debug(f"Error handling variables initiated... Filename = {filename}, Relevant data = {relevant_data}...")

# Iterate through email IDs
for emailid in items:
    # Initialize variables for error handling
    mailbody_content = ""
    error_handler = ErrorHandler()
    filename = None
    relevant_data = None
    attachment_found = False
    debug(f"Error handling variables initiated: Filename = {filename}, Relevant data = {relevant_data}, Attachment found = {attachment_found}...")
    debug("For emailid loop initiated... Fetching and parsing email data...")

    # Fetch email data and parse it
    resp, data = M.fetch(emailid, "(RFC822)")
    email_object_string = data[0][1]
    email_object_string = email_object_string.decode("utf-8")
    
    mail = email.message_from_string(email_object_string)
    debug("["+mail["From"]+"] :" + mail["Subject"])
    frm = mail["From"]
    debug('From: '+frm)
    subject = mail["Subject"]

    # Check if the email contains attachments
    if mail.get_content_maintype() == 'multipart':
        debug("Extracting email multipart...")

        # Iterate through parts of the email
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                attachment_found = True
                debug(f"Attachment found = {attachment_found}")
                continue

            # Skip non-attachment parts
            if part.get('Content-Disposition') is None:
                continue
            
            filename = part.get_filename()
            debug(f'The filename is: {filename}.')                 
        
    # Iterate through parts of the email again to find HTML content
    elif filename is None and not attachment_found:
        for part in mail.walk(): 
            if part.get_content_type() == "text/html" and not attachment_found:
                # Extract and preprocess the HTML content of the email
                if extract_email_body_and_preprocess(part):
                    relevant_data = extract_email_body_and_preprocess(part)
            
            if relevant_data is not None and getheader4table:
                debug("Relevant data found in email body...")
                # Call the function to convert and zip the relevant data
                zip_filepath = convert_and_zip_relevant_data(relevant_data, detach_dir, subject)

                if zip_filepath:                               
                    # Delete the email from the server after successful ZIP file generation
                    if delete_email_from_server(emailid, authenticate_imap(conf)):
                        log(f"Email with relevant data from {sender} deleted after zip file was generated successfully.")
                        sys.exit()
                    else:
                        debug(f"Email from {sender} with ID {emailid} already deleted.")
                else:
                    debug("No zip filepath found by the script. No files were imported.")
        

    # Rest of the code for processing other cases    
    # Get the sender's email address and verify its validity
    sender = get_allowed_sender(filename, frm, emailid, subject, M)

    if filename is not None and get_allowed_sender(filename, frm, emailid, subject, M):
        debug(f"The verified and validated sender email address is: {sender}")
        continue
    else:
        log(f"The sender email address ({sender}) is invalid for file processing.")
        
        # Handle empty emails / without attachments           
        ErrorHandler.add_error_message(
            "The previously sent email contains no content or recognized filename."
        )
        mailbody_content = ErrorHandler.get_mailbody_content()

        # Call inform_sender_about_errors with the updated mailbody_content
        if inform_sender_about_errors(
            filename, sender, subject, mailbody_content, error_handler=error_handler
        ):
            debug(f"Empty email sent from {sender}.")
            debug(f"Inform sender function value after calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

            # Delete the email from the server after the error message has been sent successfully
            delete_flag = delete_email_from_server(emailid, authenticate_imap(conf))
            if delete_flag:
                log(f"Empty email from {sender} deleted.")
            else:
                debug(f"Email from {sender} already deleted.")
                sys.exit()
        else:
            if delete_email_from_server(emailid, authenticate_imap(conf)):
                log(f"Email from {sender} deleted.")
            else:
                log(f"Email from {sender} with ID {emailid} already deleted.")
                
            log(f"Informing sender unsuccessful: Error while sending email to {sender}.")
            debug(f"Informing sender unsuccessful: Error while sending email to {sender}.")
            sys.exit()

    ### Handle allowed file extensions in case email contains attachments        
        debug(f"Filename is {filename}, which means attachment found = {attachment_found}")

    if filename is not None and attachment_found:
        debug("Checking file extension validity...")
        # List of invalid and valid file extensions
        invalid_extensions = [".png", ".jpeg", ".jpg"]
        valid_extensions = [".xls", ".xlsx", ".zip", ".csv"]
        invalid_files = []

        file_extension = os.path.splitext(filename)[1]

        if (
            file_extension in invalid_extensions
            or file_extension not in valid_extensions
        ):
            invalid_files.append(filename)

        if invalid_files:
            ErrorHandler.add_error_message(
                f"A file provided ({filename}) contains invalid extensions and cannot be processed."
            )
            mailbody_content = ErrorHandler.get_mailbody_content()

            debug(f"Inform sender function value before calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

            # Call inform_sender_about_errors with the updated mailbody_content
            if inform_sender_about_errors(
                filename, sender, subject, mailbody_content, error_handler=error_handler
            ):
                log(f"Error in files ({filename}), sent from {sender}. Invalid extension {file_extension}.")
                debug(f"Error in files sent from {sender}.")
                debug(f"Inform sender function value after calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

                # Delete the email from the server after the error message has been sent successfully
                delete_flag = delete_email_from_server(emailid, authenticate_imap(conf))
                if delete_flag:
                    log(f"Invalid email from {sender} deleted.")
                else:
                    debug(f"Email from {sender} already deleted.")
                    sys.exit()

        try:
            debug("Running TRY statement.")
            # If there are valid attachments present in the email, they will be processed here
            # 2023-08-02 Klaus: used V2 
            if relevant_data is None and attachment_found and check_sender_and_filename_V2(filename, sender):
                debug("No tables found in email body. Continuing script.")
                filename = insert_datetime(filename)
                debug(f"The updated filename is: {filename}.")
                att_path = os.path.join(detach_dir, filename)  # /Input/Filename

                # Check if it already exists in this path
                if not os.path.isfile(att_path):
                    # Write the file onto the Input folder
                    fp = open(att_path, "wb")
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    debug("The script inserted the file onto the input folder.")
                    log(f"Script completed successfully: file ({filename}) from {sender} uploaded.")
                    if delete_email_from_server(emailid, authenticate_imap(conf)):
                        debug(f"Email from {sender} deleted.")

            elif not check_sender_and_filename_V2(filename, sender):
                debug("Sender is not authorized to send the file as it is not an allowed sender.")
            
                ErrorHandler.add_error_message(f"Sender is not authorized to send the file ({filename}) provided as it is not an allowed sender."
                )
                mailbody_content = ErrorHandler.get_mailbody_content()

                debug(f"Inform sender function value before calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

                # Call inform_sender_about_errors with the updated mailbody_content
                if inform_sender_about_errors(
                    filename, sender, subject, mailbody_content, error_handler=error_handler
                ):
                    log(f"Error in files ({filename}), sent from {sender}. This email is not an allowed sender.")
                    debug(f"Error in files sent from {sender}. This email is not an allowed sender.")
                    debug(f"Inform sender function value after calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

                    # Delete the email from the server after the error message has been sent successfully
                    delete_flag = delete_email_from_server(emailid, authenticate_imap(conf))
                    if delete_flag:
                        log(f"Invalid email from {sender} deleted.")
                        sys.exit()
                    else:
                        debug(f"Email from {sender} already deleted.")
                        sys.exit()

                elif not inform_sender_about_errors(
                    filename, sender, subject, mailbody_content, error_handler=error_handler
                ):
                    log(f"Informing sender unsuccessful: Error in sending email to {sender}.")
                    debug(f"Informing sender unsuccessful: Error in sending email to {sender}.")

                    if delete_email_from_server(emailid, authenticate_imap(conf)):
                        log(f"Email from {sender} deleted.")
                    else:
                        log(f"Email from {sender} with ID {emailid} already deleted.")
                        sys.exit()

        except Exception as e:
            debug("The script raised the except statement.")
            ErrorHandler.add_error_message(
                f"It seems {sender} is not an allowed sender for {filename}. <br><br> <i><b>If you are an allowed sender:</b> there may be errors present in the filename. <br> Please try renaming the file you are trying to import in a text editor before attempting to upload again.</i>"
            )
            mailbody_content = ErrorHandler.get_mailbody_content()
            debug(f"inform_sender_about_errors() value before calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

            if inform_sender_about_errors(
                filename, sender, subject, mailbody_content, error_handler=error_handler
            ):
                debug(f"Error in files ({filename}), sent from {sender}. Function inform_sender_about_errors() was executed.")
                debug(f"Inform sender function value after calling is {inform_sender_about_errors(filename, sender, subject, mailbody_content, error_handler=error_handler)}")

                if delete_email_from_server(emailid, authenticate_imap(conf)):
                    log(f"Error present in attachments. Email from {sender} deleted.")
            else:
                log(
                    f"Error deleting email from {sender} and informing sender, with ID {emailid}."
                )
                sys.exit()

# Printing script completion
debug("The script was completed successfully.")
