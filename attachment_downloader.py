#############################################

# Last Modified: Friday, July 16, 2021
# Created by Devesh Aggarwal

#############################################


# ######### Importing Needed Files ######### #
import win32com.client  # allow python to microsoft communications (ie. outlook)
import os  # used for path management (ie. getting, saving, and
# joining directories)
from sys import exit  # imported to end program forcefully
from datetime import datetime  # get today's date

print("STATUS: Starting...")  # display msg

try:
    # ######### Initialization ######### #
    print("STATUS: Initializing...")  # display msg
    today = datetime.today().strftime('%Y-%m-%d')  # gets today's date
    # (format = yyyy-mm-dd)
    path = r'<PATH TO DIRECTORY TO SAVE>'  # path to save
    # attachments
    subject = "<SUBJECT OF EMAILS YOU WANT " \
              "TO DOWNLOAD ATTACHMENTS FROM>" + today  # assigning the expected
    # subject as a string
    found_msg = False  # initialization for validation purposes
    files_in_path = [f for f in os.listdir(path)]  # getting list of files
    # in path specified above

    # ######### Object Creation ######### #
    outlook = win32com.client.Dispatch("Outlook.Application"). \
        GetNamespace("MAPI")  # Creating an object for the outlook application.
    inbox = outlook.GetDefaultFolder(6)  # Creating an object to access
    # Inbox of the outlook.
    messages = inbox.Items  # Creating an object to access
    # items inside the inbox of outlook.

    print("STATUS: Searching for specified email...")  # display msg

    # ######### Main Logic ######### #
    for message in messages:
        if subject in message.Subject:

            print("STATUS: Found email!")
            found_msg = True

            for attachment in message.Attachments:
                if str(attachment) in files_in_path:
                    # asks user for overriding
                    u_response = input(
                        "ERROR: File name already exists in specified "
                        "directory. Would you like to override it? [y/n]")
                    # saves attachment / overrides original
                    if u_response.lower() == "y" or u_response.lower() == "yes":
                        attachment.SaveAsFile(os.path.join(path,
                                                           str(attachment)))
                    else:
                        input("ATTACHMENTS NOT SAVED (TERMINATED)  "
                              "[enter to quit]")  # display msg
                        exit(1)  # kills the program w/ exit status of "error"
                else:
                    attachment.SaveAsFile(os.path.join(path, str(attachment)))
                    # saves attachment in specified dir

            message.Unread = False
            input(f"STATUS: Successfully saved all attachments found in email "
                  f"at: {path} [enter to exit]")  # display msg
            exit(0)  # kills program w/ exit status of "success"

    if not found_msg:  # if msg doesn't exist
        input("ERROR: Couldn't find specified email.  [enter to quit]")
except FileNotFoundError:
    input("ERROR: The specified directory is invalid.  [enter to quit]")
except:
    input("ERROR: AN UNKNOWN ERROR OCCURRED")
