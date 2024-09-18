# Email Attachment Downloader

**Email Attachment Downloader** is a Python application that connects to your Microsoft Outlook inbox and downloads attachments from emails with a specified subject. It automates the task of fetching attachments and saves them to a specified directory.

## Features
- Download all attachments from emails that match a specified subject.
- Overwrite or skip files if they already exist in the download folder.
- Provides feedback on the status of the download.
- Works with Microsoft Outlook via the `win32com.client` library.

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/Email-Attachment-Downloader.git
    ```

2. Navigate to the project directory:
    ```bash
    cd Email-Attachment-Downloader
    ```

3. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

4. Set up your email credentials and directory preferences in the code.

## Usage

1. Update the following variables in the code:
    - **`path`**: Set the path where you want to save the attachments.
    - **`subject`**: Define the email subject you want to search for attachments.

2. Run the Python script:
    ```bash
    python email_downloader.py
    ```

3. The script will:
    - Search your Outlook inbox for emails with the specified subject.
    - Download and save attachments to the specified directory.
    - Handle file conflicts by prompting to overwrite or skip the file.

## Example Code Snippet
Here’s a quick look at the important parts of the script:

```python
path = r'<PATH TO DIRECTORY TO SAVE>'  # Specify the directory where attachments will be saved
subject = "<SUBJECT OF EMAILS YOU WANT TO DOWNLOAD ATTACHMENTS FROM>" + today  # Specify the email subject

# Search emails with the specified subject
for message in messages:
    if subject in message.Subject:
        for attachment in message.Attachments:
            attachment.SaveAsFile(os.path.join(path, str(attachment)))  # Save the attachment
