import win32com.client
import re
import pandas as pd

# Connect to Outlook
Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox folder
inbox = Outlook.GetDefaultFolder(6)  # "6" refers to the inbox - note that subfolders are not included
messages = inbox.Items

# Define the regular expression pattern for matching email addresses
email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Create a set to store unique email addresses
unique_email_addresses = set()

# Iterate through the messages in the inbox
for message in messages:
    try:
        # Extract the specified fields
        sender = message.Sender.Address
        recipients = [recipient.Address for recipient in message.Recipients]
        subject = message.Subject
        body = message.Body

        # Find all email addresses in the fields using the pattern
        sender_emails = re.findall(email_pattern, sender)
        recipient_emails = re.findall(email_pattern, ';'.join(recipients))
        subject_emails = re.findall(email_pattern, subject)
        body_emails = re.findall(email_pattern, body)

        # Add the email addresses to the set
        unique_email_addresses.update(sender_emails)
        unique_email_addresses.update(recipient_emails)
        unique_email_addresses.update(subject_emails)
        unique_email_addresses.update(body_emails)
    except Exception as e:
        print(f"Error occurred while processing a message: {str(e)}")

# Create a DataFrame to store the unique email addresses
df_emails = pd.DataFrame({'Email Addresses': list(unique_email_addresses)})

# Save the DataFrame to a CSV file
df_emails.to_csv('unique_email_addresses.csv', index=False)
