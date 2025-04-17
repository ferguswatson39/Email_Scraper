import imaplib
import email
import re
import os
from bs4 import BeautifulSoup
from dotenv import load_dotenv

imap_url = 'imap.gmail.com'

load_dotenv(r'C:\Users\Fergus Watson\OneDrive\Desktop\Python Web\Scraping\V1.0 Email Scraping\the_keys.env')
user = os.getenv('user')
password = os.getenv('password')

try:
    my_mail = imaplib.IMAP4_SSL(imap_url)
    my_mail.login(user, password)
    print("Login Successful!")
except imaplib.IMAP4.error:
    print("Login failed. Check credentials.")
    exit()


my_mail.select('Inbox')

# Search for emails from a specific sender
# Sent email content to myself as a draft hence why user is used as email sender - change to relevant email sender
status, email_ids = my_mail.search(None, 'FROM', user)

if status == "OK" and email_ids:
    email_ids = email_ids[0].split()  # Split into individual email IDs

    if email_ids:
        print("Data Retrieved")

        for email_id in email_ids:
            status, msg_data = my_mail.fetch(email_id, '(RFC822)')

            if status == "OK":
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        raw_email = response_part[1]
                        mail_body = email.message_from_bytes(raw_email)

                        email_text = ""

                        if mail_body.is_multipart():
                            for part in mail_body.walk():
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content-Disposition"))

                                if content_type == "text/plain" and "attachment" not in content_disposition:
                                    email_text += part.get_payload(decode=True).decode() + "\n"

                                elif content_type == "text/html" and "attachment" not in content_disposition:
                                    html_body = part.get_payload(decode=True).decode()
                                    soup = BeautifulSoup(html_body, "html.parser")
                                    email_text += soup.get_text() + "\n"
                        else:
                            email_text = mail_body.get_payload(decode=True).decode()

                        # Debugging - Print email text to see structure
                        print("\n=== Email Content Start ===\n")
                        print(email_text)
                        print("\n=== Email Content End ===\n")

                        # Regex for Desk Number and Date
                        desk_number_match = re.search(r'Desk (\d+)', email_text)
                        desk_date_match = re.search(r'When\s+([A-Za-z]+, \d{1,2} \w+ \d{4})', email_text)

                        desk_number = desk_number_match.group(1) if desk_number_match else "Not Found"
                        desk_date = desk_date_match.group(1) if desk_date_match else "Not Found"

                        print(f"Extracted Data from Email ID {email_id}:")
                        print(f"Desk Number: {desk_number}")
                        print(f"Desk Date: {desk_date}")
                        print("-" * 40)

            else:
                print(f"Failed to fetch email with ID: {email_id}")
    else:
        print("No emails found.")
else:
    print("No emails found.")

# Close and logout
my_mail.close()
my_mail.logout()
print("Logout Successful")
