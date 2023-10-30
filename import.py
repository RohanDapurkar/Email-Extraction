import imaplib
import email
from email.header import decode_header
import os
import pandas as pd

# Email account information
EMAIL = "rohandapurkar5@gmail.com"
PASSWORD = "jryc npki qigo bxvf"
IMAP_SERVER = "imap.gmail.com"

def download_attachment(msg, download_folder):
    attachments_path = ''
    for part in msg.walk():
        if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
            continue

        filename = part.get_filename()
        if filename:
            filename = decode_header(filename)[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode()
            filepath = os.path.join(download_folder, filename)
            with open(filepath, "wb") as file:
                file.write(part.get_payload(decode=True))
        else:
            filename = ''    
        attachments_path = attachments_path+filename+'\n'       
    return attachments_path    
def process_emails():
    # Connect to the email server
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")

    # Search for unread emails
    result, email_ids = mail.search(None, "UNSEEN")
    email_id_list = email_ids[0].split()

    # Folder to save attachments
    download_folder = "attachments"
    os.makedirs(download_folder, exist_ok=True)
    subject_list = []
    sender_list = []
    time_list = []
    attachments_list = []
    for email_id in email_id_list:
        result, email_data = mail.fetch(email_id, "(RFC822)")
        raw_email = email_data[0][1]

        # Parse the email
        msg = email.message_from_bytes(raw_email)
        subject = decode_header(msg["Subject"])[0][0]
        from_ = msg.get("From")

        # Extract sender
        sender, encoding = decode_header(msg["From"])[0]
        if isinstance(sender, bytes):
            sender = sender.decode(encoding or "utf-8")

        # Extract time of receipt
        date_str = msg["Date"]

        sender_list.append(sender)
        time_list.append(date_str)
        subject_list.append(subject)

        print(f"Processing email from: {from_}, Subject: {subject}")
        print("Sender:", sender)
        print("Time of Receipt:", date_str)

        # Download attachments
        attachments_path= download_attachment(msg, download_folder)
        attachments_list.append(attachments_path)    

        # Mark the email as read
        mail.store(email_id, "+FLAGS", "(\Seen)")
    Data ={
        'subject': subject_list,
        'sender': sender_list,
        'time': time_list,
        'attachments': attachments_list
    }
  
    df = pd.DataFrame(Data)
    df.to_csv('email.csv',index= False)
    # Logout
    mail.logout()

if __name__ == "__main__":
    process_emails()
