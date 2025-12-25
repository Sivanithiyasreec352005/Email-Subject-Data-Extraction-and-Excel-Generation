import imaplib
import email
from email.header import decode_header
import pdfplumber
import pandas as pd
import os
import time
import smtplib #to send emails via STMP
from email.message import EmailMessage

# ️Gmail Login (IMAP)
gmail_user = "csivanithiyasree@gmail.com"
app_password = "isoz frds eddc mvoh"  # App Password
imap_host = "imap.gmail.com"

# Connect to Gmail via IMAP
mail = imaplib.IMAP4_SSL(imap_host)# Connects securely to Gmail’s IMAP server.
mail.login(gmail_user, app_password)
mail.select("inbox")  # Select inbox folder  #Chooses the Inbox folder to read emails from.


# Search emails
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Only consider the first 10 emails (newest first)
email_ids = email_ids[-10:]
matching_emails = []

# Filter emails whose subject contains "customer details"

for email_id in reversed(email_ids):  # newest first among these 10
    status, msg_data = mail.fetch(email_id, "(RFC822)")#RFC 822 is a standard that defines the format of an email message.
#It includes headers (like From, To, Subject, Date) and the body (text, HTML, or attachments).
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    
    subject = str(email.header.make_header(email.header.decode_header(msg.get("Subject"))))
    
    if "customer details" in subject.lower():
        matching_emails.append(email_id)

if not matching_emails:
    print("❌ No matching emails found among the first 10 emails.")
    exit()

print(f"✅ Found {len(matching_emails)} matching email(s) among the first 10 emails.")

# Process each matching email

excel_files = []  # Store all generated Excel files

for email_id in matching_emails:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    
    sender = msg.get("From")
    print(f"\nProcessing email from: {sender}")
    
    # Download PDF attachment

    pdf_path = None
    for part in msg.walk(): # Loops through all parts of the email
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue

        filename = part.get_filename()
        #Check if the part is a PDF and then Decode the filename
        if filename and filename.lower().endswith(".pdf"):
            filename = decode_header(filename)[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode()
                #Create the path to save the PDF
            pdf_path = os.path.join(os.getcwd(), filename)
           # This line actually writes the file to your computer

            with open(pdf_path, "wb") as f:
                f.write(part.get_payload(decode=True))
            print(f"✅ PDF attachment saved: {filename}")
            break

    if pdf_path is None:
        print("skipping...")
        continue

    # -----------------------------
    # Convert PDF to Excel
    # -----------------------------
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                rows.extend(table)

    if not rows:
        print("❌ No table data found in PDF, skipping...")
        continue

    # Save as Excel
    df = pd.DataFrame(rows[1:], columns=rows[0])
    excel_file = f"CustomerDetails_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(excel_file, index=False)
    excel_files.append(excel_file)
    print(f"✅ PDF converted to Excel: {excel_file}")

# -----------------------------
# Send Excel files via email
# -----------------------------
if excel_files:
    msg = EmailMessage()
    msg['Subject'] = 'Customer Details Excel Files'
    msg['From'] = gmail_user
    msg['To'] = gmail_user  # Sending to self
    msg.set_content('Please find attached the Excel file(s) generated from customer details PDFs.')

    # Attach all Excel files
    for file in excel_files:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(file)
            msg.add_attachment(file_data, maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=file_name)

    # Send email via SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(gmail_user, app_password)
        smtp.send_message(msg)

    print(f"\n✅ Email sent successfully with {len(excel_files)} attachment(s) to {gmail_user}")
else:
    print("\n⚠️ No Excel files generated. Email not sent.")
    # Save as Excel
    df = pd.DataFrame(rows[1:], columns=rows[0])
    excel_file = f"CustomerDetails_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(excel_file, index=False)
    print(f"✅ PDF converted to Excel: {excel_file}")

print("\n✅ All matching emails processed successfully!")
