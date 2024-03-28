import smtplib
import mysql.connector
import argparse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Create ArgumentParser object
parser = argparse.ArgumentParser(description='A simple script with command-line arguments.')

# Add command-line arguments
parser.add_argument('--report_id', type=str, help='Report ID')
parser.add_argument('--email_id', type=str, help='Email Sender ID')
parser.add_argument('--attachment_paths', type=str, help='Attachment Paths')
parser.add_argument('--attachment_names', type=str, help='Attachment Names')

# Parse the command-line arguments
args = parser.parse_args()

email_id = args.email_id
report_id = args.report_id
attachment_paths = args.attachment_paths
attachment_names = args.attachment_names

# Replace these with your MySQL database connection details
host = "localhost"
user = "ismail"
password = "Wellings@24"
database = "db_ip"

# Establish a connection to the MySQL database
conn = mysql.connector.connect(
    host=host,
    user=user,
    password=password,
    database=database
)

# Create a cursor to interact with the database
cursor = conn.cursor()
try:
    # Example SELECT query
    query = f"select sender, sender_email, password, smtp, smtp_port from PARAM_AUTOMAIL where account_no = {email_id}"
    # Execute the query
    cursor.execute(query)
    # Fetch all rows
    rows = cursor.fetchall()
    # Print the result
    for row in rows:
        # print(row[1])
        sender_email = row[1]
        sender_alias = row[0]
        password = row[2]
        smtp = row[2]
        smtp_port = row[3]

except mysql.connector.Error as err:
    print("Error: {}".format(err))

finally:
    # Close the cursor and connection
    cursor.close()

# second query
cursor = conn.cursor()

try:
    query = f"SELECT report_name, subject, GROUP_CONCAT(email_to order by email_to SEPARATOR ', ') AS email_to, GROUP_CONCAT(email_cc order by email_cc SEPARATOR ', ') AS email_cc, GROUP_CONCAT(email_bc order by email_bc SEPARATOR ', ') AS email_bc  FROM REF_AUTOMAIL  where REPORT_ID = {report_id} and status = 1  group by report_name, subject" 
    # print(query)
    cursor.execute(query)
    rows = cursor.fetchall()
    for row in rows:
        email_subject = row[1]
        email_to = row[2]
        email_cc = row[3]
        email_bc = row[4]
    
except mysql.connector.Error as err:
    print("Error: {}".format(err))
finally:
    # Close the cursor and connection
    cursor.close()

# third query 
cursor = conn.cursor()
try:
    query = f"SELECT body FROM db_ip.ref_automail_body WHERE report_id = {report_id}" 
    cursor.execute(query)
    rows = cursor.fetchall()
    for row in rows:
        email_body = row[0]

    
except mysql.connector.Error as err:
    print("Error: {}".format(err))
finally:
    # Close the cursor and connection
    cursor.close()    
conn.close()

receiver_emails = email_to.split(',') if email_to is not None else [] 
receiver_emails += email_cc.split(',') if email_cc is not None else []
receiver_emails += email_bc.split(',') if email_bc is not None else []

# Create the MIME object
message = MIMEMultipart()
message["From"] = f"{sender_alias} <{sender_email}>" #sender_email
message["To"] =  email_to #", ".join(receiver_emails)  # Join multiple recipients with commas
message["CC"] = email_cc
message["BCC"] = email_bc
message["Subject"] = email_subject

# Add the email body
# body = "Hello, this is the body of the email!"
# HTML body with CSS styling
body = """
            <html>
            <head>
            <style>
                body {{
                font-family: 'Arial', sans-serif;
                background-color: #f2f2f2;
                color: #333;
                }}
                h1 {{
                color: #0066cc;
                }}
            </style>
            </head>
            <body>{}
            </body>
            </html>
            """.format(email_body)
# print (body)
message.attach(MIMEText(body, "html"))

# D:/ISMAIL_ERABW/filetest3.csv,D:/ISMAIL_ERABW/file1.txt
# filetest3.csv,file1.txt
# py automail_final.py --report_id 0 --email_id 1 --attachment_paths D:/ISMAIL_ERABW/filetest3.csv,D:/ISMAIL_ERABW/file1.txt --attachment_names filetest3.csv,file1.txt
# Attach multiple files
attachment_paths = args.attachment_paths.split(',')  #["D:/ISMAIL_ERABW/file1.txt" , "D:/ISMAIL_ERABW/file2.pdf"]
attachment_names = args.attachment_names.split(',')  #["file1.txt", "file2.pdf"]

for attachment_path, attachment_name in zip(attachment_paths, attachment_names):
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {attachment_name}")
        message.attach(part)

# Connect to the SMTP server
# print("start....")
# print(receiver_emails)
with smtplib.SMTP("smtp.erajaya.com", 587) as server:
    # Start TLS for security
    # print("OK HERE")

    server.starttls()
    # Login to the email account
    server.login(sender_email, password)

    # Send the email
    server.sendmail(sender_email, receiver_emails, message.as_string())

print("Email with multiple recipients and attachments sent successfully!")
