import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Email configuration
sender_email = "ismail.prasetyo@erajaya.com"
sender_alias = "Automail"
receiver_emails = ["ismail.prasetyo@gmail.com","ismail.prasetyo@erajaya.com"]
password = "Instrument@24"
subject = "Test Email from Python attach HTML"

# Create the MIME object
message = MIMEMultipart()
message["From"] = f"{sender_alias} <{sender_email}>" #sender_email
message["To"] = ", ".join(receiver_emails)  # Join multiple recipients with commas
message["Subject"] = subject

# Add the email body
# body = "Hello, this is the body of the email!"
# HTML body with CSS styling
body = """
            <html>
            <head>
            <style>
                body {
                font-family: 'Arial', sans-serif;
                background-color: #f2f2f2;
                color: #333;
                }
                h1 {
                color: #0066cc;
                }
            </style>
            </head>
            <body>
            <h1>Hello, this is an HTML email with CSS styling!</h1>
            <p>This is a paragraph in the email body.</p>
            </body>
            </html>
            """
# message.attach(MIMEText(body, "html"))
message.attach(MIMEText(body, "html"))

# Attach multiple files
attachment_paths = ["D:/ISMAIL_ERABW/file1.txt" , "D:/ISMAIL_ERABW/file2.pdf"]
attachment_names = ["file1.txt", "file2.pdf"]

for attachment_path, attachment_name in zip(attachment_paths, attachment_names):
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {attachment_name}")
        message.attach(part)

# Connect to the SMTP server
# print("start....")

with smtplib.SMTP("smtp.erajaya.com", 587) as server:
    # Start TLS for security
    print("OK HERE")

    server.starttls()
    # Login to the email account
    server.login(sender_email, password)

    # Send the email
    server.sendmail(sender_email, receiver_emails, message.as_string())

print("Email with multiple recipients and attachments sent successfully!")
