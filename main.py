import smtplib
import ssl
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from gooey import Gooey, GooeyParser
import sys

def reading_excel_file(filename):
    if filename == "" or filename is None:
        print("No file selected.")
        sys.stdout.flush()
        return

    # Read the entire Excel file
    df = pd.read_excel(filename)
    email_list = df['Email'].tolist()
    subject_list = df['Subject'].tolist()
    content_list = df['Content'].tolist()
    
    for email, subject, content in zip(email_list, subject_list, content_list):
        sending_mail(email, subject, content)

def sending_mail(receiver_email, subject, message_body):
    sender_email = "massabisgood@gmail.com"
    password = "hyfq znms beji abpf"

    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email

    # Attach message body
    message.attach(MIMEText(message_body, "plain"))

    # Create secure connection with server and send email
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            print(f"Email sent successfully to {receiver_email}")
            sys.stdout.flush()
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.stdout.flush()



@Gooey(program_description="Automated Email Sender",program_name="Email Sender")
def main():
    parser = GooeyParser()

    parser.add_argument("Select_File", widget="FileChooser",
                        help="Select Excel File That Has Email, Subject, Email Content in it",
                        action='store',  # Changed action to 'store'
                        gooey_options=dict(wildcard="Excel (.xlsx) | *.xlsx"))

    args = parser.parse_args()
    file = args.Select_File

    reading_excel_file(file)

if __name__ == "__main__":
    main()

