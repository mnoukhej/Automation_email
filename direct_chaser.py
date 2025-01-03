import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import datetime
import csv

# Outlook SMTP server credentials
outlook_server = "smtp.office365.com"
outlook_port = 587
outlook_email = "your_outlook_email@example.com"
outlook_password = "your_outlook_password"  # Store this securely

# Path to the CSV file where sent emails are logged
sent_emails_file = 'sent_emails.csv'

# Function to send follow-up (chaser) email
def send_chaser_email(client_email, client_name):
    try:
        # Set up the SMTP server
        server = smtplib.SMTP(outlook_server, outlook_port)
        server.starttls()
        server.login(outlook_email, outlook_password)

        # Create the email content
        subject = "Reminder: Your Report"
        body = f"Dear {client_name},\n\nWe have not yet received a response regarding the report we sent 72 hours ago. Please review the report and let us know if you have any questions.\n\nBest regards,\nYour Company"

        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = outlook_email
        msg['To'] = client_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Send the email
        server.send_message(msg)
        print(f"Chaser email sent to {client_email}")
        server.quit()

    except Exception as e:
        print(f"Failed to send chaser email to {client_email}: {str(e)}")

# Read the sent emails log and send chasers if necessary
try:
    with open(sent_emails_file, mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            client_name = row[0]
            client_email = row[1]
            sent_time = datetime.datetime.strptime(row[2], "%Y-%m-%d %H:%M:%S")

            # Check if 72 hours have passed
            if (datetime.datetime.now() - sent_time).total_seconds() > 72 * 3600:
                # Send the chaser email
                send_chaser_email(client_email, client_name)
except FileNotFoundError:
    print("Error: Sent emails log file not found.")
except Exception as e:
    print(f"An error occurred: {str(e)}")
