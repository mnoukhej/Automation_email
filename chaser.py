import win32com.client as win32
import pandas as pd
import datetime
import csv
import os

# Create an instance of Outlook
outlook = win32.Dispatch("outlook.application")
namespace = outlook.GetNamespace("MAPI")

# Path to the CSV file where sent emails are logged
sent_emails_file = 'sent_emails.csv'

# Function to check if a client has replied
def has_replied(client_email):
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 refers to the Inbox folder in Outlook
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        messages = messages.Restrict(f"[SenderEmailAddress] = '{client_email}'")

        if messages.Count > 0:
            return True
        return False
    except Exception as e:
        print(f"Error checking replies for {client_email}: {str(e)}")
        return False

# Function to send the follow-up (chaser) email
def send_chaser_email(client_name, client_email):
    try:
        subject = "Reminder: Your Report"
        body = f"Dear {client_name},\n\nWe have not yet received a response regarding the report we sent 72 hours ago. Please review the report and let us know if you have any questions.\n\nBest regards,\nYour Company"

        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = client_email
        mail.Body = body
        mail.Send()
        print(f"Chaser email sent to {client_name} ({client_email})")
    
    except Exception as e:
        print(f"An error occurred while sending chaser email to {client_name} ({client_email}): {str(e)}")

# Read the sent emails log and process chaser emails
try:
    with open(sent_emails_file, mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            client_name = row[0]
            client_email = row[1]
            sent_time = datetime.datetime.strptime(row[2], "%Y-%m-%d %H:%M:%S")

            # Check if 72 hours have passed
            if (datetime.datetime.now() - sent_time).total_seconds() >300:
                # Check if the client has replied
                if not has_replied(client_email):
                    # Send the chaser email
                    send_chaser_email(client_name, client_email)
except FileNotFoundError:
    print("Error: Sent emails log file not found.")
except Exception as e:
    print(f"An error occurred while reading the log file: {str(e)}")
