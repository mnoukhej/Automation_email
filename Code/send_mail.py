import win32com.client as win32
import pandas as pd
import os
import datetime
import csv
import time

# Define the path to the Excel file and the download folder
excel_file = 'clients.xlsx'
download_folder = r"D:\Codding\Projects\Automation_email"
sent_emails_file = 'sent_emails.csv'

# Load the client data from Excel
client_data = pd.read_excel(excel_file)

# Function to start Outlook if it's not running
def start_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception:
        print("Outlook is not running. Starting Outlook...")
        os.startfile("outlook.exe")
        time.sleep(5)  # Allow time for Outlook to start
        outlook = win32.Dispatch('Outlook.Application')
        return outlook

# Start Outlook (if not running)
outlook = start_outlook()

# Subject and body of the email
subject = "Your Report"
body = "Dear {name},\n\nPlease find attached your report.\n\nBest regards,\nYour Company"

# Function to log the sent email
def log_sent_email(client_name, client_email):
    with open(sent_emails_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([client_name, client_email, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

# Iterate over each client
for index, row in client_data.iterrows():
    try:
        client_name = row['Name']
        client_email = row['Email']
        
        # Create a new email
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = client_email
        
        # Customize the body for each client
        body_text = body.format(name=client_name)
        mail.Body = body_text
        
        # Attach the report
        report_path = os.path.join(download_folder, "April_2023_Report.xlsx")  # Change file format if needed
        if os.path.exists(report_path):
            mail.Attachments.Add(report_path)
        else:
            raise FileNotFoundError(f"Report not found for {client_name} at {report_path}")
        
        # Send the email
        mail.Send()
        print(f"Email sent to {client_name} ({client_email})")
        
        # Log the sent email
        log_sent_email(client_name, client_email)
    
    except FileNotFoundError as fnf_error:
        print(f"Error: {fnf_error}")
    except Exception as e:
        print(f"An error occurred while sending email to {client_name} ({client_email}): {str(e)}")
