import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load the Excel file
df = pd.read_excel('emails.xlsx')

# Email credentials (replace with your actual credentials or use environment variables)
your_email = "your_email@example.com"
your_password = "your_password"

# Set up the SMTP server (example for Gmail)
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(your_email, your_password)

for index, row in df.iterrows():
    receiver_name = row['Name']
    email = row['Email']

    # Create the email content
    msg = MIMEMultipart()
    msg['From'] = your_email
    msg['To'] = email_id
    msg['Subject'] = "SUBJECT OF THE MAIL"

    body = f"""
""MESSAGE""

    """

    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    server.send_message(msg)
    print(f"Email sent to { reciver_name} at {email_id}")

    # Clear the message after sending
    msg.clear()

# Disconnect from the server
server.quit()
