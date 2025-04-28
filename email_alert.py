import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email_alert(subject, body):
    # Retrieve email credentials from environment variables
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    receiver_email = os.getenv("RECEIVER_EMAIL")

    # Check for missing environment variables
    if not sender_email or not sender_password or not receiver_email:
        raise ValueError("Missing required email environment variables!")

    try:
        # Create the MIME message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Establish the SMTP connection and send the email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Start TLS encryption
            server.login(sender_email, sender_password)
            server.send_message(msg)

        print(f"Email sent to {receiver_email}")

    except Exception as e:
        print(f"Failed to send email: {e}")

# Self-test if running as a standalone script
if __name__ == "__main__":
    subject = "Test Email"
    body = """Dear User,

This is a test email sent from the system.

Best Regards,
ICSS Team
"""
    send_email_alert(subject, body)
