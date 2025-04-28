import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email_alert(incident_id, observation, severity, occurrence,
                     detection, rpn, priority, creation_date):
    """
    Sends an email alert for an overdue incident.

    Parameters:
    - incident_id: str or int
    - observation: str
    - severity: int
    - occurrence: int
    - detection: int
    - rpn: int
    - priority: str
    - creation_date: str or datetime
    """
    # Load SMTP credentials from environment variables
    sender_email = os.getenv('SENDER_EMAIL')
    sender_password = os.getenv('SENDER_PASSWORD')
    receiver_email = os.getenv('RECEIVER_EMAIL')

    if not sender_email or not sender_password or not receiver_email:
        print("EMAIL ALERT CONFIGURATION ERROR: Missing environment variables.")
        return

    # Prepare email content
    subject = f"Overdue Incident: {incident_id}"
    body = f"""Dear User,\n\n" \
           f"Incident {incident_id} has been open for more than 3 days. Please see the details below:\n\n" \
           f"- Observation: {observation}\n" \
           f"- Severity: {severity}\n" \
           f"- Occurrence: {occurrence}\n" \
           f"- Detection: {detection}\n" \
           f"- RPN: {rpn}\n" \
           f"- Priority: {priority}\n" \
           f"- Creation Date: {creation_date}\n\n" \
           f"Please act now.\n\n" \
           f"Best Regards,\n" \
           f"ICSS Team"

    # Construct email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Send via SMTP
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")


if __name__ == '__main__':
    # Simple test when running directly
    send_email_alert(
        incident_id='TEST123',
        observation='Test observation',
        severity=1,
        occurrence=1,
        detection=1,
        rpn=1,
        priority='Low',
        creation_date='2025-04-28'
    )
