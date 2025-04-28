import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load SMTP credentials from environment variables
SENDER_EMAIL = os.getenv('SENDER_EMAIL')  # e.g. 'lakshyarubi@gmail.com'
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')  # your app password
RECEIVER_EMAIL = os.getenv('RECEIVER_EMAIL')  # 'lakshyarubi.gnana2021@vitstudent.ac.in'


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
    # Verify that credentials are set
    if not (SENDER_EMAIL and SENDER_PASSWORD and RECEIVER_EMAIL):
        raise EnvironmentError("EMAIL_ALERT: Missing SENDER_EMAIL, SENDER_PASSWORD, or RECEIVER_EMAIL environment variable")

    # Build email subject and body
    subject = f"Overdue Incident: {incident_id}"
    body = (
        f"Dear User,\n\n"
        f"Incident {incident_id} has been open for more than 3 days. Details:\n\n"
        f"- Observation: {observation}\n"
        f"- Severity: {severity}\n"
        f"- Occurrence: {occurrence}\n"
        f"- Detection: {detection}\n"
        f"- RPN: {rpn}\n"
        f"- Priority: {priority}\n"
        f"- Creation Date: {creation_date}\n\n"
        f"Please act now.\n\n"
        f"Best Regards,\n"
        f"ICSS Team"
    )

    # Construct MIME message
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Send the email via SMTP
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {RECEIVER_EMAIL}")
    except Exception as e:
        print(f"Failed to send email: {e}")

