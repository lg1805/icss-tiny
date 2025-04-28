import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email configuration
SENDER_EMAIL = 'lakshyarubi@gmail.com'
SENDER_PASSWORD = 'selr fdih wlkm wufg'  # Replace with your app password or real password
RECEIVER_EMAIL = 'lakshyarubi.gnana2021@vitstudent.ac.in'


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
    # Email subject and body
    subject = f"Overdue Incident: {incident_id}"
    body = f"""Dear User,

Incident {incident_id} has been open for more than 3 days. Please see the details:

- Observation: {observation}
- Severity: {severity}
- Occurrence: {occurrence}
- Detection: {detection}
- RPN: {rpn}
- Priority: {priority}
- Creation Date: {creation_date}

Please act now.

Best Regards,
ICSS Team
"""

    # Construct the email message
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Send via SMTP
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {RECEIVER_EMAIL}")
    except Exception as e:
        print(f"Failed to send email to {RECEIVER_EMAIL}: {e}")
