import os
import smtplib
import ssl

# SMTP configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Load environment variables
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")


def test_smtp():
    """
    Sends a simple test email using SMTP with debug logging.
    """
    if not (SENDER_EMAIL and SENDER_PASSWORD and RECEIVER_EMAIL):
        print("ERROR: Set SENDER_EMAIL, SENDER_PASSWORD, and RECEIVER_EMAIL environment vars.")
        return

    message = "Subject: SMTP Test\n\nThis is a test email from email_test.py"
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.set_debuglevel(1)  # Enable SMTP debug output
            server.starttls(context=context)
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, message)
        print(f"Test email sent successfully to {RECEIVER_EMAIL}")
    except Exception as e:
        print("SMTP test failed:", e)


if __name__ == "__main__":
    test_smtp()

