from flask import Flask, request, render_template, send_file, abort
import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from rapidfuzz import fuzz
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)

# Where processed uploads/downloads live
dir_base     = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(dir_base, 'uploads', 'processed')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# RPN lookup file
RPN_FILE = os.path.join(dir_base, 'ProcessedData', 'RPN.xlsx')
if not os.path.exists(RPN_FILE):
    raise FileNotFoundError(f"RPN file not found at {RPN_FILE}")

# Load RPN data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data['Component'].dropna().unique().tolist()
executor = ThreadPoolExecutor(max_workers=4)

# Fixed receiver address
RECEIVER_EMAIL = 'lakshyarubi.gnana2021@vitstudent.ac.in'


def extract_component(obs):
    obs_str = str(obs).strip()
    best, best_score = None, 0
    for comp in known_components:
        score = fuzz.partial_ratio(comp.lower(), obs_str.lower())
        if score >= 80 and score > best_score:
            best, best_score = comp, score
    return best or 'Unknown'


def get_rpn_values(component):
    row = rpn_data[rpn_data['Component'] == component]
    if row.empty:
        return 1, 1, 10
    s, o, d = row.iloc[0]['Severity (S)'], row.iloc[0]['Occurrence (O)'], row.iloc[0]['Detection (D)']
    try:
        return int(s), int(o), int(d)
    except:
        return 1, 1, 10


def determine_priority(rpn):
    return 'High' if rpn >= 200 else 'Moderate' if rpn >= 100 else 'Low'


def send_email(to_email, subject, body):
    sender = 'lakshyarubi@gmail.com'
    password = 'selr fdih wlkm wufg'  # your app-specific Gmail password
    try:
        msg = MIMEMultipart()
        msg['From']    = sender
        msg['To']      = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.set_debuglevel(1)  # SMTP debug
            server.starttls()
            server.login(sender, password)
            server.send_message(msg)
        app.logger.info(f"Email sent to {to_email}")
    except Exception as e:
        app.logger.error(f"Failed to send email to {to_email}: {e}")


@app.route('/')
def index():
    return render_template('frontNEW.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        abort(400, 'No complaint_file part')
    file = request.files['complaint_file']
    if not file or not file.filename:
        abort(400, 'No selected file')

    # Save upload
    in_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(in_path)
    app.logger.info(f"Saved upload to {in_path}")

    # Read data
    try:
        df = pd.read_excel(in_path)
    except Exception as e:
        abort(400, f"Error reading Excel: {e}")

    # Required columns (no Email column)
    required = ['Observation', 'Creation Date', 'Incident Id', 'Incident Status']
    missing = [c for c in required if c not in df.columns]
    if missing:
        abort(400, f"Missing columns: {', '.join(missing)}")

    # Parse dates & compute days elapsed
    df['Creation Date'] = pd.to_datetime(df['Creation Date'], errors='coerce')
    df['Days Elapsed']   = (datetime.now() - df['Creation Date']).dt.days

    # Extract components & RPN calculations
    df['Component'] = list(executor.map(extract_component, df['Observation']))
    rpn_vals = list(executor.map(get_rpn_values, df['Component']))
    df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = \
        pd.DataFrame(rpn_vals, index=df.index)
    df['RPN']      = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']
    df['Priority'] = df['RPN'].apply(determine_priority)

    # Write processed Excel
    out_name = f"processed_{file.filename}"
    out_path = os.path.join(UPLOAD_FOLDER, out_name)
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    app.logger.info(f"Written processed file to {out_path}")

    # Send email to fixed receiver for each overdue incident
    overdue = df[(df['Incident Status'].str.lower() == 'open') & (df['Days Elapsed'] > 3)]
    app.logger.info(f"Overdue count: {len(overdue)}")
    for _, row in overdue.iterrows():
        subject = f"Incident {row['Incident Id']} – Action Required"
        body = (
            f"Dear User,\n\n"
            f"Incident {row['Incident Id']} has been open for {row['Days Elapsed']} days.\n"
            f"Details:\n"
            f"Observation: {row['Observation']}\n"
            f"Severity: {row['Severity (S)']}\n"
            f"Occurrence: {row['Occurrence (O)']}\n"
            f"Detection: {row['Detection (D)']}\n"
            f"RPN: {row['RPN']}\n"
            f"Priority: {row['Priority']}\n"
            f"Created: {row['Creation Date']}\n\n"
            f"Please address this promptly.\n"
            f"ICSS Team"
        )
        send_email(RECEIVER_EMAIL, subject, body)

    # Return the processed file for download
    return send_file(out_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
