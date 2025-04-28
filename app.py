
from flask import Flask, request, render_template, send_file, abort
import pandas as pd
import os
from datetime import datetime
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from rapidfuzz import fuzz
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads', 'processed')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

RPN_FILE = os.path.join(os.path.dirname(__file__), 'ProcessedData', 'RPN.xlsx')
if not os.path.exists(RPN_FILE):
    raise FileNotFoundError(f"RPN file not found at {RPN_FILE}")

# Load the RPN data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()
executor = ThreadPoolExecutor(max_workers=4)


def extract_component(obs):
    obs_str = str(obs).strip()
    best_match, highest_score = None, 0
    for comp in known_components:
        score = fuzz.partial_ratio(comp.lower(), obs_str.lower())
        if score >= 80 and score > highest_score:
            best_match, highest_score = comp, score
    return best_match or "Unknown"


def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        s, o, d = row.iloc[0]["Severity (S)"], row.iloc[0]["Occurrence (O)"], row.iloc[0]["Detection (D)"]
        # Type-check and convert
        try:
            s = float(s) if not isinstance(s, (int, float)) else s
            o = float(o) if not isinstance(o, (int, float)) else o
            d = float(d) if not isinstance(d, (int, float)) else d
        except Exception:
            app.logger.warning(f"Invalid RPN values for component '{component}', using defaults")
            return 1, 1, 10
        return int(s), int(o), int(d)
    return 1, 1, 10


def determine_priority(rpn):
    return "High" if rpn >= 200 else "Moderate" if rpn >= 100 else "Low"


def month_str_to_num(month_hint):
    months = {m[:3].lower(): f"{i:02d}" for i, m in enumerate(
        ["January","February","March","April","May","June",
         "July","August","September","October","November","December"], 1)}
    return months.get(month_hint[:3].lower())


def format_creation_date(date_str, month_hint):
    target_m = month_str_to_num(month_hint)
    try:
        dt = pd.to_datetime(str(date_str).strip(), errors='coerce', dayfirst=True)
        if pd.isna(dt):
            return None, None
        dd, mm, yy = dt.day, dt.month, dt.year
        if dd == 1 and mm == 1 and target_m:
            mm = int(target_m)
        formatted = f"{dd:02d}/{mm:02d}/{yy}"
        return formatted, (datetime.now() - dt).days
    except Exception as e:
        app.logger.error(f"Error parsing date '{date_str}': {e}")
        return None, None


def send_email(to_email, subject, body):
    sender_email = 'lakshyarubi@gmail.com'
    sender_password = 'selr fdih wlkm wufg'
    try:
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = sender_email, to_email, subject
        msg.attach(MIMEText(body, 'plain'))
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        app.logger.info(f"Email sent to {to_email}")
    except Exception as e:
        app.logger.error(f"Failed to send email to {to_email}: {e}")


@app.route('/')
def index():
    return render_template('frontNEW.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    # Basic upload checks
    if 'complaint_file' not in request.files:
        abort(400, 'No complaint_file part')
    file = request.files['complaint_file']
    if not file or file.filename.strip() == '':
        abort(400, 'No selected file')

    month_hint = request.form.get('month_hint', '')
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    app.logger.info(f"Saved upload to {filepath}")

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        app.logger.error(f"Error reading Excel: {e}")
        abort(400, f"Error reading file: {e}")

    # Required columns
    needed = ['Observation', 'Creation Date', 'Incident Id', 'Incident Status']
    missing = [c for c in needed if c not in df.columns]
    if missing:
        abort(400, f"Missing columns: {', '.join(missing)}")

    # Debug head
    app.logger.debug(f"DataFrame head:\n{df.head()}")

    # Date formatting
    formatted = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
    df['Creation Date'] = formatted.map(lambda x: x[0])
    days_elapsed = formatted.map(lambda x: x[1]).fillna(0).astype(int)
    app.logger.debug(f"Days elapsed sample: {days_elapsed.head()}")

    # Component extraction and RPN parallel
    df['Component'] = list(executor.map(extract_component, df['Observation']))
    rpn_vals = list(executor.map(get_rpn_values, df['Component']))
    df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = pd.DataFrame(rpn_vals, index=df.index)
    df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']
    df['Priority'] = df['RPN'].apply(determine_priority)

    # Segregate
    mask_spn = df['Observation'].str.contains('spn', case=False, na=False)
    spn_df, non_spn_df = df[mask_spn].copy(), df[~mask_spn].copy()
    order = {'High':1, 'Moderate':2, 'Low':3}
    spn_df.sort_values('Priority', key=lambda x: x.map(order), inplace=True)
    non_spn_df.sort_values('Priority', key=lambda x: x.map(order), inplace=True)

    # Write Excel
    processed = os.path.join(UPLOAD_FOLDER, f"processed_{file.filename}")
    with pd.ExcelWriter(processed, engine='xlsxwriter') as writer:
        for name, sheet in (('SPN', spn_df), ('Non-SPN', non_spn_df)):
            sheet.to_excel(writer, sheet_name=name, index=False)
            wb, ws = writer.book, writer.sheets[name]
            green = wb.add_format({'bg_color':'#C6EFCE'})
            for i, idx in enumerate(sheet.index, start=1):
                status = str(sheet.at[idx, 'Incident Status']).lower()
                col_id = sheet.columns.get_loc('Incident Id')
                col_st = sheet.columns.get_loc('Incident Status')
                if 'closed' in status or 'complete' in status:
                    ws.write(i, col_st, sheet.at[idx, 'Incident Status'], green)
                else:
                    color = ['#ADD8E6','#FFFF00','#FF1493','#FF0000'][min(days_elapsed.at[idx], 4)-1] if days_elapsed.at[idx] > 0 else None
                    if color:
                        ws.write(i, col_id, sheet.at[idx, 'Incident Id'], wb.add_format({'bg_color': color}))
    app.logger.info(f"Written processed file to {processed}")

    # Verify file exists before sending
    if not os.path.exists(processed):
        app.logger.error(f"Processed file not found: {processed}")
        abort(500, "Processed file missing")

    # Email alerts
    if 'Email' in df.columns:
        overdue = df[(df['Incident Status'].str.lower()=='open') & (days_elapsed>3)]
        for _, row in overdue.iterrows():
            subj = f"Overdue Incident: {row['Incident Id']}"
            body = (f"Dear User,\nIncident {row['Incident Id']} open >3 days. Details:\n"
                    f"Obs: {row['Observation']}\nSeverity: {row['Severity (S)']}\n"
                    f"Occurrence: {row['Occurrence (O)']}\nDetection: {row['Detection (D)']}\n"
                    f"RPN: {row['RPN']}\nPriority: {row['Priority']}\n"
                    f"Created: {row['Creation Date']}\n\nPlease act now.\nICSS Team")
            send_email(row['Email'], subj, body)

    return send_file(processed, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)


from flask import Flask, request, render_template, send_file, abort
import pandas as pd
import os
from datetime import datetime
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from rapidfuzz import fuzz
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads', 'processed')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

RPN_FILE = os.path.join(os.path.dirname(__file__), 'ProcessedData', 'RPN.xlsx')
if not os.path.exists(RPN_FILE):
    raise FileNotFoundError(f"RPN file not found at {RPN_FILE}")

# Load the RPN data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()
executor = ThreadPoolExecutor(max_workers=4)


def extract_component(obs):
    obs_str = str(obs).strip()
    best_match, highest_score = None, 0
    for comp in known_components:
        score = fuzz.partial_ratio(comp.lower(), obs_str.lower())
        if score >= 80 and score > highest_score:
            best_match, highest_score = comp, score
    return best_match or "Unknown"


def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        s, o, d = row.iloc[0]["Severity (S)"], row.iloc[0]["Occurrence (O)"], row.iloc[0]["Detection (D)"]
        # Type-check and convert
        try:
            s = float(s) if not isinstance(s, (int, float)) else s
            o = float(o) if not isinstance(o, (int, float)) else o
            d = float(d) if not isinstance(d, (int, float)) else d
        except Exception:
            app.logger.warning(f"Invalid RPN values for component '{component}', using defaults")
            return 1, 1, 10
        return int(s), int(o), int(d)
    return 1, 1, 10


def determine_priority(rpn):
    return "High" if rpn >= 200 else "Moderate" if rpn >= 100 else "Low"


def month_str_to_num(month_hint):
    months = {m[:3].lower(): f"{i:02d}" for i, m in enumerate(
        ["January","February","March","April","May","June",
         "July","August","September","October","November","December"], 1)}
    return months.get(month_hint[:3].lower())


def format_creation_date(date_str, month_hint):
    target_m = month_str_to_num(month_hint)
    try:
        dt = pd.to_datetime(str(date_str).strip(), errors='coerce', dayfirst=True)
        if pd.isna(dt):
            return None, None
        dd, mm, yy = dt.day, dt.month, dt.year
        if dd == 1 and mm == 1 and target_m:
            mm = int(target_m)
        formatted = f"{dd:02d}/{mm:02d}/{yy}"
        return formatted, (datetime.now() - dt).days
    except Exception as e:
        app.logger.error(f"Error parsing date '{date_str}': {e}")
        return None, None


def send_email(to_email, subject, body):
    sender_email = 'lakshyarubi@gmail.com'
    sender_password = 'selr fdih wlkm wufg'
    try:
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = sender_email, to_email, subject
        msg.attach(MIMEText(body, 'plain'))
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        app.logger.info(f"Email sent to {to_email}")
    except Exception as e:
        app.logger.error(f"Failed to send email to {to_email}: {e}")


@app.route('/')
def index():
    return render_template('frontNEW.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    # Basic upload checks
    if 'complaint_file' not in request.files:
        abort(400, 'No complaint_file part')
    file = request.files['complaint_file']
    if not file or file.filename.strip() == '':
        abort(400, 'No selected file')

    month_hint = request.form.get('month_hint', '')
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    app.logger.info(f"Saved upload to {filepath}")

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        app.logger.error(f"Error reading Excel: {e}")
        abort(400, f"Error reading file: {e}")

    # Required columns
    needed = ['Observation', 'Creation Date', 'Incident Id', 'Incident Status']
    missing = [c for c in needed if c not in df.columns]
    if missing:
        abort(400, f"Missing columns: {', '.join(missing)}")

    # Debug head
    app.logger.debug(f"DataFrame head:\n{df.head()}")

    # Date formatting
    formatted = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
    df['Creation Date'] = formatted.map(lambda x: x[0])
    days_elapsed = formatted.map(lambda x: x[1]).fillna(0).astype(int)
    app.logger.debug(f"Days elapsed sample: {days_elapsed.head()}")

    # Component extraction and RPN parallel
    df['Component'] = list(executor.map(extract_component, df['Observation']))
    rpn_vals = list(executor.map(get_rpn_values, df['Component']))
    df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = pd.DataFrame(rpn_vals, index=df.index)
    df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']
    df['Priority'] = df['RPN'].apply(determine_priority)

    # Segregate
    mask_spn = df['Observation'].str.contains('spn', case=False, na=False)
    spn_df, non_spn_df = df[mask_spn].copy(), df[~mask_spn].copy()
    order = {'High':1, 'Moderate':2, 'Low':3}
    spn_df.sort_values('Priority', key=lambda x: x.map(order), inplace=True)
    non_spn_df.sort_values('Priority', key=lambda x: x.map(order), inplace=True)

    # Write Excel
    processed = os.path.join(UPLOAD_FOLDER, f"processed_{file.filename}")
    with pd.ExcelWriter(processed, engine='xlsxwriter') as writer:
        for name, sheet in (('SPN', spn_df), ('Non-SPN', non_spn_df)):
            sheet.to_excel(writer, sheet_name=name, index=False)
            wb, ws = writer.book, writer.sheets[name]
            green = wb.add_format({'bg_color':'#C6EFCE'})
            for i, idx in enumerate(sheet.index, start=1):
                status = str(sheet.at[idx, 'Incident Status']).lower()
                col_id = sheet.columns.get_loc('Incident Id')
                col_st = sheet.columns.get_loc('Incident Status')
                if 'closed' in status or 'complete' in status:
                    ws.write(i, col_st, sheet.at[idx, 'Incident Status'], green)
                else:
                    color = ['#ADD8E6','#FFFF00','#FF1493','#FF0000'][min(days_elapsed.at[idx], 4)-1] if days_elapsed.at[idx] > 0 else None
                    if color:
                        ws.write(i, col_id, sheet.at[idx, 'Incident Id'], wb.add_format({'bg_color': color}))
    app.logger.info(f"Written processed file to {processed}")

    # Verify file exists before sending
    if not os.path.exists(processed):
        app.logger.error(f"Processed file not found: {processed}")
        abort(500, "Processed file missing")

    # Email alerts
    if 'Email' in df.columns:
        overdue = df[(df['Incident Status'].str.lower()=='open') & (days_elapsed>3)]
        for _, row in overdue.iterrows():
            subj = f"Overdue Incident: {row['Incident Id']}"
            body = (f"Dear User,\nIncident {row['Incident Id']} open >3 days. Details:\n"
                    f"Obs: {row['Observation']}\nSeverity: {row['Severity (S)']}\n"
                    f"Occurrence: {row['Occurrence (O)']}\nDetection: {row['Detection (D)']}\n"
                    f"RPN: {row['RPN']}\nPriority: {row['Priority']}\n"
                    f"Created: {row['Creation Date']}\n\nPlease act now.\nICSS Team")
            send_email(row['Email'], subj, body)

    return send_file(processed, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
