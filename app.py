from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from datetime import datetime
import xlsxwriter
from rapidfuzz import fuzz
from concurrent.futures import ThreadPoolExecutor
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

RPN_FILE = os.path.join(os.path.dirname(__file__), 'ProcessedData', 'RPN.xlsx')
if not os.path.exists(RPN_FILE):
    raise FileNotFoundError(f"RPN file not found at {RPN_FILE}")

# Load RPN data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

# Thread pool for parallel tasks
executor = ThreadPoolExecutor(max_workers=4)

def extract_component(obs):
    obs = str(obs).strip()
    best_match, highest_score = None, 0
    for comp in known_components:
        score = fuzz.partial_ratio(comp.lower(), obs.lower())
        if score > highest_score and score >= 80:
            best_match, highest_score = comp, score
    return best_match or "Unknown"

def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        return (int(row["Severity (S)"].values[0]),
                int(row["Occurrence (O)"].values[0]),
                int(row["Detection (D)"].values[0]))
    return 1, 1, 10

def determine_priority(rpn):
    return "High" if rpn >= 200 else "Moderate" if rpn >= 100 else "Low"

def month_str_to_num(month_hint):
    month_map = {"jan":"01","feb":"02","mar":"03","apr":"04",
                 "may":"05","jun":"06","jul":"07","aug":"08",
                 "sep":"09","oct":"10","nov":"11","dec":"12"}
    return month_map.get(month_hint.lower())

def format_creation_date(date_str, month_hint):
    target_month = month_str_to_num(month_hint)
    if not target_month:
        return None, None
    try:
        dt = pd.to_datetime(str(date_str).strip(), errors='coerce', dayfirst=True)
        if pd.notna(dt):
            dd, mm, yyyy = dt.day, dt.month, dt.year
            if str(dd).zfill(2)=="01" and str(mm).zfill(2)=="01":
                dd, mm = mm, int(target_month)
            formatted = f"{str(dd).zfill(2)}/{target_month}/{yyyy}"
            return formatted, (datetime.now() - dt).days
    except:
        return None, None
    return None, None

# Email alert function
def send_alert_email(df_filtered):
    if df_filtered.empty:
        return
    sender_email   = "lakshyarubi@gmail.com"
    receiver_email = "lakshyarubi.gnana2021@vitstudent.ac.in"
    app_password   = "selr fdih wlkm wufg"

    html_table = df_filtered.to_html(index=False)
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "🚨 Escalated Incidents (3+ days)"
    msg["From"]    = sender_email
    msg["To"]      = receiver_email

    html_body = f"""
    <html>
      <body style="font-family:Arial,sans-serif;">
        <h3>🚨 Open & Pending Incidents Escalated ≥ 3 Days</h3>
        <p>Generated: {datetime.now().strftime('%d %b %Y, %H:%M:%S')}</p>
        {html_table}
        <p>Regards,<br/>ICSS Team</p>
      </body>
    </html>
    """
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
            print("Email alert sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

@app.route('/')
def index():
    return render_template('frontNEW.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('complaint_file')
    if not file or file.filename=='':
        return "No file provided", 400

    month_hint = request.form.get('month_hint','default')
    path       = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    try:
        df = pd.read_excel(path)
    except Exception as e:
        return f"Error reading file: {e}", 400

    required = ['Observation','Creation Date','Incident Id']
    if not all(col in df.columns for col in required):
        return "Required columns missing", 400

    # Format dates & compute elapsed days
    fmt = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
    df['Creation Date'] = fmt.apply(lambda x: x[0])
    days_elapsed         = fmt.apply(lambda x: x[1])
    df['Days Elapsed']   = days_elapsed

    # Month abbreviation
    df['Creation_DT']    = pd.to_datetime(df['Creation Date'], dayfirst=True, errors='coerce')
    df['Month']          = df['Creation_DT'].dt.strftime('%b')
    df.drop(columns=['Creation_DT'], inplace=True)

    # Component matching & RPN calculation
    df['Component'] = list(executor.map(extract_component, df['Observation']))
    rpn_vals       = list(executor.map(get_rpn_values, df['Component']))
    df[['Severity (S)','Occurrence (O)','Detection (D)']] = pd.DataFrame(rpn_vals, index=df.index)
    df['RPN']      = df['Severity (S)']*df['Occurrence (O)']*df['Detection (D)']
    df['Priority'] = df['RPN'].apply(determine_priority)

    # Segregate and write Excel
    spn_df    = df[df['Observation'].str.contains('spn',case=False,na=False)]
    non_spn   = df[~df['Observation'].str.contains('spn',case=False,na=False)]
    order_map = {'High':1,'Moderate':2,'Low':3}
    spn_df    = spn_df.sort_values(by='Priority', key=lambda x: x.map(order_map))
    non_spn   = non_spn.sort_values(by='Priority', key=lambda x: x.map(order_map))

    out_path = os.path.join(UPLOAD_FOLDER, 'processed_'+file.filename)
    with pd.ExcelWriter(out_path, engine='xlsxwriter', engine_kwargs={'options':{'nan_inf_to_errors':True}}) as writer:
        for name, sheet in [('SPN',spn_df),('Non-SPN',non_spn)]:
            sheet.fillna('', inplace=True)
            sheet.to_excel(writer, sheet_name=name, index=False)
            wb = writer.book
            ws = writer.sheets[name]
            green = wb.add_format({'bg_color':'#C6EFCE'})
            for i, idx in enumerate(sheet.index):
                elapsed = days_elapsed.loc[idx]
                status  = str(sheet.at[idx,'Incident Status']).lower()
                if 'closed' in status or 'complete' in status:
                    ws.write(i+1, sheet.columns.get_loc('Incident Status'), sheet.at[idx,'Incident Status'], green)
    
    # Send only escalated open/pending
    alert_df = df[(df['Incident Status'].str.lower().isin(['open','pending'])) & (df['Days Elapsed']>=3)]
    cols     = ['Sr.No','Incident Id','Creation Date','Month','Days Elapsed','KVA Rating',
                'Engine no','Customer VOC','SR Status','Account Name','Service Dealer Name',
                'Running hours','Observation','Incident Status','ASM Name']
    alert_df = alert_df[cols]
    executor.submit(send_alert_email, alert_df)

    return send_file(out_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',5000)))
