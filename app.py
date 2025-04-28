import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
from rapidfuzz import fuzz
from concurrent.futures import ThreadPoolExecutor
import logging
app.config['PROPAGATE_EXCEPTIONS'] = True
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

RPN_FILE = os.path.join(os.path.dirname(__file__), 'ProcessedData', 'RPN.xlsx')
if not os.path.exists(RPN_FILE):
    raise FileNotFoundError(f"RPN file not found at {RPN_FILE}")

# Load the RPN data file here
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

# Use ThreadPoolExecutor for parallel execution
executor = ThreadPoolExecutor(max_workers=4)  # You can adjust based on available CPU cores

# Helper function to send email alert
def send_email_alert(row):
    sender_email = "lakshyarubi@gmail.com"
    receiver_email = "lakshyarubi@gmail.com"
    password = "your_app_password"

    subject = f"Alert: Issue in Incident ID {row['Incident Id']}"
    body = f"""
    Incident Id: {row['Incident Id']}
    Observation: {row['Observation']}
    Creation Date: {row['Creation Date']}
    Component: {row['Component']}
    Severity: {row['Severity (S)']}
    Occurrence: {row['Occurrence (O)']}
    Detection: {row['Detection (D)']}
    RPN: {row['RPN']}
    Priority: {row['Priority']}
    """

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)  # Port 587 for STARTTLS
        server.starttls()  # Start TLS encryption
        server.login(sender_email, password)
        text = message.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        print(f"Email sent for Incident ID: {row['Incident Id']}")
    except Exception as e:
        print(f"Error sending email for {row['Incident Id']}: {e}")

@app.route('/')
def index():
    return render_template('frontNEW.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No complaint_file part", 400

    file = request.files['complaint_file']
    if file.filename == '':
        return "No selected file", 400

    month_hint = request.form.get('month_hint', 'default')

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            return f"Error reading file: {e}", 400

        if 'Observation' not in df.columns or 'Creation Date' not in df.columns or 'Incident Id' not in df.columns:
            return "Required columns missing", 400

        formatted_dates = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
        df['Creation Date'] = formatted_dates.apply(lambda x: x[0])
        days_elapsed = formatted_dates.apply(lambda x: x[1])

        # Step 3: Run NLP and matching in parallel using ThreadPoolExecutor
        df["Component"] = list(executor.map(extract_component, df["Observation"]))

        # Step 4: Get RPN values and assign priority in parallel
        rpn_values = list(executor.map(get_rpn_values, df["Component"]))
        df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = pd.DataFrame(rpn_values, index=df.index)
        df["RPN"] = df["Severity (S)"] * df["Occurrence (O)"] * df["Detection (D)"]
        df["Priority"] = df["RPN"].apply(determine_priority)

        # Step 5: Segregate and sort the Data
        spn_df = df[df["Observation"].str.contains("spn", case=False, na=False)]
        non_spn_df = df[~df["Observation"].str.contains("spn", case=False, na=False)]

        priority_order = {"High": 1, "Moderate": 2, "Low": 3}
        spn_df = spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))
        non_spn_df = non_spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))

        # Highlight and process RED highlighted rows (for Incident Status = "closed" or "complete")
        red_highlighted_rows = df[df["Incident Status"].str.contains("closed|complete", case=False, na=False)]

        # Send email for each RED highlighted row
        for _, row in red_highlighted_rows.iterrows():
            send_email_alert(row)

        # Generate Processed Excel File
        processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)

        spn_df = spn_df.fillna("")
        non_spn_df = non_spn_df.fillna("")

        with pd.ExcelWriter(processed_filepath, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            for sheet_name, sheet_df in zip(["SPN", "Non-SPN"], [spn_df, non_spn_df]):
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                green_fmt = workbook.add_format({'bg_color': '#C6EFCE'})

                for idx, row_idx in enumerate(sheet_df.index):
                    elapsed = days_elapsed.loc[row_idx]
                    color = get_color(elapsed)
                    incident_status = str(sheet_df.loc[row_idx, "Incident Status"]).lower()

                    if "closed" in incident_status or "complete" in incident_status:
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident Status"), sheet_df.loc[row_idx, "Incident Status"], green_fmt)
                    elif color:
                        fmt = workbook.add_format({'bg_color': color})
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident Id"), sheet_df.loc[row_idx, "Incident Id"], fmt)

        return send_file(processed_filepath, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

