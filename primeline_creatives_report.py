import os
import json
import smtplib
import ssl
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials
from jinja2 import Template
import base64
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

# Configuration
SPREADSHEET_ID = '1onzfa2zoPGgEAD8efVaGpcQYGex9s5iKCFke0btmnk0'
SMTP_HOST = 'secure.emailsrvr.com'
SMTP_PORT = 465
SENDER_EMAIL = 'mayur.kambli@artworkservicesusa.com'

# Recipients
RECIPIENTS_TO = ['mayur.online9@gmail.com']
RECIPIENTS_CC = ['mayur.kambli@artworkservicesusa.com']
RECIPIENTS_BCC = ['mayur.kambli@artworkservicesusa.com']

def get_google_client():
    creds_raw = os.environ.get('GOOGLE_CREDENTIALS_JSON', '').strip()
    if not creds_raw:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")
    
    creds_dict = None
    # 1. Try Base64 decoding
    try:
        decoded = base64.b64decode(creds_raw).decode('utf-8')
        if decoded.strip().startswith('{'):
            creds_dict = json.loads(decoded)
            print("Successfully loaded credentials via Base64.")
    except Exception:
        pass

    # 2. Fallback to raw JSON
    if not creds_dict:
        try:
            cleaned_raw = creds_raw.replace('\r\n', '\\n').replace('\n', '\\n')
            creds_dict = json.loads(cleaned_raw)
            print("Successfully loaded credentials via Raw JSON (auto-fixed).")
        except Exception:
            try:
                creds_dict = json.loads(creds_raw)
                print("Successfully loaded credentials via Raw JSON.")
            except json.JSONDecodeError as e:
                raise ValueError(f"CRITICAL: GOOGLE_CREDENTIALS_JSON is malformed. Error: {e}")

    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def get_target_date():
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        target_date = today - timedelta(days=3)
    elif today.weekday() == 6:  # Sunday
        target_date = today - timedelta(days=2)
    else:
        target_date = today - timedelta(days=1)
    
    return f"{target_date.day}-{target_date.strftime('%b-%y')}"

def fetch_data(target_date):
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    
    print(f"Target Date: {target_date}")
    
    today = datetime.now()
    today_str = f"{today.day}-{today.strftime('%b-%y')}"
    
    # Get Data from Primeline Creatives
    try:
        ws_trends = sh.worksheet("Primeline Creatives")
    except gspread.exceptions.WorksheetNotFound:
        try:
            ws_trends = sh.worksheet("Primeline-Creatives")
        except gspread.exceptions.WorksheetNotFound:
            print(f"Available sheets: {[w.title for w in sh.worksheets()]}")
            # Fallback to the first sheet
            ws_trends = sh.worksheets()[0]
            print(f"Fell back to using sheet: {ws_trends.title}")
            
    all_data = ws_trends.get_all_values()
    if not all_data:
        return target_date, 0, 0, 0, 0, []

    rows = all_data[1:]
    
    emails_received = 0
    emails_completed = 0
    total_completed_items = 0
    pending_count = 0
    detailed_rows = []

    for row in rows:
        if len(row) < 14: # Pad row if shorter than index N
            row = row + [""] * (14 - len(row))
        
        row_date = str(row[0]).strip()
        email_subject = str(row[6]).strip() # Column G
        done_date = str(row[13]).strip() # Column N
        total_items = str(row[9]).strip() # Column J
        
        if row_date == today_str:
            continue
            
        is_pending = not done_date or done_date.lower() == 'pending'
        
        if is_pending and not email_subject:
            continue
            
        # 1. Emails Received: Date matches target_date
        if row_date == target_date:
            emails_received += 1
            
        # 2. Emails Completed: Done date matches target_date
        if done_date == target_date:
            emails_completed += 1
            try:
                total_completed_items += int(total_items) if total_items else 0
            except:
                pass
        
        # 3. Pending: Done date is "Pending" or empty (but has a Date in col A)
        if is_pending and row_date:
            pending_count += 1
            
        # 4. Detailed Rows
        if (done_date == target_date or is_pending) and row_date:
            detailed_rows.append([
                row[0], 
                row[5], 
                row[6], 
                row[9] if not is_pending else "", # Column J
                row[13] if row[13].strip() else "Pending" # Column N
            ])

    return emails_received, emails_completed, total_completed_items, pending_count, detailed_rows

def format_html(target_date, emails_received, emails_completed, total_completed_items, pending_count, detailed_rows):
    if emails_received == 0 and emails_completed == 0 and pending_count == 0:
        template_str = """
        <html>
        <head>
        <style>
            body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
            table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; }
            td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
            .header-cell { font-weight: bold; }
        </style>
        </head>
        <body>
            <p>Hi team,</p>
            <p>Please see below summary.</p>
            <table>
                <tr><td class="header-cell">Primeline-creatives</td><td></td></tr>
                <tr><td class="header-cell">Date</td><td class="header-cell">Emails received</td></tr>
                <tr><td>{{ target_date }}</td><td>No orders received</td></tr>
            </table>
            <br>
            <p>Thanks and Regards,<br>Mayur</p>
        </body>
        </html>
        """
        template = Template(template_str)
        return template.render(target_date=target_date)

    template_str = """
    <html>
    <head>
    <style>
        body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
        table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; width: auto; min-width: 400px; }
        td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
        .header-cell { font-weight: bold; }
        .detail-table { width: 100%; max-width: 800px; }
        .section-title { font-weight: bold; text-align: center; }
    </style>
    </head>
    <body>
        <p>Hi team,</p>
        <p>Please see below mentioned summary report for your reference.</p>
        
        <!-- Summary Table -->
        <table>
            <tr><td colspan="4" class="header-cell">Primeline-creatives</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails received</td><td>Emails completed</td><td>Total virtual completed</td>
            </tr>
            <tr>
                <td>{{ target_date }}</td><td>{{ emails_received }}</td><td>{{ emails_completed }}</td><td>{{ total_completed_items }}</td>
            </tr>
            <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
            <tr><td class="header-cell">Pending</td><td>{{ pending_count }}</td><td></td><td></td></tr>
        </table>

        <br>

        <!-- Detailed Table -->
        <table class="detail-table">
            <tr><td colspan="5" class="section-title">Primeline-creatives</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails from</td><td>Email subject</td><td>Count</td><td>Done date</td>
            </tr>
            {% for row in detailed_rows %}
            <tr>
                <td>{{ row[0] }}</td><td>{{ row[1] }}</td><td>{{ row[2] }}</td><td>{{ row[3] }}</td><td>{{ row[4] }}</td>
            </tr>
            {% endfor %}
        </table>
        
        <br>
        <p>Thanks and Regards,<br>Mayur</p>
    </body>
    </html>
    """
    template = Template(template_str)
    return template.render(
        target_date=target_date,
        emails_received=emails_received,
        emails_completed=emails_completed,
        total_completed_items=total_completed_items,
        pending_count=pending_count,
        detailed_rows=detailed_rows
    )

def capture_screenshot(html_content, output_path="report_screenshot.png"):
    # Write temporary HTML file
    temp_html = "temp_report.html"
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    # We want a full-path URI for playwright
    abs_path = "file:///" + os.path.abspath(temp_html).replace("\\", "/")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={'width': 1000, 'height': 800})
        page.goto(abs_path)
        # Wait for any potential rendering
        page.wait_for_load_state("networkidle")
        
        # Take screenshot of the body so it frames the content nicely
        body = page.locator("body")
        body.screenshot(path=output_path)
        browser.close()

def send_email(subject, html_content, screenshot_path):
    password = os.environ.get('SMTP_PASSWORD')
    if not password:
        raise ValueError("SMTP_PASSWORD environment variable not set")
    
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENTS_TO)
    msg['Cc'] = ", ".join(RECIPIENTS_CC)
    msg['Bcc'] = ", ".join(RECIPIENTS_BCC)
    
    msg.set_content("Please enable HTML to view this report.")
    msg.add_alternative(html_content, subtype='html')
    
    if os.path.exists(screenshot_path):
        with open(screenshot_path, 'rb') as f:
            img_data = f.read()
        msg.add_attachment(img_data, maintype='image', subtype='png', filename='report_summary.png')
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SENDER_EMAIL, password)
        server.send_message(msg)

def main():
    try:
        target_date = get_target_date()
        print(f"Fetching data from Google Sheets for {target_date}...")
        emails_received, emails_completed, total_items, pending, detailed = fetch_data(target_date)
        
        print("Generating HTML report...")
        html_content = format_html(target_date, emails_received, emails_completed, total_items, pending, detailed)
        
        screenshot_path = "report_screenshot.png"
        print("Capturing screenshot...")
        capture_screenshot(html_content, screenshot_path)
        
        subject = f"Primeline-Creatives Summary : {target_date}"
        print(f"Sending email: {subject}")
        send_email(subject, html_content, screenshot_path)
        
        print("Success!")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        exit(1)

if __name__ == "__main__":
    main()
