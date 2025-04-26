from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
from threading import Thread
import pandas as pd
import smtplib
from email.message import EmailMessage
import ssl
import os, time, re
import docx

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

status_log = []
MAX_EMAILS_PER_SESSION = 300
DELAY_BETWEEN_EMAILS = 4  # seconds

DEFAULT_SMTP_CONFIGS = {
    'brevo': {
        'server': 'smtp-relay.brevo.com',
        'port': 587
    },
    'gmail': {
        'server': 'smtp.gmail.com',
        'port': 587
    },
    'yandex': {
        'server': 'smtp.yandex.com',
        'port': 465
    },
    'zoho': {
        'server': 'smtp.zoho.com',
        'port': 465
    }
}

@app.route('/')
def form():
    return render_template('index.html')

@app.route('/status')
def get_status():
    return jsonify(status_log)

def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def create_smtp_connection(server, port, sender, password):
    try:
        context = ssl.create_default_context()
        if port == 465:
            smtp = smtplib.SMTP_SSL(server, port, context=context)
        else:
            smtp = smtplib.SMTP(server, port)
            smtp.ehlo()
            smtp.starttls(context=context)
        smtp.login(sender, password)
        return smtp
    except Exception as e:
        status_log.append(f"❌ SMTP connection failed: {str(e)}")
        return None

def send_bulk_emails(smtp_server, smtp_port, sender, password, reply_to, subject, body_template, df, is_html):
    global status_log
    status_log.clear()
    failed_emails = []
    total_sent = 0

    smtp = create_smtp_connection(smtp_server, smtp_port, sender, password)
    if not smtp:
        return

    for index, row in df.iterrows():
        if total_sent >= MAX_EMAILS_PER_SESSION:
            status_log.append("🛑 Spam limit reached (100 emails per session).")
            break

        email_address = row.get('email', '').strip()
        if not is_valid_email(email_address):
            status_log.append(f"⚠️ Skipped invalid email: {email_address}")
            continue

        retries = 0
        sent = False

        while retries < 3 and not sent:
            try:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['To'] = email_address
                msg['From'] = sender
                msg['Reply-To'] = reply_to
                msg['Return-Path'] = sender
                msg['X-Priority'] = '3'  # Normal priority
                msg['X-Mailer'] = 'Python SMTP Bulk Mailer'
                msg['List-Unsubscribe'] = f"<mailto:{reply_to}>"  # Helps reduce spam score

                try:
                    body = body_template.format(
                        first_name=row.get('first_name', ''),
                        last_name=row.get('last_name', '')
                    )
                except KeyError as ke:
                    status_log.append(f"❌ Formatting error for {email_address}: missing {ke}")
                    failed_emails.append(email_address)
                    break

                if is_html:
                    msg.add_alternative(body, subtype='html')
                else:
                    msg.set_content(body)

                smtp.send_message(msg)
                status_log.append(f"✅ Sent to {email_address}")
                sent = True
                total_sent += 1
                time.sleep(DELAY_BETWEEN_EMAILS)

            except Exception as e:
                retries += 1
                status_log.append(f"❌ Retry {retries} for {email_address}: {str(e)}")
                time.sleep(DELAY_BETWEEN_EMAILS)

        if not sent and email_address not in failed_emails:
            failed_emails.append(email_address)

    smtp.quit()

    if failed_emails:
        status_log.append(f"❌ Failed to send to: {failed_emails}")
    else:
        status_log.append("✅ All emails sent successfully!")

@app.route('/send', methods=['POST'])
def send_emails():
    try:
        smtp_server = request.form['smtp_server'].strip()
        smtp_port = int(request.form['smtp_port'])
        sender = request.form['sender_email'].strip()
        password = request.form['sender_password']
        reply_to = request.form['reply_to'].strip()
        subject = request.form['subject']
        body_template = request.form['body']
        file = request.files['file']

        provider = request.form.get('provider', '').lower()
        if provider in DEFAULT_SMTP_CONFIGS:
            smtp_server = DEFAULT_SMTP_CONFIGS[provider]['server']
            smtp_port = DEFAULT_SMTP_CONFIGS[provider]['port']

        html_file = request.files.get('html_file')
        docx_file = request.files.get('docx_file')
        is_html = False

        if html_file and html_file.filename.endswith('.html'):
            body_template = html_file.read().decode('utf-8')
            is_html = True
        elif docx_file and docx_file.filename.endswith('.docx'):
            doc = docx.Document(docx_file)
            body_template = '\n'.join([para.text for para in doc.paragraphs])
            is_html = False

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        if filepath.endswith('.xlsx'):
            df = pd.read_excel(filepath)
        elif filepath.endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            return "Unsupported file format. Use CSV or XLSX."

        if 'email' not in df.columns:
            return "Missing 'email' column in uploaded file."

        thread = Thread(target=send_bulk_emails, args=(
            smtp_server, smtp_port, sender, password, reply_to, subject, body_template, df, is_html
        ))
        thread.start()

        return "📨 Sending emails... Check progress below."

    except Exception as e:
        return f"❌ Error during processing: {str(e)}"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000)
