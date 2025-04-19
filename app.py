from flask import Flask, render_template, request, jsonify
from threading import Thread
import pandas as pd
import smtplib
from email.message import EmailMessage
import os, time, re

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

status_log = []
MAX_EMAILS_PER_SESSION = 100


@app.route('/')
def form():
    return render_template('index.html')


@app.route('/status')
def get_status():
    return jsonify(status_log)


def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)


def send_bulk_emails(smtp_server, smtp_port, sender, password, reply_to, subject, body_template, df):
    global status_log
    status_log.clear()
    failed_emails = []
    total_sent = 0

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(sender, password)
            status_log.append("✅ Logged in to SMTP server.")

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

                        try:
                            body = body_template.format(
                                first_name=row.get('first_name', ''),
                                last_name=row.get('last_name', '')
                            )
                        except KeyError as ke:
                            status_log.append(f"❌ Formatting error for {email_address}: missing {ke}")
                            failed_emails.append(email_address)
                            break

                        msg.set_content(body)
                        smtp.send_message(msg)
                        status_log.append(f"✅ Sent to {email_address}")
                        sent = True
                        total_sent += 1
                        time.sleep(5)

                    except Exception as e:
                        retries += 1
                        status_log.append(f"❌ Retry {retries} for {email_address}: {str(e)}")
                        time.sleep(5)

                if not sent and email_address not in failed_emails:
                    failed_emails.append(email_address)

        if failed_emails:
            status_log.append(f"❌ Failed to send to: {failed_emails}")
        else:
            status_log.append("✅ All emails sent successfully!")

    except smtplib.SMTPAuthenticationError:
        status_log.append("❌ Authentication failed. Check your email or app password.")
    except Exception as e:
        status_log.append(f"❌ General error: {str(e)}")


@app.route('/send', methods=['POST'])
def send_emails():
    try:
        smtp_server = request.form['smtp_server']
        smtp_port = int(request.form['smtp_port'])
        sender = request.form['sender_email']
        password = request.form['sender_password']
        reply_to = request.form['reply_to']
        subject = request.form['subject']
        body_template = request.form['body']
        file = request.files['file']

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
            smtp_server, smtp_port, sender, password, reply_to, subject, body_template, df))
        thread.start()

        return "📨 Sending emails... Check progress below."

    except Exception as e:
        return f"❌ Error during processing: {str(e)}"


if __name__ == '__main__':
    app.run(debug=True)
