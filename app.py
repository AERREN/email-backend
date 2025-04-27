from flask import Flask, render_template, request, jsonify, redirect, url_for, session
from flask_cors import CORS
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from threading import Thread
import pandas as pd
import smtplib
from email.message import EmailMessage
import ssl
import os, time, re
import docx
from werkzeug.security import generate_password_hash, check_password_hash

# Initialize Flask app and Flask-Login
app = Flask(__name__)
app.secret_key = os.urandom(24)  # Secret key for sessions
CORS(app)

login_manager = LoginManager()
login_manager.init_app(app)

# User Database (for demo purposes)
users = {}

# Constants and File Handling
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

status_log = []
MAX_EMAILS_PER_SESSION = 300
DELAY_BETWEEN_EMAILS = 4  # seconds

DEFAULT_SMTP_CONFIGS = {
    'brevo': {'server': 'smtp-relay.brevo.com', 'port': 587},
    'gmail': {'server': 'smtp.gmail.com', 'port': 587},
    'yandex': {'server': 'smtp.yandex.com', 'port': 465},
    'zoho': {'server': 'smtp.zoho.com', 'port': 465}
}

# Flask-Login User Class
class User(UserMixin):
    def __init__(self, id):
        self.id = id

@login_manager.user_loader
def load_user(user_id):
    return User(user_id) if user_id in users else None

# ------------ FRONTEND ROUTES ------------
@app.route('/')
def form():
    return render_template('index.html')

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=current_user.id)

@app.route('/status')
def get_status():
    return jsonify(status_log)

# ------------ AUTH APIs (for frontend JS) ------------

@app.route('/register', methods=['POST'])
def api_register():
    data = request.get_json()
    username = data['username']
    email = data['email']
    password = data['password']
    if username in users:
        return jsonify({"error": "User already exists!"}), 400
    hashed_password = generate_password_hash(password)
    users[username] = {'email': email, 'password': hashed_password, 'activities': []}
    return jsonify({"message": "Registered successfully!"})

@app.route('/login', methods=['POST'])
def api_login():
    data = request.get_json()
    username = data['username']
    password = data['password']
    user = users.get(username)
    if user and check_password_hash(user['password'], password):
        login_user(User(username))
        return jsonify({"message": "Login successful!"})
    return jsonify({"error": "Invalid credentials"}), 401

@app.route('/userdata', methods=['GET'])
@login_required
def userdata():
    username = current_user.id
    user_data = users.get(username)
    if not user_data:
        return jsonify({"error": "User not found"}), 404
    return jsonify({
        "username": username,
        "email": user_data['email'],
        "activities": user_data['activities']
    })

@app.route('/logout', methods=['GET'])
@login_required
def api_logout():
    logout_user()
    return jsonify({"message": "Logged out successfully!"})

# ------------ EMAIL SENDER ------------

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

@app.route('/send', methods=['POST'])
@login_required
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

def send_bulk_emails(smtp_server, smtp_port, sender, password, reply_to, subject, body_template, df, is_html):
    smtp = create_smtp_connection(smtp_server, smtp_port, sender, password)
    if not smtp:
        return

    try:
        for index, row in df.iterrows():
            if index >= MAX_EMAILS_PER_SESSION:
                break

            to_email = str(row['email']).strip()
            if not is_valid_email(to_email):
                status_log.append(f"❌ Invalid email address: {to_email}")
                continue

            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = sender
            msg['To'] = to_email
            msg['Reply-To'] = reply_to
            if is_html:
                msg.add_alternative(body_template, subtype='html')
            else:
                msg.set_content(body_template)

            try:
                smtp.send_message(msg)
                status_log.append(f"✅ Sent to {to_email}")
            except Exception as e:
                status_log.append(f"❌ Failed to send to {to_email}: {str(e)}")

            time.sleep(DELAY_BETWEEN_EMAILS)
    finally:
        smtp.quit()

# ------------ START SERVER ------------

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000)
