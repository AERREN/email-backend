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

# Sign-up route
@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        password = request.form['password']
        hashed_password = generate_password_hash(password)
        users[username] = {'email': email, 'password': hashed_password}
        return redirect(url_for('login'))
    return render_template('signup.html')

# Login route
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = users.get(username)
        if user and check_password_hash(user['password'], password):
            login_user(User(username))
            return redirect(url_for('dashboard'))
        return "❌ Invalid credentials"
    return render_template('login.html')

# Dashboard route (protected)
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=current_user.id)

# Logout route
@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

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
