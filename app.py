import os
import re
import smtplib
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, flash, session, url_for, jsonify
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

UPLOAD_FOLDER = './uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.secret_key = 'supersecretkeychange'  # Replace for prod
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Neon Tech DB connection string
app.config['SQLALCHEMY_DATABASE_URI'] = "postgresql://neondb_owner:npg_1RVftoYeCM3U@ep-plain-base-a1i2tceb-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# --- DB Model mapping your table ---
class Email(db.Model):
    __tablename__ = 'email'  # your existing table
    id = db.Column(db.Integer, primary_key=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_sent = db.Column(db.Boolean, default=False)
    sent_at = db.Column(db.DateTime)
    email = db.Column(db.String(255), unique=True, nullable=False)

# --- Helpers ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def get_sender_accounts():
    return session.get('sender_accounts', [])

def set_sender_accounts(accounts):
    session['sender_accounts'] = accounts

# --- Routes ---

@app.route('/')
def dashboard():
    total = Email.query.count()
    sent = Email.query.filter_by(is_sent=True).count()
    pending = total - sent
    today_start = datetime.combine(date.today(), datetime.min.time())
    today_sent = Email.query.filter(Email.sent_at >= today_start).count()
    sending_progress = f"{today_sent} / 4000"
    return render_template('dashboard.html',
                           total=total, sent=sent, pending=pending,
                           sending_progress=sending_progress,
                           sender_accounts=get_sender_accounts())

@app.route('/upload-excel', methods=['GET', 'POST'])
def upload_excel():
    log = None
    if request.method == 'POST':
        files = request.files.getlist('files')
        processed = valid = duplicates = inserted = 0
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                df = pd.read_excel(filepath)
                # Auto detect email column by common keywords
                email_col = next((col for col in df.columns if 'email' in col.lower()), None)
                if not email_col:
                    flash(f"No email column found in file {filename}", "danger")
                    continue
                emails = df[email_col].dropna().astype(str).str.strip()
                processed += len(emails)
                for em in emails:
                    if validate_email(em):
                        valid += 1
                        if not Email.query.filter_by(email=em).first():
                            new_email = Email(email=em)
                            db.session.add(new_email)
                            inserted += 1
                        else:
                            duplicates += 1
                    else:
                        # invalid email skips
                        pass
                db.session.commit()
        log = dict(processed=processed, valid=valid, duplicates=duplicates, inserted=inserted)
    return render_template('upload_excel.html', log=log)

@app.route('/sender-accounts', methods=['GET', 'POST'])
def sender_accounts():
    if 'sender_accounts' not in session:
        session['sender_accounts'] = []
    accounts = get_sender_accounts()
    if request.method == 'POST':
        if len(accounts) >= 10:
            flash("Max 10 sender accounts allowed", "danger")
        else:
            email = request.form['email'].strip()
            password = request.form['password'].strip()
            if not validate_email(email):
                flash("Invalid email address", "danger")
            else:
                if any(a['email'] == email for a in accounts):
                    flash("Sender account already added", "danger")
                else:
                    accounts.append({"email": email, "password": password})
                    set_sender_accounts(accounts)
                    flash("Sender account added", "success")
        return redirect(url_for('sender_accounts'))
    return render_template('sender_accounts.html', accounts=accounts)

@app.route('/remove-account/<email>', methods=['POST'])
def remove_account(email):
    accounts = get_sender_accounts()
    accounts = [acc for acc in accounts if acc['email'] != email]
    set_sender_accounts(accounts)
    flash(f"Removed sender account {email}", "success")
    return redirect(url_for('sender_accounts'))

@app.route('/campaign', methods=['GET', 'POST'])
def campaign():
    if request.method == 'POST':
        subject = request.form.get('subject', '').strip()
        body = request.form.get('body', '').strip()
        attachment = request.files.get('attachment')
        filename = None
        if attachment and attachment.filename != '':
            filename = secure_filename(attachment.filename)
            attachment.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        session['campaign'] = {
            'subject': subject,
            'body': body,
            'attachment': filename
        }
        flash("Campaign setup saved for this run", "success")
        return redirect(url_for('dashboard'))
    campaign = session.get('campaign', {})
    return render_template('campaign.html', campaign=campaign)

@app.route('/activate-sending', methods=['POST'])
def activate_sending():
    sender_accounts = get_sender_accounts()
    if not sender_accounts:
        flash("Add at least one sender account to send emails", "danger")
        return redirect(url_for('dashboard'))
    campaign = session.get('campaign')
    if not campaign or not campaign.get('subject') or not campaign.get('body'):
        flash("Set campaign subject and body before sending", "danger")
        return redirect(url_for('dashboard'))

    limit_per_account = 400
    total_limit = 4000
    emails_to_send = Email.query.filter_by(is_sent=False).limit(total_limit).all()
    if not emails_to_send:
        flash("No unsent emails available", "warning")
        return redirect(url_for('dashboard'))

    accounts_count = len(sender_accounts)
    emails_per_account = min(limit_per_account, total_limit // accounts_count)

    # Prepare attachment path if any
    attachment_path = None
    if campaign.get('attachment'):
        attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], campaign['attachment'])

    sent_count = 0
    errors = []

    for i, account in enumerate(sender_accounts):
        emails_batch = emails_to_send[i*emails_per_account:(i+1)*emails_per_account]
        for email_obj in emails_batch:
            try:
                send_email(
                    to_email=email_obj.email,
                    subject=campaign['subject'],
                    body=campaign['body'],
                    from_email=account['email'],
                    app_password=account['password'],
                    attachment_path=attachment_path
                )
                email_obj.is_sent = True
                email_obj.sent_at = datetime.utcnow()
                db.session.commit()
                sent_count += 1
            except Exception as e:
                errors.append(f"{email_obj.email} - {str(e)}")

    flash(f"Sent {sent_count} emails in this batch.", "success")
    if errors:
        flash(f"Errors occurred for {len(errors)} emails. Check logs.", "danger")
        print("Email send errors:", errors)
    return redirect(url_for('dashboard'))

# Email sending function using smtplib
def send_email(to_email, subject, body, from_email, app_password, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Replace placeholders if any (extend if needed)
    body_text = body.replace("{email}", to_email)
    msg.attach(MIMEText(body_text, 'plain'))

    if attachment_path and os.path.isfile(attachment_path):
        part = MIMEBase('application', 'octet-stream')
        with open(attachment_path, 'rb') as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, app_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

# --- Run app ---
if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    with app.app_context():
        db.create_all()  # Create tables if not exist inside app context
    app.run(host='0.0.0.0', port=5000, debug=True)
