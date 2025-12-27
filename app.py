import json
import pandas as pd
import platform
import subprocess
import os
import re
import smtplib
import random
import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from xhtml2pdf import pisa
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)
CORS(app)

# --- CONFIGURATION ---
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'students.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
EXCEL_FILE = os.path.join(basedir, 'Student_Data.xlsx')

# ==========================================
# 👇 MASTER SYSTEM EMAIL (Isi se sab mail jayenge) 👇
# ==========================================
SYSTEM_EMAIL = "manumaurya006@gmail.com"  
SYSTEM_PASSWORD = "dptk jxcg szmo rkyy"  # 16-digit App Password

SERVICE_ACCOUNT_FILE = 'credentials.json'
SCOPES = ['https://www.googleapis.com/auth/drive']
DRIVE_FOLDER_ID = '1aBcD_xYz123456789_QrStUvWxYz' 
# ==========================================

# --- Database Models ---
class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(100), unique=True, nullable=False) 
    password = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), nullable=False)
    otp = db.Column(db.String(6))
    otp_expiry = db.Column(db.DateTime)

class Student(db.Model):
    id = db.Column(db.String(50), primary_key=True)
    enrollment_number = db.Column(db.String(50), unique=True, nullable=False)
    student_name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), nullable=False) 
    course = db.Column(db.String(50))
    semester = db.Column(db.String(20))
    subjects_json = db.Column(db.Text, nullable=False, default="[]")
    attendance_marks = db.Column(db.Float, default=0.0)
    internship_marks = db.Column(db.Float, default=0.0)
    project_marks = db.Column(db.Float, default=0.0)
    timestamp = db.Column(db.String(50))

    def to_dict(self):
        data = {c.name: getattr(self, c.name) for c in self.__table__.columns}
        data['subjects'] = json.loads(self.subjects_json)
        return data

with app.app_context():
    db.create_all()

# --- Google Drive Logic ---
def authenticate_drive():
    if not os.path.exists(SERVICE_ACCOUNT_FILE): return None
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(filename):
    try:
        service = authenticate_drive()
        if not service: return
        file_metadata = {'name': os.path.basename(filename)}
        if DRIVE_FOLDER_ID: file_metadata['parents'] = [DRIVE_FOLDER_ID]
        media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        query = f"name = '{os.path.basename(filename)}' and trashed = false"
        if DRIVE_FOLDER_ID: query += f" and '{DRIVE_FOLDER_ID}' in parents"
        results = service.files().list(q=query, fields="files(id)").execute()
        items = results.get('files', [])

        if not items:
            service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        else:
            service.files().update(fileId=items[0]['id'], media_body=media).execute()
        print(">> Drive Sync Successful")
    except Exception as e: print(f"Drive Error: {e}")

# --- Excel Sync Logic ---
def sync_with_excel():
    students = Student.query.all()
    data_list = []
    for s in students:
        subjects = json.loads(s.subjects_json)
        subj_str = ", ".join([f"{sub['name']} (E:{sub['exam']}+I:{sub['internal']})" for sub in subjects])
        theory = sum([sub['exam'] + sub['internal'] for sub in subjects])
        total = theory + s.attendance_marks + s.internship_marks + s.project_marks
        
        data_list.append({
            'Department': s.department, 
            'Enrollment': s.enrollment_number, 'Name': s.student_name, 'Email': s.email,
            'Password': s.password, 'Course': s.course, 'Semester': s.semester, 
            'Subjects': subj_str, 'Attendance': s.attendance_marks, 
            'Internship': s.internship_marks, 'Project': s.project_marks, 'Grand Total': total
        })

    try:
        if data_list:
            pd.DataFrame(data_list).to_excel(EXCEL_FILE, index=False)
            upload_to_drive(EXCEL_FILE)
        elif os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
    except Exception as e: print(f"Excel Error: {e}")

# --- HELPER: GENERATE HTML FOR PDF ---
def get_marksheet_html(student):
    subjects = json.loads(student.subjects_json)
    rows = "".join([f"<tr><td style='padding:8px; border:1px solid #ddd;'>{s['name']}</td><td style='padding:8px; border:1px solid #ddd;'>{s['exam']}</td><td style='padding:8px; border:1px solid #ddd;'>{s['internal']}</td><td style='padding:8px; border:1px solid #ddd;'><b>{s['exam']+s['internal']}</b></td></tr>" for s in subjects])
    grand_total = sum([s['exam']+s['internal'] for s in subjects]) + student.attendance_marks + student.internship_marks + student.project_marks
    perc = (grand_total / ((len(subjects)*100)+110)) * 100 if subjects else 0
    status = "PASS" if perc >= 40 else "FAIL"
    color = "green" if status == "PASS" else "red"
    
    return f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <div style="border: 2px solid #333; padding: 20px; max-width: 800px; margin: auto;">
            <h1 style="text-align: center; color: #2563eb; margin-bottom: 5px;">JNCT COLLEGE, BHOPAL</h1>
            <h3 style="text-align: center; color: #555; margin-top: 0px;">Department of {student.department}</h3>
            <hr>
            <p><b>Name:</b> {student.student_name}</p>
            <p><b>Enrollment:</b> {student.enrollment_number}</p>
            <p><b>Course:</b> {student.course} | <b>Sem:</b> {student.semester}</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
                <tr style="background-color: #f2f2f2;">
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Subject</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Exam</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Internal</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Total</th>
                </tr>
                {rows}
            </table>
            <div style="margin-top: 20px;">
                <p><b>Attendance:</b> {student.attendance_marks} | <b>Internship:</b> {student.internship_marks} | <b>Project:</b> {student.project_marks}</p>
                <h2 style="text-align: right;">Total: {grand_total}</h2>
                <h2 style="text-align: right; color: {color};">RESULT: {status} ({perc:.2f}%)</h2>
            </div>
            <hr>
            <p style="text-align: center; font-size: 12px; color: #777;">Computer Generated Report - JNCT College</p>
        </div>
    </body>
    </html>"""

# --- API ROUTES ---

@app.route('/')
def index():
    with open('index.html', 'r', encoding='utf-8') as f: return f.read()

@app.route('/api/admin/register', methods=['POST'])
def admin_register():
    data = request.json
    if Admin.query.filter_by(email=data['email']).first():
        return jsonify({'error': 'Email already registered'}), 400
    
    new_admin = Admin(
        email=data['email'], 
        password=data['password'], 
        department=data['department']
    )
    db.session.add(new_admin)
    db.session.commit()
    return jsonify({'message': 'Department Registered Successfully!'})

@app.route('/api/admin/login', methods=['POST'])
def admin_login():
    data = request.json
    admin = Admin.query.filter_by(email=data['email'], password=data['password']).first()
    if admin:
        return jsonify({'message': 'Login Successful', 'email': admin.email, 'department': admin.department})
    return jsonify({'error': 'Invalid Credentials'}), 401

@app.route('/api/admin/send-otp', methods=['POST'])
def send_otp():
    data = request.json
    email = data.get('email')
    admin = Admin.query.filter_by(email=email).first()
    if not admin: return jsonify({'error': 'Email not registered'}), 404
    
    otp = str(random.randint(100000, 999999))
    admin.otp = otp
    admin.otp_expiry = datetime.datetime.now() + datetime.timedelta(minutes=10)
    db.session.commit()
    
    try:
        msg = MIMEText(f"Your Password Reset OTP is: {otp}", 'plain')
        msg['From'] = SYSTEM_EMAIL
        msg['To'] = email 
        msg['Subject'] = "JNCT Admin Password Reset"
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SYSTEM_EMAIL, SYSTEM_PASSWORD)
        server.send_message(msg)
        server.quit()
        return jsonify({'message': 'OTP Sent'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/api/admin/reset-password', methods=['POST'])
def reset_password():
    data = request.json
    admin = Admin.query.filter_by(email=data['email']).first()
    if not admin or admin.otp != data['otp']: return jsonify({'error': 'Invalid OTP'}), 400
    if datetime.datetime.now() > admin.otp_expiry: return jsonify({'error': 'OTP Expired'}), 400
    
    admin.password = data['new_password']
    admin.otp = None
    db.session.commit()
    return jsonify({'message': 'Password Changed'})

@app.route('/api/students', methods=['GET'])
def get_students():
    dept_filter = request.args.get('department')
    if dept_filter:
        students = Student.query.filter_by(department=dept_filter).all()
    else:
        students = Student.query.all()
    return jsonify([s.to_dict() for s in students])

@app.route('/api/students', methods=['POST'])
def add_or_update_student():
    data = request.json
    existing = Student.query.filter_by(enrollment_number=data['enrollment_number']).first()
    subjects = data.get('subjects', [])
    subjects_json = json.dumps(subjects)
    
    if existing:
        existing.student_name = data['student_name']
        existing.email = data.get('email', '')
        existing.password = data['password']
        existing.department = data['department'] 
        existing.course = data['course']
        existing.semester = data['semester']
        existing.subjects_json = subjects_json
        existing.attendance_marks = data.get('attendance_marks', 0)
        existing.internship_marks = data.get('internship_marks', 0)
        existing.project_marks = data.get('project_marks', 0)
        existing.timestamp = str(pd.Timestamp.now())
        msg = "Record Updated"
    else:
        if 'email' not in data: data['email'] = ""
        del data['subjects']
        data['subjects_json'] = subjects_json
        db.session.add(Student(**data))
        msg = "New Student Added"

    db.session.commit()
    sync_with_excel()
    return jsonify({'message': msg}), 201

@app.route('/api/students/<id>', methods=['DELETE'])
def delete_student(id):
    student = Student.query.get(id)
    if student:
        db.session.delete(student)
        db.session.commit()
        sync_with_excel()
    return jsonify({'message': 'Deleted Successfully'})

@app.route('/api/open-excel', methods=['GET'])
def open_excel():
    if not os.path.exists(EXCEL_FILE): return jsonify({'error': 'No file created yet'}), 404
    try:
        if platform.system() == "Windows": os.startfile(EXCEL_FILE)
        else: subprocess.call(["xdg-open", EXCEL_FILE])
        return jsonify({'message': 'Opening Excel...'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/api/upload-excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    dept_context = request.form.get('department') 

    try:
        df = pd.read_excel(file)
        count = 0
        for _, row in df.iterrows():
            enroll = str(row.get('Enrollment', '')).strip()
            if not enroll or enroll == 'nan': continue
            student = Student.query.filter_by(enrollment_number=enroll).first()
            row_dept = str(row.get('Department', dept_context)) 

            if not student:
                student = Student(id=str(pd.Timestamp.now().timestamp() + count), enrollment_number=enroll)
                db.session.add(student)
            
            student.student_name = str(row.get('Name', 'Unknown'))
            student.email = str(row.get('Email', ''))
            student.department = row_dept
            student.password = str(row.get('Password', enroll))
            student.course = str(row.get('Course', ''))
            student.semester = str(row.get('Semester', ''))
            student.attendance_marks = float(row.get('Attendance', 0))
            student.internship_marks = float(row.get('Internship', 0))
            student.project_marks = float(row.get('Project', 0))
            
            subj_str = str(row.get('Subjects', ''))
            subjects = []
            if subj_str and subj_str.lower() != 'nan':
                matches = re.findall(r"(.+?) \(E:(\d+(?:\.\d+)?)\+I:(\d+(?:\.\d+)?)\)", subj_str)
                for m in matches: subjects.append({'name': m[0].strip(), 'exam': float(m[1]), 'internal': float(m[2])})
            student.subjects_json = json.dumps(subjects)
            count += 1
        db.session.commit()
        sync_with_excel()
        return jsonify({'message': f'Processed {count} records'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/api/download-marksheet/<id>', methods=['GET'])
def download_marksheet(id):
    student = Student.query.get(id)
    if not student: return jsonify({'error': 'Student not found'}), 404
    
    try:
        html_content = get_marksheet_html(student)
        pdf_buffer = BytesIO()
        pisa.CreatePDF(html_content, dest=pdf_buffer)
        pdf_buffer.seek(0)
        
        return send_file(
            pdf_buffer,
            download_name=f"{student.student_name}_Marksheet.pdf",
            as_attachment=True,
            mimetype='application/pdf'
        )
    except Exception as e: return jsonify({'error': str(e)}), 500

# --- ROBUST & FAIL-SAFE EMAIL ROUTE ---
@app.route('/api/send-marksheet/<id>', methods=['POST'])
def send_marksheet(id):
    student = Student.query.get(id)
    if not student:
        return jsonify({'error': 'Student not found in DB'}), 404
    
    if not student.email or '@' not in student.email:
        return jsonify({'error': f'Invalid Email for student: {student.student_name}'}), 400
    
    # 1. HOD/Admin dhundo (Reply-To ke liye)
    # Agar Admin nahi mila, toh 'reply_to' bhi Master Email ban jayega. Error nahi aayega.
    admin = Admin.query.filter_by(department=student.department).first()
    reply_to_email = admin.email if admin else SYSTEM_EMAIL
    sender_display = f"JNCT {student.department}" if admin else "JNCT Admin"

    try:
        # PDF Generate
        html_content = get_marksheet_html(student)
        pdf_buffer = BytesIO()
        pisa.CreatePDF(html_content, dest=pdf_buffer)
        pdf_bytes = pdf_buffer.getvalue()
        
        # Email Draft
        msg = MIMEMultipart()
        msg['From'] = f"{sender_display} <{SYSTEM_EMAIL}>" # Master ID se jayega
        msg['Reply-To'] = reply_to_email # Reply HOD ko jayega
        msg['To'] = student.email
        msg['Subject'] = f"Result: {student.student_name} ({student.department})"
        
        msg.attach(MIMEText(f"Dear {student.student_name},\n\nPlease find your official marksheet attached for {student.course} ({student.department}).\n\nRegards,\nJNCT College", 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{student.student_name}_Marksheet.pdf"')
        msg.attach(part)
        
        # 2. Login & Send (Using MASTER CREDENTIALS)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SYSTEM_EMAIL, SYSTEM_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        return jsonify({'message': f'Email Sent Successfully to {student.email}!'})

    except smtplib.SMTPAuthenticationError:
        return jsonify({'error': 'Master Email Login Failed! Check App Password in app.py'}), 500
    except Exception as e:
        return jsonify({'error': f"Server Error: {str(e)}"}), 500

if __name__ == '__main__':
    with app.app_context():
        if os.path.exists(EXCEL_FILE): sync_with_excel()
    app.run(debug=True)