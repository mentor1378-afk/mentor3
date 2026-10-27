from dotenv import load_dotenv
load_dotenv()  # reads .env file automatically

import os
import json
import re
import io
import time
import urllib.request
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from werkzeug.utils import secure_filename
import cloudinary
import cloudinary.uploader
import cloudinary.api
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from datetime import datetime
import pytz
from pptx.oxml.xmlchemy import OxmlElement

# --- Dynamic model selection -------------------------------------------------

_GEMINI_SYSTEM_INSTRUCTION = (
    "You are a caring, professional college mentor writing an individualized student progress comment. "
    "Use the supplied reference style: begin with the student's name, describe attendance with exact subject "
    "names and percentages, acknowledge stronger subjects when relevant, address event participation and "
    "online-course status, and close with sincere encouragement. Use only supplied facts; never invent "
    "achievements, grades, risks, family circumstances, or actions. Do not copy private remarks verbatim. "
    "Write one natural paragraph without headings or bullets and return only a JSON object with "
    '"reg_no" and "comment" fields.'
)

def _get_model_choice():
    """Return the configured provider, model name, and API key."""
    gemini_key = (
        (os.getenv("GEMINI_API_KEY") or "").strip()
        or (os.getenv("GOOGLE_API_KEY") or "").strip()
    )
    openai_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    groq_key = (os.getenv("GROQ_API_KEY") or "").strip()
    choice = (os.getenv("MODEL_SELECTION") or "").strip().upper()

    if not choice:
        choice = (
            "GROQ" if groq_key else
            "GEMINI" if gemini_key else
            "OPENAI" if openai_key else
            "GEMINI"
        )

    if choice == "OPENAI":
        return ("openai", os.getenv("OPENAI_MODEL", "gpt-4o-mini"), openai_key or None)
    if choice == "GROQ":
        return (
            "groq",
            os.getenv("GROQ_MODEL", "openai/gpt-oss-120b"),
            groq_key or None,
        )
    if choice == "GEMINI":
        return (
            "gemini",
            os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
            gemini_key or None,
        )

    return (choice.lower(), "", None)

PROVIDER, MODEL_NAME, API_KEY = _get_model_choice()
AI_PROVIDER_AVAILABLE = False
_gemini_model = None
_chat_client = None

if PROVIDER == "gemini":
    if API_KEY:
        try:
            import google.generativeai as genai
            genai.configure(api_key=API_KEY)
            _gemini_model = genai.GenerativeModel(
                model_name=MODEL_NAME,
                system_instruction=_GEMINI_SYSTEM_INSTRUCTION,
            )
            AI_PROVIDER_AVAILABLE = True
        except Exception as exc:
            print(f"[AI Report] Gemini not available: {exc}")
            AI_PROVIDER_AVAILABLE = False
            _gemini_model = None
    else:
        print("[AI Report] Gemini not configured. Set GEMINI_API_KEY or GOOGLE_API_KEY to enable AI comment generation.")
elif PROVIDER in {"openai", "groq"}:
    if API_KEY:
        try:
            import openai
            client_options = {"api_key": API_KEY}
            if PROVIDER == "groq":
                client_options["base_url"] = "https://api.groq.com/openai/v1"
            _chat_client = openai.OpenAI(**client_options)
            AI_PROVIDER_AVAILABLE = True
        except Exception as exc:
            print(f"[AI Report] {PROVIDER.title()} not available: {exc}")
            AI_PROVIDER_AVAILABLE = False
            _chat_client = None
    else:
        api_key_name = "GROQ_API_KEY" if PROVIDER == "groq" else "OPENAI_API_KEY"
        print(f"[AI Report] {PROVIDER.title()} not configured. Set {api_key_name} to enable AI comment generation.")
else:
    print(f"[AI Report] Unsupported provider: {PROVIDER}")

# --- Document generation libraries ------------------------------------------------
try:
    from docx import Document
    from docx.shared import Pt as DocxPt, Inches as DocxInches, RGBColor as DocxRGB
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement as DocxOxmlElement
    DOCX_AVAILABLE = True
except Exception as _de:
    print(f"[AI Report] python-docx not available: {_de}")
    DOCX_AVAILABLE = False

def add_highlight(run, color_hex):
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement('a:highlight')
    srgbClr = OxmlElement('a:srgbClr')
    srgbClr.set('val', color_hex)
    highlight.append(srgbClr)
    rPr.append(highlight)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY') or os.urandom(32)
uri = os.environ.get("DATABASE_URL", "sqlite:///database.db")
if uri.startswith("postgres://"):
    uri = uri.replace("postgres://", "postgresql://", 1)
app.config['SQLALCHEMY_DATABASE_URI'] = uri

# Use QueuePool for better concurrency. NullPool opens a new connection per query,
# which causes connection exhaustion / 500 errors when many students use the app at once.
if 'postgresql' in uri or 'postgres' in uri:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_size': 10,
        'max_overflow': 20,
        'pool_timeout': 30,
        'pool_recycle': 280,
        'pool_pre_ping': True,
        'connect_args': {
            'sslmode': 'require',
            'connect_timeout': 10,
        },
    }
else:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True,
        'pool_recycle': 300,
    }
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)

# Models
class User(db.Model):
    username = db.Column(db.String(50), primary_key=True)
    password_hash = db.Column(db.String(100), nullable=False)
    role = db.Column(db.String(20), nullable=False) # 'faculty' or 'student'

class StudentDetail(db.Model):
    reg_num = db.Column(db.String(50), primary_key=True)
    name = db.Column(db.String(100))
    course = db.Column(db.String(100))
    # JSON strings for dynamic data
    slot_info = db.Column(db.Text, default='[]') # List of slot names
    attendance_data = db.Column(db.Text, default='{}') # {"Slot A": 81}
    marks_data = db.Column(db.Text, default='{}') # {"Slot A": {"model": 20}}
    
    registered_new_course = db.Column(db.String(200))
    online_course = db.Column(db.String(200))
    event_participation = db.Column(db.Text)
    additional_description = db.Column(db.Text)
    results_data = db.Column(db.Text, default='[]')  # [{"subject": "Maths", "grade": "A"}]
    custom_advisory = db.Column(db.Text)
    photo_path = db.Column(db.String(200))
    last_updated = db.Column(db.DateTime, nullable=True)

    def get_attendance(self):
        try: return json.loads(self.attendance_data)
        except: return {}
    
    def get_marks(self):
        try: return json.loads(self.marks_data)
        except: return {}
    
    def get_slots(self):
        try: return json.loads(self.slot_info)
        except: return []

class GlobalAdvisory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text, nullable=False, default="All students are advised to pay their 2nd-year tuition fees on time through the Viana Portal.\nAdditionally, kindly upload your recent passport-size photograph to your Viana profile at the earliest, if you have not already done so....")

class GlobalMentorObservation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text, nullable=False, default="I personally advised the student to concentrate more on study and skill development. The student is currently attending an online course to improve their technical skills, which is really appreciable.")

class GlobalSettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    show_grades_in_ppt = db.Column(db.Boolean, nullable=False, default=False)

# Helper functions
def delete_photo(photo_path):
    if not photo_path:
        return
    if 'cloudinary.com' in photo_path:
        try:
            public_id = "simats_profiles/" + photo_path.split('/')[-1].split('.')[0]
            cloudinary.uploader.destroy(public_id)
        except Exception as e:
            print(f'Failed to delete from Cloudinary: {e}')
    elif os.path.exists(photo_path):
        try:
            os.unlink(photo_path)
        except Exception as e:
            print(f'Failed to delete local photo: {e}')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def seed_db():
    # Rename the existing seeded faculty login without changing its password.
    previous_faculty = User.query.get('Ashwin@123')
    renamed_faculty = User.query.get('Gnanavel@123')
    if previous_faculty and not renamed_faculty:
        previous_faculty.username = 'Gnanavel@123'
    elif previous_faculty and renamed_faculty:
        print("[Account migration] Both faculty usernames exist; keeping both accounts unchanged.")
    elif not renamed_faculty:
        db.session.add(User(
            username='Gnanavel@123',
            password_hash=bcrypt.generate_password_hash(
                os.getenv('GNANAVEL_INITIAL_PASSWORD', 'welcome')
            ).decode('utf-8'),
            role='faculty'
        ))

    # Faculty: admin@123
    if not User.query.get('admin@123'):
        db.session.add(User(
            username='admin@123',
            password_hash=bcrypt.generate_password_hash(
                os.getenv('ADMIN_INITIAL_PASSWORD', 'welcome')
            ).decode('utf-8'),
            role='faculty'
        ))

    # Remove only the two known demo students previously inserted by seed_db.
    demo_student = StudentDetail.query.filter_by(
        reg_num='24REG01',
        name='Ibrahim',
        course='CSA0708 - Computer Networks',
    ).first()
    if demo_student:
        db.session.delete(demo_student)
        demo_account = User.query.filter_by(username='24REG01', role='student').first()
        if demo_account:
            db.session.delete(demo_account)

    demo_account = User.query.filter_by(username='Ibrahim@123', role='student').first()
    if demo_account and not StudentDetail.query.get('Ibrahim@123'):
        db.session.delete(demo_account)

    db.session.commit()

with app.app_context():
    db.create_all()  # Only creates tables if they don't exist — never wipes data
    try:
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        columns = [col['name'] for col in inspector.get_columns('student_detail')]

        def add_column_if_missing(column_name, ddl):
            if column_name not in columns:
                db.session.execute(db.text(f'ALTER TABLE student_detail ADD COLUMN {column_name} {ddl}'))
                db.session.commit()
                columns.append(column_name)

        add_column_if_missing('last_updated', 'DATETIME' if db.engine.name != 'postgresql' else 'TIMESTAMP')
        add_column_if_missing('results_data', "TEXT DEFAULT '[]'")
        add_column_if_missing('custom_advisory', 'TEXT')
        add_column_if_missing('mentor_username', 'VARCHAR(50)')
    except Exception as e:
        db.session.rollback()
        print(f"Migration error: {e}")
    seed_db()

# Routes
@app.route('/')
def index():
    if 'user' in session:
        if session['role'] == 'faculty':
            return redirect(url_for('faculty_dashboard'))
        return redirect(url_for('student_dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        
        if user and bcrypt.check_password_hash(user.password_hash, password):
            session['user'] = user.username
            session['role'] = user.role
            if user.role == 'faculty':
                return redirect(url_for('faculty_dashboard'))
            return redirect(url_for('student_dashboard'))
        flash('Invalid username or password', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/change_password', methods=['POST'])
def change_password():
    if 'user' not in session:
        return redirect(url_for('login'))
    
    current_password = request.form['current_password']
    new_password = request.form['new_password']
    
    user = User.query.get(session['user'])
    if user and bcrypt.check_password_hash(user.password_hash, current_password):
        user.password_hash = bcrypt.generate_password_hash(new_password).decode('utf-8')
        db.session.commit()
        flash('Password changed successfully!', 'success')
    else:
        flash('Incorrect current password', 'danger')
    
    return redirect(url_for('faculty_dashboard' if session['role'] == 'faculty' else 'student_dashboard'))

@app.route('/student', methods=['GET', 'POST'])
def student_dashboard():
    if 'user' not in session or session['role'] != 'student':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.get(session['user'])
    if not student:
        student = StudentDetail(reg_num=session['user'], name=session['user'])
        db.session.add(student)
        db.session.commit()

    if request.method == 'POST':
        try:
            student.name = (request.form.get('name') or '').strip() or student.name
            student.course = (request.form.get('course') or '').strip() or student.course

            # Dynamic Slots handling
            slots = [s.strip() for s in request.form.getlist('slot_names[]') if s and s.strip()]
            if not slots:
                slots = student.get_slots() or []
            student.slot_info = json.dumps(slots)

            att_data = {}
            marks_data = {}
            current_attendance = student.get_attendance()
            current_marks = student.get_marks()
            for slot in slots:
                att_val = request.form.get(f'att_{slot}', current_attendance.get(slot, 0))
                att_data[slot] = str(att_val).strip() if att_val is not None else 0
                marks_data[slot] = {
                    'model': (request.form.get(f'model_{slot}') or current_marks.get(slot, {}).get('model', '-')).strip(),
                    'test1': (request.form.get(f'test1_{slot}') or current_marks.get(slot, {}).get('test1', '-')).strip(),
                    'test2': (request.form.get(f'test2_{slot}') or current_marks.get(slot, {}).get('test2', '-')).strip(),
                    'avg': (request.form.get(f'avg_{slot}') or current_marks.get(slot, {}).get('avg', '-')).strip(),
                    'total_marks': (request.form.get(f'total_marks_{slot}') or current_marks.get(slot, {}).get('total_marks', '')).strip(),
                    'course': (request.form.get(f'course_{slot}') or current_marks.get(slot, {}).get('course', '')).strip()
                }

            student.attendance_data = json.dumps(att_data)
            student.marks_data = json.dumps(marks_data)

            student.registered_new_course = (request.form.get('registered_new_course') or '').strip()
            student.online_course = (request.form.get('online_course') or '').strip()
            student.event_participation = (request.form.get('event_participation') or '').strip()
            student.additional_description = (request.form.get('description') or '').strip()

            # Results data — save subject/grade entries from My Results tab
            result_subjects = request.form.getlist('result_subject[]')
            result_grades = request.form.getlist('result_grade[]')
            if result_subjects:  # Only update if results fields were submitted
                results = [{'subject': s.strip(), 'grade': g.strip()} for s, g in zip(result_subjects, result_grades) if s and s.strip()]
                student.results_data = json.dumps(results)

            file = request.files.get('photo')
            if file and file.filename and allowed_file(file.filename):
                if os.environ.get('CLOUDINARY_URL'):
                    try:
                        upload_result = cloudinary.uploader.upload(file, folder="simats_profiles")
                        student.photo_path = upload_result.get('secure_url')
                    except Exception as e:
                        flash(f'Failed to upload photo to cloud storage: {str(e)}', 'warning')
                else:
                    try:
                        filename = secure_filename(f"{session['user']}_{file.filename}")
                        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                        file.save(filepath)
                        student.photo_path = filepath
                    except Exception as e:
                        print(f"Local photo upload warning: {e}")

            # Safe Kolkata timestamp calculation (pure Python, immune to OS pytz missing tzdata)
            from datetime import timedelta
            student.last_updated = datetime.utcnow() + timedelta(hours=5, minutes=30)
            
            db.session.commit()
            flash('Details updated successfully!', 'success')
        except Exception as e:
            db.session.rollback()
            print(f"Student submit error: {e}")
            flash('An error occurred while saving your details. Please try again.', 'danger')

        return redirect(url_for('student_dashboard'))
    
    # Passing current data as dicts
    try:
        results = json.loads(student.results_data) if student.results_data else []
    except:
        results = []
    return render_template('student.html', student=student,
                           attendance=student.get_attendance(),
                           marks=student.get_marks(),
                           slots=student.get_slots(),
                           results=results)

@app.route('/add_student', methods=['POST'])
def add_student():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    reg_num = request.form['reg_num']
    name = request.form['name']
    password = request.form['password']
    
    if User.query.get(reg_num):
        flash('Student with this Registration Number already exists.', 'danger')
    else:
        new_user = User(
            username=reg_num,
            password_hash=bcrypt.generate_password_hash(password).decode('utf-8'),
            role='student'
        )
        new_detail = StudentDetail(reg_num=reg_num, name=name, last_updated=datetime.now(pytz.timezone('Asia/Kolkata')).replace(tzinfo=None))
        db.session.add(new_user)
        db.session.add(new_detail)
        db.session.commit()
        flash(f'Student {name} added successfully!', 'success')
    
    return redirect(url_for('faculty_dashboard'))

@app.route('/edit_student_faculty/<reg_num>', methods=['POST'])
def edit_student_faculty(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.get(reg_num)
    if student:
        student.name = request.form.get('name', student.name) or student.name
        student.course = request.form.get('course', student.course) or student.course
        student.registered_new_course = request.form.get('registered_new_course', student.registered_new_course or '')
        student.online_course = request.form.get('online_course', student.online_course or '')
        student.event_participation = request.form.get('event_participation', student.event_participation or '')
        student.additional_description = request.form.get('description', student.additional_description or '')
        student.custom_advisory = request.form.get('custom_advisory', student.custom_advisory or '')

        # Save results from faculty edit modal
        result_subjects = request.form.getlist('result_subject[]')
        result_grades = request.form.getlist('result_grade[]')
        results = [{'subject': s.strip(), 'grade': g.strip()} for s, g in zip(result_subjects, result_grades) if s.strip()]
        student.results_data = json.dumps(results)

        student.last_updated = datetime.now(pytz.timezone('Asia/Kolkata')).replace(tzinfo=None)
        db.session.commit()
        flash(f'Details for {student.name or reg_num} updated successfully!', 'success')
    else:
        flash('Student not found.', 'danger')
        
    return redirect(url_for('faculty_dashboard'))

@app.route('/update_global_advisory', methods=['POST'])
def update_global_advisory():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    advisory_text = request.form.get('global_advisory', '').strip()
    global_adv = GlobalAdvisory.query.first()
    if not global_adv:
        global_adv = GlobalAdvisory(id=1, content=advisory_text)
        db.session.add(global_adv)
    else:
        global_adv.content = advisory_text
    
    db.session.commit()
    flash('Global Institutional Advisory updated for all students!', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/update_global_observation', methods=['POST'])
def update_global_observation():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    observation_text = request.form.get('global_observation', '').strip()
    global_obs = GlobalMentorObservation.query.first()
    if not global_obs:
        global_obs = GlobalMentorObservation(id=1, content=observation_text)
        db.session.add(global_obs)
    else:
        global_obs.content = observation_text
    
    db.session.commit()
    flash('Global Mentor Observation updated for all students!', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/faculty')
def faculty_dashboard():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    students = StudentDetail.query.all()
    global_adv = GlobalAdvisory.query.first()
    global_advisory_text = global_adv.content if global_adv else ""
    global_obs = GlobalMentorObservation.query.first()
    global_observation_text = global_obs.content if global_obs else ""
    
    # Simplified stats for global update
    total_att_a = 0
    count_a = 0
    for s in students:
        att = s.get_attendance()
        if 'Slot A' in att:
            total_att_a += int(att['Slot A'] or 0)
            count_a += 1
    
    stats = {
        'total_students': len(students),
        'reports_generated': 152,
        'avg_attendance_a': int(total_att_a / count_a) if count_a > 0 else 0,
        'avg_attendance_b': 92
    }

    students_json = [
        {
            'reg_num': s.reg_num,
            'name': s.name,
            'course': s.course,
            'attendance_data': s.attendance_data,
            'marks_data': s.marks_data,
            'slot_info': s.slot_info,
            'registered_new_course': s.registered_new_course,
            'online_course': s.online_course,
            'event_participation': s.event_participation,
            'additional_description': s.additional_description,
            'custom_advisory': s.custom_advisory,
            'photo_path': s.photo_path,
            'last_updated': s.last_updated.strftime('%d-%b-%Y %I:%M %p') if s.last_updated else 'N/A',
        }
        for s in students
    ]

    return render_template('faculty.html', students=students, students_json=students_json, stats=stats,
                           global_advisory=global_advisory_text,
                           global_observation=global_observation_text,
                           show_grades_in_ppt=(GlobalSettings.query.first().show_grades_in_ppt if GlobalSettings.query.first() else False))

@app.route('/toggle_grade_ppt', methods=['POST'])
def toggle_grade_ppt():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    settings = GlobalSettings.query.first()
    if not settings:
        settings = GlobalSettings(id=1, show_grades_in_ppt=True)
        db.session.add(settings)
    else:
        settings.show_grades_in_ppt = not settings.show_grades_in_ppt
    db.session.commit()
    return redirect(url_for('faculty_dashboard'))

@app.route('/generate_report')
def generate_report():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    students = StudentDetail.query.all()
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    for student in students:
        slots = student.get_slots()
        att_data = student.get_attendance()
        marks_data = student.get_marks()
        mentor_name = "GNANAVEL V K"
        
        # --- SLIDE 1: ACADEMIC PERFORMANCE ---
        slide_layout = prs.slide_layouts[6] # Blank
        slide1 = prs.slides.add_slide(slide_layout)
        
        # Header Box removed as per request

        # Banner Table for Name/Reg/Mentor (replacing the previous textbox for better alignment)
        banner_tbl = slide1.shapes.add_table(1, 6, Inches(0.2), Inches(1.4), Inches(13.0), Inches(0.5)).table
        banner_content = [
            "Mentee Name", student.name or 'N/A',
            "Reg.NO", student.reg_num or 'N/A',
            "Mentor name", mentor_name
        ]
        
        for i, text in enumerate(banner_content):
            cell = banner_tbl.cell(0, i)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.name = 'Times New Roman'
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.LEFT if i % 2 == 0 else PP_ALIGN.CENTER

        # Student Photo with gray border (matching image 1)
        left_img = Inches(0.5)
        top_img = Inches(2.2)
        width_img = Inches(2.5)
        height_img = Inches(3.0)
        
        # Draw a gray frame
        frame = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_img - Inches(0.1), top_img - Inches(0.1), width_img + Inches(0.2), height_img + Inches(0.2))
        frame.fill.solid()
        frame.fill.fore_color.rgb = RGBColor(230, 230, 230)
        frame.line.color.rgb = RGBColor(200, 200, 200)

        photo_added = False
        try:
            if student.photo_path:
                if student.photo_path.startswith('http'):
                    req = urllib.request.Request(student.photo_path, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req) as response:
                        image_stream = BytesIO(response.read())
                    slide1.shapes.add_picture(image_stream, left_img, top_img, width=width_img, height=height_img)
                    photo_added = True
                elif os.path.exists(student.photo_path):
                    slide1.shapes.add_picture(student.photo_path, left_img, top_img, width=width_img, height=height_img)
                    photo_added = True
        except Exception as e:
            print(f"Error adding photo to PPT: {e}")

        if not photo_added:
            rect = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_img, top_img, width_img, height_img)
            rect.fill.solid()
            rect.fill.fore_color.rgb = RGBColor(100, 100, 100)
            text_frame = rect.text_frame
            text_frame.clear()
            p = text_frame.paragraphs[0]
            p.text = "No Photo"
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        # --- Dynamic Marks Table: one row per slot ---
        all_slots = slots if slots else ["Slot A"]
        n_data_rows = len(all_slots)              # 1 row per slot
        total_rows  = 1 + n_data_rows             # + 1 header row
        cols = 4
        table_width  = Inches(9.5)
        # Scale height: 0.55 in per data row, min 2.0 in
        table_height = max(Inches(2.0), Inches(0.55 + n_data_rows * 0.55))
        left_tbl = Inches(3.3)
        top_tbl  = Inches(2.2)

        table = slide1.shapes.add_table(total_rows, cols, left_tbl, top_tbl, table_width, table_height).table

        # Header Row
        h_labels = ["Exam", "Total Marks", "Marks Obtained", "Class Average Mark"]
        for i, h in enumerate(h_labels):
            cell = table.cell(0, i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(112, 173, 71)   # SIMATS Green
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True
            p.font.size = Pt(18)
            p.alignment = PP_ALIGN.CENTER

        # Data rows – one row per slot (Test 1 only)
        for row_idx, slot in enumerate(all_slots, start=1):
            s_marks      = marks_data.get(slot, {})
            test1_val    = str(s_marks.get('test1', '0') or '0')
            avg_val      = str(s_marks.get('avg',   '0') or '0')
            total_m_val  = str(s_marks.get('total_marks', '-') or '-')

            r_vals = [f"Test 1 ({slot})", total_m_val, test1_val, avg_val]
            for c_idx, val in enumerate(r_vals):
                cell = table.cell(row_idx, c_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(226, 239, 218)   # Light green
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(18)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER

        # Footer Date
        # Removed as per request

        # --- SLIDE 2: MENTOR NOTES & ATTENDANCE ---
        slide2 = prs.slides.add_slide(slide_layout)

        # Header Box removed as per request

        # Read grade toggle setting
        ppt_settings = GlobalSettings.query.first()
        show_grades = ppt_settings.show_grades_in_ppt if ppt_settings else False

        # Main Gray Content Box
        body_box = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.8))
        body_box.fill.solid()
        body_box.fill.fore_color.rgb = RGBColor(245, 245, 245)
        body_box.line.color.rgb = RGBColor(200, 200, 200)

        tf_body = slide2.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12.1), Inches(5.6)).text_frame
        tf_body.word_wrap = True

        # "Welcome to SIMATS ENGINEERING" with green highlight
        p = tf_body.paragraphs[0]
        p.text = "Welcome to SIMATS ENGINEERING"
        p.font.bold = True
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(0, 0, 0)
        if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        p = tf_body.add_paragraph()
        p.text = "Dear Parent,"
        p.font.size = Pt(18)
        p.font.bold = True
        p.space_after = Pt(10)

        p = tf_body.add_paragraph()

        low_attendance = False
        for slot in slots:
            try:
                if int(att_data.get(slot, 0)) < 80:
                    low_attendance = True
                    break
            except (ValueError, TypeError):
                pass

        if low_attendance:
            p.text = f"{student.name or 'The student'} has attendance below 80%. Please maintain the attendance % above 80%."
            p.font.color.rgb = RGBColor(0, 0, 0)
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FF0000')
        else:
            p.text = f"So far {student.name or 'the student'} has maintained consistent attendance in the course."
            p.font.color.rgb = RGBColor(0, 0, 0)
            if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        p.font.size = Pt(18)
        p.font.bold = True

        for slot in slots:
            p = tf_body.add_paragraph()
            course_name = marks_data.get(slot, {}).get('course', '')
            if course_name:
                p.text = f"Attendance for {slot}: {course_name}: {att_data.get(slot, 0)}%"
            else:
                p.text = f"Attendance for {slot}: {att_data.get(slot, 0)}%"
            p.font.size = Pt(18)
            p.font.bold = True
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        # Use student-specific observation if set, otherwise fall back to global observation
        global_obs_record = GlobalMentorObservation.query.first()
        global_obs_text = global_obs_record.content if global_obs_record else 'I personally advised the student to concentrate more on study and skill development.'
        observation_to_use = student.additional_description.strip() if (student.additional_description and student.additional_description.strip()) else global_obs_text
        p = tf_body.add_paragraph()
        p.space_before = Pt(15)
        p.text = observation_to_use
        p.font.size = Pt(18)
        p.font.bold = True
        if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        if show_grades:
            # --- GRADES MODE: show subject results, skip new course & event ---
            try:
                results = json.loads(student.results_data) if student.results_data else []
            except:
                results = []
            if results:
                p = tf_body.add_paragraph()
                p.text = "Subject Results:"
                p.font.size = Pt(18)
                p.font.bold = True
                p.space_before = Pt(10)
                if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')
                for r in results:
                    p = tf_body.add_paragraph()
                    p.text = f"  {r.get('subject', '')}  —  Grade: {r.get('grade', '')}"
                    p.font.size = Pt(18)
                    p.font.bold = True
                    if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')
            else:
                p = tf_body.add_paragraph()
                p.text = "No subject results added yet."
                p.font.size = Pt(18)
                p.font.bold = True
                p.space_before = Pt(10)
        else:
            # --- STANDARD MODE: show new course & event, skip grades ---
            p = tf_body.add_paragraph()
            p.text = f"New course: {student.registered_new_course or 'N/A'}"
            p.font.size = Pt(18)
            p.space_before = Pt(10)
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

            p = tf_body.add_paragraph()
            p.space_before = Pt(15)
            if student.event_participation and student.event_participation.strip():
                p.text = f"Your ward participated in: {student.event_participation.strip()} and gave his very best throughout the journey. His dedication, hard work, and sincere efforts are truly appreciable."
            else:
                p.text = "We encourage your ward to actively participate in upcoming events and extracurricular activities to build their skills and gain valuable experience."
            p.font.size = Pt(18)
            p.font.bold = True
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        # Add green institutional advisory lines (Global or Custom Override)
        advisory_to_use = student.custom_advisory.strip() if (student.custom_advisory and student.custom_advisory.strip()) else (GlobalAdvisory.query.first().content if GlobalAdvisory.query.first() else "All students are advised to pay their 2nd-year tuition fees on time through the Viana Portal.\nAdditionally, kindly upload your recent passport-size photograph to your Viana profile at the earliest...")

        for adv_line in advisory_to_use.split('\n'):
            if adv_line.strip():
                p = tf_body.add_paragraph()
                p.text = adv_line.strip()
                p.font.size = Pt(18)
                p.font.bold = True
                if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        # Apply Times New Roman font to all paragraphs
        for paragraph in tf_body.paragraphs:
            paragraph.font.name = 'Times New Roman'
            if len(paragraph.runs) > 0:
                paragraph.runs[0].font.name = 'Times New Roman'

        # Footer Date
        # Removed as per request

    report_path = 'Mentor_Dashboard_Report.pptx'
    prs.save(report_path)
    return send_file(report_path, as_attachment=True)

@app.route('/clear_student/<reg_num>', methods=['POST'])
def clear_student(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.get(reg_num)
    if student:
        student.slot_info = '[]'
        student.attendance_data = '{}'
        student.marks_data = '{}'
        student.registered_new_course = None
        student.online_course = None
        student.event_participation = None
        student.additional_description = None
        delete_photo(student.photo_path)
        student.photo_path = None
        student.last_updated = datetime.now(pytz.timezone('Asia/Kolkata')).replace(tzinfo=None)
        db.session.commit()
        flash(f'Report data for {student.name or reg_num} has been cleared.', 'success')
    else:
        flash('Student not found.', 'danger')
    
    return redirect(url_for('faculty_dashboard'))

@app.route('/remove_student/<reg_num>', methods=['POST'])
def remove_student(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.get(reg_num)
    if student:
        delete_photo(student.photo_path)
        db.session.delete(student)
    
    # Delete the student's login account
    user = User.query.get(reg_num)
    if user:
        db.session.delete(user)
    
    db.session.commit()
    flash(f'Student {reg_num} has been permanently removed.', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/delete_all_reports', methods=['POST'])
def delete_all_reports():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    # Only clear PPT/report fields — keep student accounts and rows intact
    students = StudentDetail.query.all()
    for student in students:
        delete_photo(student.photo_path)
        student.slot_info = '[]'
        student.attendance_data = '{}'
        student.marks_data = '{}'
        student.registered_new_course = None
        student.online_course = None
        student.event_participation = None
        student.additional_description = None
        student.photo_path = None

    db.session.commit()
    flash('All student report data has been cleared. Accounts are still active.', 'success')
    return redirect(url_for('faculty_dashboard'))


# ===========================================================================
# AI MENTOR COMMENT REPORT
# ===========================================================================

def _get_grades(student):
    """Return a list like [{'subject': 'Math', 'grade': 'A'}, ...] from student.results_data."""
    try:
        results = json.loads(student.results_data or "[]")
        return [
            {"subject": r.get("subject", "N/A"), "grade": r.get("grade", "N/A")}
            for r in results
        ]
    except Exception:
        return []

def correct_student_text(text):
    """
    Lightweight correction of common student-side typos:
    - Remove duplicate consecutive words (e.g. "Cryptography Cryptography")
    - Strip extra whitespace/symbols
    - Title-case names that are all-lowercase (heuristic)
    Returns (corrected_text, was_changed: bool)
    """
    if not text:
        return text, False

    original = text.strip()

    # Remove duplicate consecutive words (case-insensitive)
    corrected = re.sub(r'\b(\w+)\s+\1\b', r'\1', original, flags=re.IGNORECASE)

    # Collapse multiple spaces
    corrected = re.sub(r' {2,}', ' ', corrected).strip()

    # Title-case if the whole string is lowercase (names, subjects)
    if corrected == corrected.lower() and corrected:
        corrected = corrected.title()

    # Fix 5g/6g -> 5G/6G style (digit+letter combinations should be uppercase)
    def uppercase_tech(m):
        return m.group(0).upper()
    corrected = re.sub(r'\b(\d+[a-zA-Z]+|\b(?:ai|ml|iot|nlp|dsa|dbms|os|cn|se|pdsd|coa|psp|oop|oops|tcp|ip|rdbms|sql|html|css|js|php|ui|ux|api|sdk|ide|lms|erp|crm|gps|ble|rfid|otp|sms|cse|it|ece|eee|be|btech|mtech|mba|mca)\b)', uppercase_tech, corrected, flags=re.IGNORECASE)

    changed = corrected != original
    return corrected, changed




def _build_ai_payload(student, show_grades=False):
    """
    Compute facts in pure Python for one student.
    show_grades must be passed from the main thread (not queried here)
    to avoid SQLAlchemy thread-safety issues.
    Returns (payload_dict, corrections_list)
    """
    att_data = student.get_attendance()
    marks_data = student.get_marks()
    corrections = []

    # --- Correct student-entered text fields ---
    def maybe_correct(field_value, field_name):
        corrected, changed = correct_student_text(field_value or '')
        if changed and field_value:
            corrections.append({
                'field': field_name,
                'original': field_value,
                'corrected': corrected
            })
        return corrected if corrected else (field_value or '')

    name = maybe_correct(student.name, 'Name')
    event = maybe_correct(student.event_participation, 'Event Participation')
    online_course = maybe_correct(student.online_course, 'Online Course')
    remark = maybe_correct(student.additional_description, 'Remark / Observation')

    # --- Compute attendance facts ---
    low_subjects = []      # below 80%
    borderline_subjects = []  # exactly 80%
    good_count = 0

    for slot, att_val in att_data.items():
        try:
            att_pct = float(str(att_val).replace('%', '').strip() or 0)
        except (ValueError, TypeError):
            att_pct = 0

        # Get subject/course name for this slot from marks_data
        subject_name = marks_data.get(slot, {}).get('course', '') or slot

        if att_pct < 80:
            low_subjects.append({'subject': subject_name, 'percent': att_pct})
        elif att_pct == 80:
            borderline_subjects.append(subject_name)
        else:
            good_count += 1

    other_slots_good = good_count > 0 and len(low_subjects) > 0

    # Infer gender from name (simple heuristic for Indian names)
    lower_name = (student.name or '').lower().split()[0] if student.name else ''
    female_endings = ('a', 'i', 'ee', 'ya', 'na', 'ra', 'la', 'ita', 'ika', 'ika', 'eeka')
    gender = 'female' if any(lower_name.endswith(e) for e in female_endings) else 'male'

    payload = {
        'reg_no': student.reg_num,
        'name': name,
        'gender': gender,
        'low_subjects': low_subjects,
        'borderline_subjects': borderline_subjects,
        'other_slots_good': other_slots_good,
        'event': event,
        'online_course': online_course,
        'student_remark': remark,
        'institution': 'SIMATS Engineering',
    }

    # Include grades only when the flag is explicitly passed as True
    # (flag must be read in the main thread before spawning worker threads)
    if show_grades:
        payload['grades'] = _get_grades(student)

    return payload, corrections


def _call_model(payload):
    """Unified LLM call – works with Gemini, OpenAI, or Groq.
    Builds the prompt (including grades & remark) and returns the comment.
    """
    # ---- Build extra instructions ------------------------------------------------
    extra_parts = []
    if payload.get('grades'):
        grades_txt = ', '.join(f"{g['subject']} ({g['grade']})" for g in payload['grades'])
        extra_parts.append(
            f"The student has the following grades that MUST be woven naturally into the comment: {grades_txt}."
        )
    remark = (payload.get('student_remark') or '').strip()
    if remark:
        extra_parts.append(
            f"The student has shared this personal remark/reason (e.g. for low attendance): '{remark}'. "
            f"Acknowledge it briefly and sensitively in the comment — do NOT copy it word for word, "
            f"just mention that the student has explained the situation and encourage them going forward."
        )
    extra = ("\n" + " ".join(extra_parts)) if extra_parts else ""
    prompt = f"Input: {json.dumps(payload, ensure_ascii=False)}{extra}"

    # ---- Dispatch ---------------------------------------------------------------
    if PROVIDER == "gemini":
        response = _gemini_model.generate_content(prompt)
        raw = response.text.strip()
    elif PROVIDER in {"openai", "groq"}:
        system_msg = (
            "You are a college mentor personally writing a short comment about one student for the semester mentor report. "
            "Follow this reference style: start with the student's name; state attendance accurately with subject names "
            "and percentages; recognize subjects with stronger attendance when relevant; describe any recorded event "
            "participation accurately; mention online-course status when provided; and end with specific, warm "
            "encouragement. Write one natural, personal paragraph of about 90 to 150 words.\n"
            "Rules:\n"
            "1. Mention low attendance only for subjects below 80%, with exact subject names and percentages. "
            "If multiple subjects are low, identify each and recommend attending upcoming classes. Do not claim "
            "detention risk, parent meetings, or other consequences unless explicitly provided.\n"
            "2. If there are no low-attendance subjects, say attendance is consistent; if some subjects are good, "
            "acknowledge them without implying that all are good.\n"
            "3. If event details are supplied, describe only the recorded participation or achievement. If none are "
            "supplied, say no event participation was recorded and encourage a suitable future opportunity.\n"
            "4. Mention online-course participation only when the supplied facts support it; if blank, say no course "
            "was recorded and encourage one where appropriate.\n"
            "5. Weave provided grades naturally. Use only the supplied facts; do not invent subjects, percentages, "
            "achievements, course enrollment, or personal circumstances. Do not repeat private remarks verbatim.\n"
            "6. Use a sincere mentor voice, vary sentence structure, and close with encouragement. Output no headings, "
            "bullets, or quotation marks. Silently correct spelling and grammar.\n"
            "Return ONLY a JSON object: {\"reg_no\": \"...\", \"comment\": \"...\"}"
        )
        resp = _chat_client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
        raw = resp.choices[0].message.content.strip()
    else:
        raise RuntimeError(f"Unsupported AI provider: {PROVIDER}")

    # ---- Clean up possible markdown fences --------------------------------------
    raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.DOTALL)
    raw = re.sub(r'\s*```$', '', raw, flags=re.DOTALL)

    # ---- Try to parse JSON; fall back to raw text if parsing fails -----------
    try:
        result = json.loads(raw)
        return result.get('comment', '').strip()
    except Exception:
        return raw


def fallback_comment(payload):
    """Fixed template used when AI generation fails or is unavailable."""
    name = payload.get('name', 'The student')
    low = payload.get('low_subjects', [])
    event = payload.get('event', '')
    online = payload.get('online_course', '')
    remark = (payload.get('student_remark') or '').strip()
    gender_pronoun = 'she' if payload.get('gender') == 'female' else 'he'
    gender_pos = 'her' if payload.get('gender') == 'female' else 'his'
    first = name.split()[0]

    if low:
        att_str = ', '.join(f"{s['subject']} ({int(s['percent'])}%)" for s in low)
        att_sentence = (f"{name} has recorded attendance below 80% in the following subject(s): "
                        f"{att_str}. I encourage {first} to attend every upcoming class "
                        f"to bring {gender_pos} attendance above 80%.")
    else:
        att_sentence = (f"{name} has maintained consistent attendance across all subjects, "
                        f"which reflects {gender_pos} commitment and dedication.")

    remark_sentence = ''
    if remark:
        remark_sentence = (f" I have noted {first}'s personal explanation regarding this, "
                           f"and I appreciate {gender_pos} transparency. "
                           f"I encourage {first} to manage {gender_pos} responsibilities proactively going forward.")

    event_sentence = ''
    if event:
        event_sentence = f" {first} participated in {event}, which is highly appreciable."
    else:
        event_sentence = (f" I encourage {first} to participate in upcoming hackathons, "
                          f"symposiums, or workshops to further develop {gender_pos} skills.")

    course_sentence = ''
    if online:
        course_sentence = (f" {first} is also pursuing an online course in {online}, "
                           f"which shows {gender_pos} eagerness to learn beyond the classroom.")

    # Include grades sentence when grades are in payload
    grades_sentence = ''
    if payload.get('grades'):
        grades_list = ', '.join(f"{g['subject']} ({g['grade']})" for g in payload['grades'])
        grades_sentence = f" {first} has achieved the following grades: {grades_list}."

    closing = (f" Keep it up, {first}! "
               f"I believe in {gender_pos} potential and look forward to seeing "
               f"{gender_pronoun} achieve great success.")

    return f"{att_sentence}{remark_sentence}{event_sentence}{course_sentence}{grades_sentence}{closing}".strip()


def _validate_ai_comment(comment, payload):
    if not comment:
        raise ValueError("The AI provider returned an empty comment.")
    for subject in payload.get('low_subjects', []):
        name = subject.get('subject', '')
        percent = subject.get('percent', -1)
        if name and name.lower() not in comment.lower():
            raise ValueError(f"Comment missing subject: {name}")
        if percent >= 0 and str(int(percent)) not in comment:
            raise ValueError(f"Comment missing percentage for: {name}")


def _is_rate_limit_error(error):
    message = str(error).lower()
    return '429' in message or 'quota' in message or 'rate limit' in message


def _rate_limit_retry_delay(error):
    match = re.search(r'please retry in\s+([0-9]+(?:\.[0-9]+)?)\s*seconds?', str(error), re.IGNORECASE)
    return float(match.group(1)) if match else None


def _generate_for_one(student, show_grades=False):
    """Generate comment for a single student. Returns result dict."""
    payload, corrections = _build_ai_payload(student, show_grades=show_grades)

    comment = None
    ai_generated = False
    ai_warning = None
    if AI_PROVIDER_AVAILABLE:
        try:
            comment = _call_model(payload)
            _validate_ai_comment(comment, payload)
            ai_generated = bool(comment)
        except Exception as first_err:
            print(f"[AI Report] First attempt failed for {student.reg_num}: {first_err}")
            if _is_rate_limit_error(first_err):
                retry_delay = _rate_limit_retry_delay(first_err)
                if retry_delay is not None and retry_delay <= 5:
                    time.sleep(retry_delay + 0.25)
                    try:
                        comment = _call_model(payload)
                        _validate_ai_comment(comment, payload)
                        ai_generated = bool(comment)
                    except Exception as retry_err:
                        print(f"[AI Report] Retry failed for {student.reg_num}: {retry_err}")
                        comment = None
                if not comment:
                    ai_warning = (
                        f"{PROVIDER.title()}'s request limit was reached; a template comment was used. "
                        "Wait for the quota window to reset, then try again."
                    )
            else:
                # Retry once for errors other than provider quota limits.
                try:
                    comment = _call_model(payload)
                    _validate_ai_comment(comment, payload)
                    ai_generated = bool(comment)
                except Exception as retry_err:
                    print(f"[AI Report] Retry failed for {student.reg_num}: {retry_err}")
                    comment = None
                    if _is_rate_limit_error(retry_err):
                        ai_warning = (
                            f"{PROVIDER.title()}'s request limit was reached; a template comment was used. "
                            "Wait for the quota window to reset, then try again."
                        )
                    else:
                        ai_warning = "AI generation failed; a template comment was used. Check the server logs and API configuration."
    else:
        ai_warning = f"{PROVIDER} is unavailable. Check the configured API key and server dependencies."

    if not comment:
        if not ai_warning:
            ai_warning = "AI returned no comment; a template comment was used. Check the server logs and API configuration."
        comment = fallback_comment(payload)
        ai_generated = False

    return {
        'reg_no': student.reg_num,
        'name': payload['name'],
        'comment': comment,
        'ai_generated': ai_generated,
        'ai_warning': ai_warning,
        'corrections': corrections,
        'gender': payload['gender'],
        'low_count': len(payload['low_subjects']),
        'has_event': bool(payload.get('event', '').strip()),
    }


@app.route('/generate_comment_report', methods=['POST'])
def generate_comment_report():
    """
    POST /generate_comment_report
    Accepts optional JSON body { "reg_nos": ["24REG01"] } to regenerate a subset.
    Returns JSON list of { reg_no, name, comment, corrections }.
    """
    if 'user' not in session or session['role'] != 'faculty':
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.get_json(silent=True) or {}
    reg_nos = data.get('reg_nos')  # None means all students

    if reg_nos:
        students = [StudentDetail.query.get(r) for r in reg_nos if StudentDetail.query.get(r)]
    else:
        students = StudentDetail.query.all()

    if not students:
        return jsonify({'error': 'No students found'}), 404

    # Read the grades toggle ONCE in the main thread (SQLAlchemy is not thread-safe)
    try:
        settings = GlobalSettings.query.first()
        show_grades = bool(settings and settings.show_grades_in_ppt)
    except Exception:
        show_grades = False

    results = []
    # Keep Gemini calls sequential; OpenAI-compatible providers can use a small parallel batch.
    BATCH = 5
    # Serialize Gemini requests to avoid exceeding low free-tier per-minute quotas.
    max_workers = 1 if PROVIDER == "gemini" else BATCH
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_generate_for_one, s, show_grades): s for s in students}
        for future in as_completed(futures):
            try:
                results.append(future.result())
            except Exception as e:
                s = futures[future]
                print(f"[AI Report] Exception for {s.reg_num}: {e}")
                # Fallback even on executor crash
                payload, corrections = _build_ai_payload(s, show_grades=show_grades)
                results.append({
                    'reg_no': s.reg_num,
                    'name': payload['name'],
                    'comment': fallback_comment(payload),
                    'ai_generated': False,
                    'ai_warning': "AI generation failed; a template comment was used. Check the server logs.",
                    'corrections': corrections,
                    'gender': payload['gender'],
                    'low_count': len(payload['low_subjects']),
                    'has_event': bool(payload.get('event', '').strip()),
                })

    # Sort by name for consistent ordering
    results.sort(key=lambda x: (x['name'] or '').lower())
    return jsonify(results)


@app.route('/download_comment_report', methods=['POST'])
def download_comment_report():
    """
    POST /download_comment_report
    Accepts JSON body { "mentor_name": "GNANAVEL V K", "comments": [{ "reg_no", "name", "comment" }] }
    Returns a .docx file download.
    """
    if 'user' not in session or session['role'] != 'faculty':
        return jsonify({'error': 'Unauthorized'}), 401

    if not DOCX_AVAILABLE:
        return jsonify({'error': 'python-docx not installed on server'}), 500

    data = request.get_json(silent=True) or {}
    mentor_name = data.get('mentor_name', 'GNANAVEL V K').strip()
    comments = data.get('comments', [])

    doc = Document()

    # --- Page Setup: A4 ---
    section = doc.sections[0]
    section.page_width  = DocxInches(8.27)   # A4 width
    section.page_height = DocxInches(11.69)  # A4 height
    section.left_margin   = DocxInches(1.0)
    section.right_margin  = DocxInches(1.0)
    section.top_margin    = DocxInches(1.0)
    section.bottom_margin = DocxInches(1.0)

    # --- Helper: set run font ---
    def set_run_font(run, size_pt=11, bold=False, italic=False, colour=None):
        run.font.name = 'Arial'
        run.font.size = DocxPt(size_pt)
        run.font.bold = bold
        run.font.italic = italic
        if colour:
            run.font.color.rgb = DocxRGB(*colour)

    # --- Title ---
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run('Mentee Performance Review')
    set_run_font(title_run, size_pt=18, bold=True)

    # --- Subtitle ---
    total = len(comments)
    today_str = datetime.now().strftime('%d %B %Y')
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_para.add_run(
        f"Mentor: {mentor_name}  |  Institution: SIMATS Engineering  |  Mentees: {total}  |  Date: {today_str}"
    )
    set_run_font(subtitle_run, size_pt=10, italic=True, colour=(80, 80, 80))

    doc.add_paragraph()  # spacer

    # --- Summary sentence ---
    low_att_count = sum(1 for c in comments if c.get('low_count', 0) > 0)
    event_count   = sum(1 for c in comments if c.get('has_event', False))
    summary_text = (
        f"Of the {total} mentee(s) in this group, {low_att_count} have at least one subject with attendance "
        f"below 80%, and {event_count} have notable event participation or achievements."
    )
    summary_para = doc.add_paragraph()
    summary_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    summary_run = summary_para.add_run(summary_text)
    set_run_font(summary_run, size_pt=11, italic=True)

    doc.add_paragraph()  # spacer

    # --- Per-student entries ---
    for idx, c in enumerate(comments, start=1):
        name    = (c.get('name') or 'N/A').strip()
        reg_no  = (c.get('reg_no') or '').strip()
        comment = (c.get('comment') or '').strip()

        # Heading: "1. Sayyad Ibrahim (Reg. No: 24REG01)"
        heading_para = doc.add_paragraph()
        heading_run = heading_para.add_run(f"{idx}. {name} (Reg. No: {reg_no})")
        set_run_font(heading_run, size_pt=12, bold=True)

        # Comment paragraph (justified)
        body_para = doc.add_paragraph()
        body_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        body_run = body_para.add_run(comment)
        set_run_font(body_run, size_pt=11)

        doc.add_paragraph()  # spacer between students

    # --- Footer with page numbers ---
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run()
    footer_run.font.name = 'Arial'
    footer_run.font.size = DocxPt(9)

    # Insert PAGE field into footer using XML
    fldChar_begin = DocxOxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    instrText = DocxOxmlElement('w:instrText')
    instrText.text = ' PAGE '
    fldChar_end = DocxOxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    footer_run._r.append(fldChar_begin)
    footer_run._r.append(instrText)
    footer_run._r.append(fldChar_end)

    # Save to in-memory buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name='Mentor_Comment_Report.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        seed_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
