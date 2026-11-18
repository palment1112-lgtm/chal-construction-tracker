from flask import Flask, render_template, request, redirect, url_for, send_from_directory, send_file
import os
import psycopg2
import psycopg2.extras
import werkzeug.utils
import openpyxl
from io import BytesIO

app = Flask(__name__)

BASE_UPLOAD_FOLDER = 'uploads'
app.config['BASE_UPLOAD_FOLDER'] = BASE_UPLOAD_FOLDER

if not os.path.exists(BASE_UPLOAD_FOLDER):
    os.makedirs(BASE_UPLOAD_FOLDER)

import os

# ดึงค่าจาก Environment บน Render หรือใช้พอร์ต 6543 เป็นค่าสำรอง
DATABASE_URL = os.environ.get('DATABASE_URL') or "postgresql://postgres:4Kwn9Z4T04v0AyYd@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"

def get_db_connection():
    return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # สร้างตาราง Projects
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id SERIAL PRIMARY KEY,
            project_name TEXT NOT NULL
        )
    ''')
    
    # สร้างตาราง Memos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS memos (
            id SERIAL PRIMARY KEY,
            project_id INTEGER REFERENCES projects(id),
            phase TEXT,
            memo_no TEXT,
            detail TEXT,
            send_date TEXT,
            summary_date TEXT,
            final_price TEXT,
            wo_no TEXT,
            boq_status TEXT,
            doc_status TEXT,
            filename TEXT,
            wo_filename TEXT
        )
    ''')
    conn.commit()
    
    # ตรวจสอบว่ามีข้อมูลโครงการเริ่มต้นหรือยัง
    cursor.execute('SELECT COUNT(*) FROM projects')
    if cursor.fetchone()['count'] == 0:
        initial_projects = [
            "BOTANICA FOUR SEASON",
            "BOTANICA GREEN BOULEVARD",
            "BOTANICA MAJESTIA",
            "BOTANICA MONTAZURE",
            "BOTANICA FORESTA",
            "โครงการในอนาคต (รอระบุชื่อ)"
        ]
        for name in initial_projects:
            cursor.execute('INSERT INTO projects (project_name) VALUES (%s)', (name,))
        conn.commit()
        
    cursor.close()
    conn.close()

init_db()

@app.route('/')
def index():
    search_query = request.args.get('search', '')
    filter_status = request.args.get('filter_status', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM projects ORDER BY id ASC')
    projects_rows = cursor.fetchall()
    
    total_memos = 0
    finished_count = 0
    wo_received_count = 0
    
    projects_data = []
    for p in projects_rows:
        proj_id = p['id']
        proj_name = p['project_name']
        
        query = 'SELECT * FROM memos WHERE project_id = %s'
        params = [proj_id]
        
        conditions = []
        if search_query:
            conditions.append('(memo_no ILIKE %s OR detail ILIKE %s OR phase ILIKE %s)')
            params.extend([f'%{search_query}%', f'%{search_query}%', f'%{search_query}%'])
            
        if filter_status == 'รอส่ง BOQ':
            conditions.append('boq_status = %s')
            params.append('รอส่ง BOQ')
        elif filter_status == 'รอ WO':
            conditions.append("(wo_no IS NULL OR wo_no = '' OR doc_status = 'รอราคา')")
            
        if conditions:
            query += ' AND ' + ' AND '.join(conditions)
            
        cursor.execute(query, params)
        memos_rows = cursor.fetchall()
        memos = [dict(row) for row in memos_rows]
        
        cursor.execute('SELECT COUNT(*) FROM memos WHERE project_id = %s', (proj_id,))
        total_memos += cursor.fetchone()['count']
        
        cursor.execute("SELECT COUNT(*) FROM memos WHERE project_id = %s AND doc_status = 'จบราคา'", (proj_id,))
        finished_count += cursor.fetchone()['count']
        
        cursor.execute("SELECT COUNT(*) FROM memos WHERE project_id = %s AND wo_no IS NOT NULL AND wo_no != ''", (proj_id,))
        wo_received_count += cursor.fetchone()['count']
        
        if search_query and not memos and search_query.lower() not in proj_name.lower():
            continue
            
        projects_data.append({
            "id": proj_id,
            "project_name": proj_name,
            "memos": memos
        })
        
    cursor.close()
    conn.close()
    
    stats = {
        "total": total_memos,
        "finished": finished_count,
        "wo_received": wo_received_count
    }
    
    return render_template('index.html', projects=projects_data, search=search_query, filter_status=filter_status, stats=stats)

@app.route('/project/<int:proj_id>/add_memo', methods=['POST'])
def add_memo(proj_id):
    phase = request.form.get('phase', '')
    memo_no = request.form.get('memo_no', '')
    detail = request.form.get('detail', '')
    send_date = request.form.get('send_date', '')
    summary_date = request.form.get('summary_date', '')
    final_price = request.form.get('final_price', '')
    wo_no = request.form.get('wo_no', '')
    boq_status = request.form.get('boq_status', 'รอส่ง BOQ')
    doc_status = request.form.get('doc_status', 'รอราคา')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT project_name FROM projects WHERE id = %s', (proj_id,))
    proj_name = cursor.fetchone()['project_name']
    cursor.close()
    conn.close()
    
    safe_project_name = "".join(c if c.isalnum() else "_" for c in proj_name)
    project_folder = os.path.join(app.config['BASE_UPLOAD_FOLDER'], safe_project_name)
    if not os.path.exists(project_folder):
        os.makedirs(project_folder)
        
    filename = ""
    file = request.files.get('file')
    if file and file.filename != '':
        filename = werkzeug.utils.secure_filename(file.filename)
        file.save(os.path.join(project_folder, filename))
        
    wo_filename = ""
    wo_file = request.files.get('wo_file')
    if wo_file and wo_file.filename != '':
        wo_filename = werkzeug.utils.secure_filename(wo_file.filename)
        wo_file.save(os.path.join(project_folder, wo_filename))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO memos (project_id, phase, memo_no, detail, send_date, summary_date, final_price, wo_no, boq_status, doc_status, filename, wo_filename)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    ''', (proj_id, phase, memo_no, detail, send_date, summary_date, final_price, wo_no, boq_status, doc_status, filename, wo_filename))
    conn.commit()
    cursor.close()
    conn.close()
    return redirect(url_for('index'))

@app.route('/memo/delete/<int:memo_id>', methods=['POST'])
def delete_memo(memo_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM memos WHERE id = %s', (memo_id,))
    conn.commit()
    cursor.close()
    conn.close()
    return redirect(url_for('index'))

@app.route('/memo/edit/<int:memo_id>', methods=['POST'])
def edit_memo(memo_id):
    phase = request.form.get('phase', '')
    memo_no = request.form.get('memo_no', '')
    detail = request.form.get('detail', '')
    send_date = request.form.get('send_date', '')
    summary_date = request.form.get('summary_date', '')
    final_price = request.form.get('final_price', '')
    wo_no = request.form.get('wo_no', '')
    boq_status = request.form.get('boq_status', '')
    doc_status = request.form.get('doc_status', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT p.project_name, m.filename, m.wo_filename FROM memos m 
        JOIN projects p ON m.project_id = p.id 
        WHERE m.id = %s
    ''', (memo_id,))
    row = cursor.fetchone()
    proj_name = row['project_name']
    old_filename = row['filename']
    old_wo_filename = row['wo_filename']
    
    safe_project_name = "".join(c if c.isalnum() else "_" for c in proj_name)
    project_folder = os.path.join(app.config['BASE_UPLOAD_FOLDER'], safe_project_name)
    if not os.path.exists(project_folder):
        os.makedirs(project_folder)
        
    filename = old_filename
    file = request.files.get('file')
    if file and file.filename != '':
        filename = werkzeug.utils.secure_filename(file.filename)
        file.save(os.path.join(project_folder, filename))
        
    wo_filename = old_wo_filename
    wo_file = request.files.get('wo_file')
    if wo_file and wo_file.filename != '':
        wo_filename = werkzeug.utils.secure_filename(wo_file.filename)
        wo_file.save(os.path.join(project_folder, wo_filename))
        
    cursor.execute('''
        UPDATE memos SET phase=%s, memo_no=%s, detail=%s, send_date=%s, summary_date=%s, final_price=%s, wo_no=%s, boq_status=%s, doc_status=%s, filename=%s, wo_filename=%s
        WHERE id=%s
    ''', (phase, memo_no, detail, send_date, summary_date, final_price, wo_no, boq_status, doc_status, filename, wo_filename, memo_id))
        
    conn.commit()
    cursor.close()
    conn.close()
    return redirect(url_for('index'))

@app.route('/export_excel')
def export_excel():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM projects ORDER BY id ASC')
    projects_rows = cursor.fetchall()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CHAL BOQ Summary"
    
    headers = ["ชื่อโครงการ", "PHASE", "Memo No.", "รายละเอียด", "วันที่ส่งราคา", "วันที่สรุปราคา", "ราคาสุดท้าย", "เลขที่ WO", "ไฟล์ WO", "ไฟล์ BOQ", "สถานะ BOQ", "สถานะเอกสาร"]
    ws.append(headers)
    
    for p in projects_rows:
        cursor.execute('SELECT * FROM memos WHERE project_id = %s', (p['id'],))
        memos = cursor.fetchall()
        for m in memos:
            ws.append([
                p['project_name'],
                m['phase'],
                m['memo_no'],
                m['detail'],
                m['send_date'],
                m['summary_date'] or '',
                m['final_price'],
                m['wo_no'],
                m['wo_filename'] or '',
                m['filename'] or '',
                m['boq_status'],
                m['doc_status']
            ])
            
    cursor.close()
    conn.close()
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='CHAL_BOQ_Report.xlsx')

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['BASE_UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)