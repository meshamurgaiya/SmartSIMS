from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_mysqldb import MySQL
from openpyxl import Workbook

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

app = Flask(__name__)
app.secret_key = "smartsims123"

# ===========================
# MySQL Configuration
# ===========================

app.config["MYSQL_HOST"] = "localhost"
app.config["MYSQL_USER"] = "root"
app.config["MYSQL_PASSWORD"] = "your_password"
app.config["MYSQL_DB"] = "studentdb"

mysql = MySQL(app)


# ===========================
# Login Page
# ===========================

@app.route("/")
def login():
    return render_template("login.html")


# ===========================
# Login Check
# ===========================

@app.route("/login", methods=["POST"])
def login_user():

    email = request.form["email"]
    password = request.form["password"]

    cursor = mysql.connection.cursor()

    cursor.execute(
        "SELECT * FROM users WHERE email=%s AND password=%s",
        (email, password)
    )

    user = cursor.fetchone()

    cursor.close()

    if user:
        return redirect(url_for("dashboard"))
    else:
        return "Invalid Email or Password"


# ===========================
# Dashboard
# ===========================

@app.route("/dashboard")
def dashboard():

    cursor = mysql.connection.cursor()

    # Total Students
    cursor.execute("SELECT COUNT(*) FROM students")
    total_students = cursor.fetchone()[0]

    # Average Attendance
    cursor.execute("SELECT AVG(attendance) FROM students")
    avg = cursor.fetchone()[0]

    if avg is None:
        avg = 0

    average_attendance = round(avg)

    # High Risk Students
    cursor.execute(
        "SELECT COUNT(*) FROM students WHERE attendance < 75"
    )
    high_risk = cursor.fetchone()[0]

    # Department Statistics
    cursor.execute("""
        SELECT department, COUNT(*)
        FROM students
        GROUP BY department
    """)

    department_data = cursor.fetchall()

    departments = []
    counts = []

    for row in department_data:
        departments.append(row[0])
        counts.append(row[1])

    cursor.close()

    return render_template(
        "dashboard.html",
        total_students=total_students,
        average_attendance=average_attendance,
        high_risk=high_risk,
        departments=departments,
        counts=counts
    )

# ===========================
# Student Analytics
# ===========================

@app.route("/analytics")
def analytics():

    cursor = mysql.connection.cursor()

    # Total Students
    cursor.execute("SELECT COUNT(*) FROM students")
    total_students = cursor.fetchone()[0]

    # Safe Students
    cursor.execute("SELECT COUNT(*) FROM students WHERE attendance >= 75")
    safe_students = cursor.fetchone()[0]

    # High Risk Students
    cursor.execute("SELECT COUNT(*) FROM students WHERE attendance < 75")
    high_risk = cursor.fetchone()[0]

    # Average Attendance
    cursor.execute("SELECT AVG(attendance) FROM students")
    avg = cursor.fetchone()[0]

    if avg is None:
        avg = 0

    average_attendance = round(avg)

    # Department Statistics
    cursor.execute("""
        SELECT department, COUNT(*)
        FROM students
        GROUP BY department
    """)

    department_data = cursor.fetchall()

    departments = []
    counts = []

    for row in department_data:
        departments.append(row[0])
        counts.append(row[1])

    cursor.close()

    return render_template(
        "analytics.html",
        total_students=total_students,
        safe_students=safe_students,
        high_risk=high_risk,
        average_attendance=average_attendance,
        departments=departments,
        counts=counts
    )

# ===========================
# Student List
# ===========================

@app.route("/students")
def students():

    cursor = mysql.connection.cursor()

    cursor.execute("SELECT * FROM students")

    students = cursor.fetchall()

    cursor.close()

    return render_template(
        "students.html",
        students=students
    )


# ===========================
# Add Student Page
# ===========================

@app.route("/add_student")
def add_student():

    return render_template("add_student.html")

# ===========================
# Save Student
# ===========================

@app.route("/save_student", methods=["POST"])
def save_student():

    name = request.form["name"]
    age = request.form["age"]
    department = request.form["department"]
    year = request.form["year"]
    email = request.form["email"]

    attendance = 100

    cursor = mysql.connection.cursor()

    cursor.execute(
        """
        INSERT INTO students
        (name, age, department, year, email, attendance)
        VALUES (%s,%s,%s,%s,%s,%s)
        """,
        (name, age, department, year, email, attendance)
    )

    mysql.connection.commit()

    cursor.close()

    flash("Student Added Successfully!", "success")

    return redirect(url_for("students"))


# ===========================
# Edit Student
# ===========================

@app.route("/edit_student/<int:id>")
def edit_student(id):

    cursor = mysql.connection.cursor()

    cursor.execute(
        "SELECT * FROM students WHERE id=%s",
        (id,)
    )

    student = cursor.fetchone()

    cursor.close()

    return render_template(
        "edit_student.html",
        student=student
    )



# ===========================
# Update Student
# ===========================

@app.route("/update_student/<int:id>", methods=["POST"])
def update_student(id):

    name = request.form["name"]
    age = request.form["age"]
    department = request.form["department"]
    year = request.form["year"]
    email = request.form["email"]
    attendance = request.form["attendance"]

    cursor = mysql.connection.cursor()

    cursor.execute(
        """
        UPDATE students
        SET
            name=%s,
            age=%s,
            department=%s,
            year=%s,
            email=%s,
            attendance=%s
        WHERE id=%s
        """,
        (
            name,
            age,
            department,
            year,
            email,
            attendance,
            id
        )
    )

    mysql.connection.commit()

    cursor.close()

    flash("Student Updated Successfully!", "success")

    return redirect(url_for("students"))

# ===========================
# Delete Student
# ===========================

@app.route("/delete_student/<int:id>")
def delete_student(id):

    cursor = mysql.connection.cursor()

    cursor.execute(
        "DELETE FROM students WHERE id=%s",
        (id,)
    )

    mysql.connection.commit()

    cursor.close()

    flash("Student Deleted Successfully!", "success")

    return redirect(url_for("students"))

# ===========================
# Export Students to Excel
# ===========================

@app.route("/export_excel")
def export_excel():

    cursor = mysql.connection.cursor()

    cursor.execute("""
        SELECT id, name, age, department, year, email, attendance
        FROM students
    """)

    students = cursor.fetchall()

    cursor.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Students"

    ws.append([
        "ID",
        "Name",
        "Age",
        "Department",
        "Year",
        "Email",
        "Attendance"
    ])

    for student in students:
        ws.append(student)

    file_name = "students.xlsx"

    wb.save(file_name)

    return send_file(
        file_name,
        as_attachment=True
    )

# ===========================
# Export Students to PDF
# ===========================

@app.route("/export_pdf")
def export_pdf():

    cursor = mysql.connection.cursor()

    cursor.execute("""
        SELECT id, name, age, department, year, email, attendance
        FROM students
    """)

    students = cursor.fetchall()

    cursor.close()

    file_name = "students_report.pdf"

    pdf = SimpleDocTemplate(file_name)

    data = []

    data.append([
        "ID",
        "Name",
        "Age",
        "Department",
        "Year",
        "Email",
        "Attendance"
    ])

    for student in students:
        data.append(list(student))

    table = Table(data)

    table.setStyle(TableStyle([

        ("BACKGROUND",(0,0),(-1,0),colors.darkblue),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black),
        ("BACKGROUND",(0,1),(-1,-1),colors.beige),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("BOTTOMPADDING",(0,0),(-1,0),12),

    ]))

    pdf.build([table])

    return send_file(
        file_name,
        as_attachment=True
    )
# ===========================
# Admin Profile
# ===========================

@app.route("/admin_profile")
def admin_profile():

    admin = {
        "name": "Admin",
        "email": "admin@smartsims.com",
        "role": "System Administrator",
        "phone": "+91 9876543210"
    }

    return render_template(
        "admin_profile.html",
        admin=admin
    )


# ===========================
# Logout
# ===========================
@app.route("/high_risk_students")
def high_risk_students():

    cursor = mysql.connection.cursor()

    cursor.execute("SELECT * FROM students WHERE attendance < 75")

    students = cursor.fetchall()

    cursor.close()

    return render_template(
        "high_risk_students.html",
        students=students
    )

# ===========================
# Student Profile
# ===========================

@app.route("/student_profile/<int:id>")
def student_profile(id):

    cursor = mysql.connection.cursor()

    cursor.execute(
        "SELECT * FROM students WHERE id=%s",
        (id,)
    )

    student = cursor.fetchone()

    cursor.close()

    return render_template(
        "student_profile.html",
        student=student
    )
# ===========================
# Change Password
# ===========================

@app.route("/change_password")
def change_password():

    return render_template("change_password.html")

@app.route("/logout")
def logout():

    return redirect(url_for("login"))


# ===========================
# Run Application
# ===========================

if __name__ == "__main__":
    app.run(debug=True)