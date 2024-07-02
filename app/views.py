"""
Flask Documentation:     http://flask.pocoo.org/docs/
Jinja2 Documentation:    http://jinja.pocoo.org/2/documentation/
Werkzeug Documentation:  http://werkzeug.pocoo.org/documentation/
This file creates your application.
"""

from app import app, db
from flask import render_template, request, redirect, url_for, flash, make_response, send_file, jsonify
import pandas as pd
import pyexcel
from app.forms import UserForm
from app.models import User
import uuid
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
from werkzeug.utils import secure_filename
import random
import string
# import sqlite3

# CONSTANT
ADMIN_TK = "123456789" # extract name
ADMIN_ROW_INDEX = 264499
FINISH_ROW_INDEX = 900 # to Xac nhan cua can bo thu truong don vi
app.config['UPLOAD_FOLDER'] = os.getcwd()
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
###
# Routing for your application.
###
def format_currency(amount):
    try:
        """
        Format a number as currency in Vietnamese dong (VND).
        """
        # Convert the number to a string and reverse it
        reversed_amount = str(amount)[::-1]
        
        # Insert a dot (.) after every three digits
        formatted_amount = '.'.join(reversed_amount[i:i+3] for i in range(0, len(reversed_amount), 3))
        
        # Reverse the formatted string back to the original order
        formatted_amount = formatted_amount[::-1]
        
        # Add the currency symbol (₫)
        formatted_amount += ' ₫'
    except:
        formatted_amount = amount
    return formatted_amount

def generate_random_string(length=6):
    characters = string.ascii_letters + string.digits
    random_string = ''.join(random.choices(characters, k=length))
    return random_string

@app.route('/home')
def home():
    uuid = request.cookies.get('uuid')
    if not uuid:
        return redirect(url_for('login'))
    user = User.query.filter_by(uuid=uuid).first()
    print(user, user.uuid, ADMIN_TK)
    if uuid == ADMIN_TK:
        return render_template('admin.html')
    # Read the Excel file
    records = pyexcel.get_records(file_name='data.xlsx')  # Replace 'data.xls' with the path to your Excel file
    # del records[0] # remove blank record
    del records[FINISH_ROW_INDEX:]
    # Format the records with specific columns
    formatted_records = []
    for index, record in enumerate(records):
        data = list(record.values())
        # print(list(record.keys()))
        # print(list(record.values()))
        currency_money = data[user.rowIndex]
        if currency_money is None or currency_money == '':
            continue
        try:
            if index >= 6:
                if int(currency_money) == 0:
                    continue
                if currency_money is not None and currency_money != '':
                    currency_money = format_currency(int(currency_money))
        except:
            currency_money = data[user.rowIndex]

        # Select specific columns from the record
        formatted_record = {
            # 'TT': data[0],
            'Ngày, tháng': data[1],
            'CK': data[2],
            'TM/CK': data[3],
            'Nội dung': data[4],
            'Số liệu': currency_money,
            # Add more columns as needed
        }
        formatted_records.append(formatted_record)
    # Render the records in an HTML table
    return render_template('index.html', records=formatted_records, username=user.name)

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        uuid = request.form['uuid']
        # Check if the UUID exists in the User table
        user = User.query.filter_by(uuid=uuid).first()
        if user:
            # Create a response object
            response = make_response(redirect(url_for('home')))
            # Set the cookie
            response.set_cookie('uuid', uuid, max_age=3600)

            # Redirect to the home page after successful login
            return response
        else:
            # Flash an alert message if login fails
            flash('UUID không chính xác. Xin mời nhập lại!', 'error')
            # Redirect back to the login page
            return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    # Create a response object
    response = make_response(redirect(url_for('login')))
    # Clear the UUID cookie by setting its value to an empty string and setting its max_age to 0
    response.set_cookie('uuid', '', max_age=0)
    return response

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file part'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'No selected file'})
    
    if file and allowed_file(file.filename):
        filename = 'data.xlsx'  # Fixed filename
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        return jsonify({'status': 'success', 'message': 'File successfully uploaded'})
    else:
        return jsonify({'status': 'error', 'message': 'Allowed file types are xlsx only'})

@app.route('/clear_db')
def clear_db():
    try:
        db.create_all()
        # Delete all records from the User table
        User.query.delete()
        # Commit the changes to the database
        db.session.commit()
        return "All data in the User table deleted successfully."
    except Exception as e:
        # If any exception occurs, print the error message
        print(f"An error occurred: {str(e)}")
        # Optionally, you can rollback the database session in case of an error
        db.session.rollback()
        return "Error"
    
@app.route('/import_excel')
def import_excel():
    data_sample = [
    {"name": "Nguyễn Văn Hiệu", "code": "lVKcor"},
    {"name": "Nguyễn Kiều Oanh", "code": "RHGe9i"},
    {"name": "Hoàng Trọng Nghĩa", "code": "QW3Ryx"},
    {"name": "Nguyễn Việt Khôi", "code": "Wbq50e"},
    {"name": "Nguyễn Văn Thái", "code": "uYdHud"},
    {"name": "Nguyễn Hồng Hà", "code": "lHKmMV"},
    {"name": "Võ Thị Thanh Tâm", "code": "CSGf0p"},
    {"name": "Phạm Văn Hứa", "code": "oKIer4"},
    {"name": "Bàng Xuân Hùng", "code": "oX1NrT"},
    {"name": "Dư Đức Thắng", "code": "ECPZWr"},
    {"name": "Hồ Xuân Hương", "code": "SDMAkB"},
    {"name": "Nguyễn Thị Vân Tú", "code": "vDCnvd"},
    {"name": "Hoàng Thị Tuyết Mai", "code": "xDHThM"},
    {"name": "Đỗ Huy Thưởng", "code": "LafQCZ"},
    {"name": "Nguyễn Thu Thủy", "code": "zGQpca"},
    {"name": "Nguyễn Ngọc Trực", "code": "5V3Jq0"},
    {"name": "Trần Thị An", "code": "vzK23X"},
    {"name": "Trần Nhật Lam Duyên", "code": "zScg03"},
    {"name": "Đinh Việt Hưng", "code": "FMosbn"},
    {"name": "Hoàng Thúy Quỳnh", "code": "WEWtkj"},
    {"name": "Vũ Hoài Đức", "code": "sIlOnZ"},
    {"name": "Đặng Thị Mùi", "code": "D1kz2b"},
    {"name": "Bùi Đại Dũng", "code": "tF38ps"},
    {"name": "Nguyễn Thị Hiền", "code": "j388MR"},
    {"name": "Nguyễn Thị Thanh Huyền", "code": "0L3jqI"},
    {"name": "Mai Thị Hạnh", "code": "P9HFWn"},
    {"name": "Phạm Quỳnh Phương", "code": "2PNh8E"},
    {"name": "Lê Phước Anh", "code": "NWQIrE"},
    {"name": "Nguyễn Thu Hương", "code": "lHqwvd"},
    {"name": "Lư Thị Thanh Lê", "code": "WENCLC"},
    {"name": "Đỗ Xuân Đức", "code": "8Ho28R"},
    {"name": "Vũ Đường Luân", "code": "dHoF3g"},
    {"name": "Đinh Việt Hải", "code": "aGDcmp"},
    {"name": "Trần Thị Hoan", "code": "wo5Tah"},
    {"name": "Nguyễn Thu Thủy-1", "code": "3gs3fZ"},
    {"name": "Nguyễn Thị Lan Anh", "code": "ncC61l"},
    {"name": "Nguyễn Hữu Cung", "code": "hVir3f"},
    {"name": "Bùi Thị Thanh Hương", "code": "uwmYoA"},
    {"name": "Dương Văn Hào", "code": "Uebxez"},
    {"name": "Nguyễn Thị Mai Lan", "code": "t3UoWY"},
    {"name": "Nguyễn Anh Thư", "code": "r8qhor"},
    {"name": "Nguyễn Cẩm Chi", "code": "2c9Raj"},
    {"name": "Bùi Thị Thanh Hoa", "code": "AIaLol"},
    {"name": "Trần Điệp Thành", "code": "kZguCv"},
    {"name": "Lê Thị Hà", "code": "JhMET3"},
    {"name": "Trần Hoài", "code": "EPXcOE"},
    {"name": "Nguyễn Thị Thanh Mai", "code": "v8qZRO"},
    {"name": "Trần Yên Thế", "code": "smhMaf"},
    {"name": "Trương Thị Thu Thủy", "code": "8yZPLa"},
    {"name": "Trần Thị Thy Trà", "code": "BkPg5H"},
    {"name": "Đào Mạnh Đạt", "code": "JCwiHg"},
    {"name": "Hoàng Thị Thu Hà", "code": "TMch3F"},
    {"name": "Trần Quốc Trung", "code": "jabDvc"},
    {"name": "Lê Xuân Thái", "code": "FokgZy"},
    {"name": "Phan Quang Anh", "code": "FpRLQQ"},
    {"name": "Nguyễn Thị Thu Hương", "code": "yEVQZJ"},
    {"name": "Phạm Thị Kiều Ly", "code": "Erg7Ja"},
    {"name": "Nguyễn Văn Huấn", "code": "k0XTmj"},
    {"name": "Lê Minh Sơn", "code": "D23iM3"},
    {"name": "Chu Mạnh Hùng", "code": "uLQGey"},
    {"name": "Vũ Kim Yến", "code": "fl1w4o"},
    {"name": "Nguyễn Ngọc Minh", "code": "MLXHCT"},
    {"name": "Nguyễn Thị Oanh", "code": "0RPo25"},
    {"name": "Vũ Thành Trung", "code": "wsqqKo"},
    {"name": "Nhữ Mạnh Tiến", "code": "q6NNok"},
    {"name": "Vũ Thanh Ngọc", "code": "nNvfqf"},
    {"name": "Nguyễn Thị Minh Thảo", "code": "giLzpB"},
    {"name": "Nguyễn Thị Thu Hương-1", "code": "HAVv34"},
    {"name": "Hoàng Văn Hiệp", "code": "lXMHw4"},
    {"name": "Vũ Đình Hoàng Anh Tuấn", "code": "Fut0bc"},
    {"name": "Phạm Thị Thanh Hằng", "code": "BZWwSc"},
    {"name": "Nguyễn Hà Khoa Học", "code": "2cZQL2"},
    {"name": "Nguyễn Văn Minh", "code": "CfgGnI"},
    {"name": "Đinh Thế Anh", "code": "nH3xtT"},
    {"name": "Uông Thị Huyền Hạnh", "code": "cd6nf9"},
    {"name": "Nguyễn Thế Sơn", "code": "CvaVaW"},
    {"name": "Triệu Kim Trường", "code": "lORYR1"},
    {"name": "Lương Thị Hường", "code": "SCL4Gi"},
    {"name": "Nguyễn Đức Thái", "code": "CRv8C9"},
    {"name": "Trần Việt Tùng", "code": "7ggMbF"},
    {"name": "Đặng Thu Phương", "code": "FRINeO"},
    {"name": "Nguyễn Bích Ngọc", "code": "2yXgSf"},
    {"name": "Đỗ Ngọc Anh", "code": "BhQ8Wf"},
    {"name": "Nguyễn Thị Thanh Xuân", "code": "AW9OYG"},
    {"name": "Nguyễn Thị Thu Hà", "code": "79J5xd"},
    {"name": "Huỳnh Thị Hòa", "code": "Bpj2zM"},
    {"name": "Kiều Trung Kiên", "code": "SN1xjx"},
    {"name": "Hoàng Diễn Thanh", "code": "33LSST"},
    {"name": "Nguyễn Hồng Nhung", "code": "u2RrMr"},
    {"name": "Ngô Xuân Phú", "code": "eYgNt8"},
    {"name": "Nguyễn Quang Vinh", "code": "W3Vs36"},
    {"name": "Nguyễn Thị Tuyết Trinh", "code": "ZmltMm"},
    {"name": "Nguyễn Văn Minh-1", "code": "KdUJCX"},
    {"name": "Nguyễn Hoàng Phương Minh", "code": "vQDkE1"},
    {"name": "Nguyễn Thị Thanh Hồng", "code": "gl7tCL"},
    {"name": "Nguyễn Thị Giang Nam", "code": "8RkNXQ"},
    {"name": "Phạm Thị Hồng Nhung", "code": "6XtZ9G"},
    {"name": "Nguyễn Hồng Hạnh", "code": "DiRUAf"},
    {"name": "Nguyễn Tùng Linh", "code": "sLXDES"},
    {"name": "Triệu Minh Hải", "code": "t1bpDB"},
    {"name": "Lê Quang Pháp", "code": "RFvkSF"},
    {"name": "Thái Nhật Minh", "code": "9qskfG"},
    {"name": "Ngô Đức Duy", "code": "yqwpMT"},
    {"name": "Đinh Thị Thu", "code": "949pRP"},
    {"name": "Vũ Thị Trâm Anh", "code": "HVZsRZ"},
    {"name": "Nguyễn Thị Tuệ Thư", "code": "1TwMfd"},
    {"name": "Phạm Minh Quân", "code": "S2Ltlx"},
    {"name": "Kiều Thị Yến", "code": "h2umM7"},
    {"name": "Lê Duy Khương", "code": "W3UQrM"},
    {"name": "Nguyễn Đức Tuấn", "code": "vwkQ4n"},
    {"name": "Chu Trung Tiến", "code": "e0ak2X"},
    {"name": "Phạm Minh Tâm", "code": "3Rz3oB"},
    {"name": "Phạm Thị Thanh Xuân", "code": "iWpAuo"},
    {"name": "Trần Minh Anh", "code": "plHR6L"},
    {"name": "Hoàng Huy Dương", "code": "xt1RkQ"},
    {"name": "Nguyễn Thị Hoài Thương", "code": "xWV0x6"},
    {"name": "Nguyễn Chí Trung", "code": "DoqSaa"},
    {"name": "Nguyễn Thảo Ly", "code": "ejnb4s"},
    {"name": "Tạ Ngọc Ánh", "code": "wVubnx"},
    {"name": "Phạm Thị Mai Hương", "code": "LWomP3"},
    {"name": "Lăng Thị Hồng Nhung", "code": "kiNvO3"},
    {"name": "Nguyễn Thị Minh Châu", "code": "eElSDl"}
]
    # Check if the user already exists
    existing_user = User.query.filter_by(name="ADMIN").first()
    if existing_user is None:
        # User does not exist, create and add to the session
        user = User("9999", "ADMIN", str(ADMIN_TK), ADMIN_ROW_INDEX)
        db.session.add(user)
        db.session.commit()

    status = True
    try:
        # Read the Excel file
        records = pyexcel.get_records(file_name='data.xlsx')  # Replace 'data.xls' with the path to your Excel file
        list_user = list(records[0].keys()) # danh sách users
        # Danh sách mã số
        list_code = list(records[1].values())
        start_user_index = 5 # bắt đầu từ index User
        end_user_index = len(list_user) # 
        list_user_format = list_user[start_user_index:end_user_index]
        for user_name in list_user_format:
            # Check if the user already exists
            existing_user = User.query.filter_by(name=user_name).first()
            found_item = next((item for item in data_sample if item["name"] == user_name), None)
            if found_item is None:
                found_item = {"name": user_name, "code": generate_random_string()}
            if existing_user is None:
                # User does not exist, create and add to the session
                user = User(str(list_code[start_user_index]), user_name, found_item["code"], start_user_index)
                db.session.add(user)
            else:
                existing_user.uuid = found_item["code"]
                existing_user.code = str(list_code[start_user_index])
            start_user_index = start_user_index + 1
        db.session.commit()
    except Exception as e:
        # If any exception occurs, print the error message
        print(f"An error occurred: {str(e)}")
        # Optionally, you can rollback the database session in case of an error
        db.session.rollback()
        status = False
    return render_template('about.html', status=status)

def generate_message(name, uuid):
    message = f"Phòng KHTC xin gửi thầy, cô đường link http://103.167.89.184:6868 và mật khẩu {uuid} để truy cập dữ liệu thu nhập hàng tháng. Trân trọng!"
    return message
@app.route('/export_excel')
def export_excel():
    uuid = request.cookies.get('uuid')
    print(uuid, ADMIN_TK, uuid== ADMIN_TK)
    if not uuid or uuid != ADMIN_TK:
        return redirect(url_for('login'))
    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["UUID", "Họ và tên", "Lời nhắn"])
    # Set column widths and styles for header row
    header_row = ws[1]
    for cell in header_row:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        column_letter = get_column_letter(cell.column)
        ws.column_dimensions[column_letter].width = 50  # Set width for all columns

    # Add data to the worksheet
    users = db.session.query(User).all()
    for row_index, user in enumerate(users, start=2):  # Start from row 2 for data rows
        ws.cell(row=row_index, column=1, value=user.uuid)
        ws.cell(row=row_index, column=2, value=user.name)
        message = generate_message(user.name, user.uuid)
        ws.cell(row=row_index, column=3, value=message)
        # Adjust column width to fit content
        # Adjust column width to fit content and set alignment to justified
    for col in range(1, 4):  # Adjusting columns 1 to 3
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].auto_size = True  # Auto size column width
        for row in ws.iter_rows(min_row=2, min_col=col, max_row=ws.max_row, max_col=col):
            for cell in row:
                cell.alignment = Alignment(horizontal='justify')  # Justify align for data cells

    # Save the workbook to a BytesIO buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # Send the file to the client for download
    return send_file(buffer, as_attachment=True, attachment_filename='data.xlsx')

@app.route('/users')
def show_users():
    uuid = request.cookies.get('uuid')
    if not uuid or uuid != ADMIN_TK:
        return redirect(url_for('login'))
    users = db.session.query(User).all() # or you could have used User.query.all()

    return render_template('show_users.html', users=users)

# Flash errors from the form if validation fails
def flash_errors(form):
    for field, errors in form.errors.items():
        for error in errors:
            flash(u"Error in the %s field - %s" % (
                getattr(form, field).label.text,
                error
            ))

###
# The functions below should be applicable to all Flask apps.
###

@app.route('/<file_name>.txt')
def send_text_file(file_name):
    """Send your static text file."""
    file_dot_text = file_name + '.txt'
    return app.send_static_file(file_dot_text)


@app.after_request
def add_header(response):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    response.headers['X-UA-Compatible'] = 'IE=Edge,chrome=1'
    response.headers['Cache-Control'] = 'public, max-age=600'
    return response


@app.errorhandler(404)
def page_not_found(error):
    """Custom 404 page."""
    return render_template('404.html'), 404


if __name__ == '__main__':
    app.run(debug=True,host="0.0.0.0",port="8080")
