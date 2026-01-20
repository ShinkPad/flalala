from flask import Flask, render_template, request, redirect, url_for, session
import random
import datetime
import os
try:
    import openpyxl
except Exception:
    openpyxl = None

app = Flask(__name__)
app.secret_key = 'change-this-secret'

@app.route('/')
def index():
    dt_now = datetime.datetime.now()
    return render_template('htmls/login.html', time=dt_now.strftime('%Y年%m月%d日 %H:%M:%S'))

@app.route('/password')
def password():
    login = request.args.get('login')
    forgot = str(random.randint(0, 4))  # 0〜4の乱数
    return render_template('htmls/password.html', login=login, forgot=forgot)


def read_users_from_excel(path=None):
    """Return dict of username->password from an Excel file if available."""
    users = {}
    if path is None:
        path = os.path.join(app.root_path, 'data', 'users.xlsx')
    if openpyxl and os.path.exists(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            username = str(row[0]).strip() if row[0] is not None else ''
            password = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ''
            if username:
                users[username] = password
        return users
 


@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username', '')
    password = request.form.get('password', '')
    users = read_users_from_excel()
    # authenticate
    if username in users and users[username] == password:
        session['username'] = username
        return redirect(url_for('home'))
    # login failed
    return render_template('htmls/login.html', time=datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S'), error='ユーザー名かパスワードが違います')


@app.route('/home')
def home():
    username = session.get('username')
    if not username:
        return redirect(url_for('index'))
    dt_now = datetime.datetime.now()
    time = dt_now.strftime('%Y年%m月%d日 %H:%M:%S')
    reports = []
    return render_template('htmls/home.html', username=username, time=time, reports=reports)

if __name__ == '__main__':
    app.run(debug=True)
