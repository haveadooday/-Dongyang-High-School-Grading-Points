
## 성적 수정 제대로 안되는 부분 수정
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from openpyxl import Workbook, load_workbook
import logging
from datetime import datetime


logging.basicConfig(filename = 'logs/'+datetime.today().strftime("%Y.%m.%d")+'.log', level = logging.DEBUG)
app = Flask(__name__)
app.secret_key = 'Raeto0322'

@app.route('/download/dongyang.ico')
def download_dongyang_icon():
    directory = 'static/'  # 실제 파일이 위치하는 디렉토리 경로를 입력해주세요
    return send_file(f'{directory}/dongyang.ico', as_attachment=True)

@app.after_request
def add_security_headers(response):
    # X-Frame-Options 헤더 추가 (클릭재킹 방지)
    response.headers['X-Frame-Options'] = 'DENY'
    
    # X-Content-Type-Options 헤더 추가 (MIME 타입 스니핑 방지)
    response.headers['X-Content-Type-Options'] = 'nosniff'

    return response

@app.route('/')
def index():
    return render_template('emergency.html')
    #return render_template('servicestop.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port="80", debug=False)
