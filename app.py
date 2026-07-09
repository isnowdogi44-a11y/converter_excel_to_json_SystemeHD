import sys
import os
import json
import io
import webbrowser
import threading
import time
import pandas as pd
from flask import Flask, request, send_file, jsonify

# Определяем путь к папке с шаблонами
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
else:
    template_folder = 'templates'

app = Flask(__name__, template_folder=template_folder)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def convert_excel_to_json_data(excel_file):
    """Конвертирует Excel-файл в JSON (возвращает список словарей)"""
    df = pd.read_excel(excel_file, header=0, dtype=str)

    required_columns = [
        'PlcTerminalName', 'SysObj', 'SysObjNode', 'BACNetType',
        'BACNetId', 'OpcType', 'VarField', 'PlcVarDir',
        'Hmi', 'FsaSchemeName', 'VarDesc'
    ]

    existing_columns = [col for col in required_columns if col in df.columns]
    df = df[existing_columns]
    df = df.fillna('')

    # Обработка BACNetId
    if 'BACNetId' in df.columns:
        def clean_bacnetid(val):
            if val == '':
                return ''
            try:
                f = float(val)
                if f.is_integer():
                    return str(int(f))
                else:
                    return str(f)
            except ValueError:
                return val

        df['BACNetId'] = df['BACNetId'].apply(clean_bacnetid)

    # Обработка Hmi
    if 'Hmi' in df.columns:
        def clean_hmi(val):
            if val == '':
                return ''
            if isinstance(val, str):
                val_low = val.lower()
                if val_low in ('true', '1', 'yes', 'on'):
                    return True
                else:
                    return False
            return bool(val)

        df['Hmi'] = df['Hmi'].apply(clean_hmi)

    data = df.to_dict(orient='records')

    # Удаляем FsaSchemeName, если он пустой
    for row in data:
        if row.get('FsaSchemeName') == '':
            del row['FsaSchemeName']

    return data


@app.route('/')
def index():
    html = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel → JSON конвертер</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            padding: 50px 40px;
            max-width: 600px;
            width: 100%;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            text-align: center;
        }
        h1 { color: #333; margin-bottom: 10px; font-weight: 600; font-size: 28px; }
        .subtitle { color: #777; margin-bottom: 30px; font-size: 14px; }
        .upload-area {
            border: 2px dashed #ccc;
            border-radius: 12px;
            padding: 40px 20px;
            transition: all 0.3s ease;
            cursor: pointer;
            background: #fafafa;
        }
        .upload-area:hover { border-color: #667eea; background: #f0f2ff; }
        .upload-area.dragover { border-color: #667eea; background: #e8ecff; }
        .upload-icon { font-size: 48px; margin-bottom: 10px; }
        .upload-text { color: #555; font-size: 16px; }
        .upload-text strong { color: #667eea; }
        .file-info { margin-top: 15px; color: #333; font-weight: 500; display: none; }
        .file-info.show { display: block; }
        #fileInput { display: none; }
        .btn-convert {
            margin-top: 25px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 14px 40px;
            border-radius: 50px;
            font-size: 18px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            width: 100%;
        }
        .btn-convert:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4); }
        .btn-convert:disabled { background: #ccc; cursor: not-allowed; }
        .loader { display: none; margin: 20px auto; border: 4px solid #f3f3f3; border-top: 4px solid #667eea; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; }
        .loader.show { display: block; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .error { color: #e74c3c; margin-top: 15px; display: none; background: #fde8e8; padding: 12px; border-radius: 8px; }
        .error.show { display: block; }
        .footer { margin-top: 25px; font-size: 12px; color: #aaa; }
        .version { font-size: 12px; color: #aaa; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel → JSON</h1>
        <p class="subtitle">Загрузите Excel-файл и получите JSON</p>
        <div class="upload-area" id="dropZone">
            <div class="upload-icon">📁</div>
            <div class="upload-text"><strong>Нажмите</strong> или перетащите файл сюда</div>
            <div class="file-info" id="fileInfo"></div>
            <input type="file" id="fileInput" accept=".xlsx,.xls">
        </div>
        <button class="btn-convert" id="convertBtn" disabled>🔄 Конвертировать</button>
        <div class="loader" id="loader"></div>
        <div class="error" id="error"></div>
        <div class="footer">Поддерживаются .xlsx и .xls файлы</div>
        <div class="version">Версия 1.0</div>
    </div>
    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const convertBtn = document.getElementById('convertBtn');
        const loader = document.getElementById('loader');
        const error = document.getElementById('error');
        let selectedFile = null;

        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handleFile(e.target.files[0]);
        });
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('dragover'); });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
        });

        function handleFile(file) {
            const ext = file.name.split('.').pop().toLowerCase();
            if (!['xlsx', 'xls'].includes(ext)) {
                showError('Пожалуйста, загрузите файл .xlsx или .xls');
                return;
            }
            selectedFile = file;
            fileInfo.textContent = '📎 ' + file.name + ' (' + (file.size / 1024).toFixed(1) + ' KB)';
            fileInfo.classList.add('show');
            convertBtn.disabled = false;
            error.classList.remove('show');
        }

        function showError(msg) {
            error.textContent = '❌ ' + msg;
            error.classList.add('show');
            setTimeout(() => error.classList.remove('show'), 5000);
        }

        convertBtn.addEventListener('click', async () => {
            if (!selectedFile) return;
            convertBtn.disabled = true;
            loader.classList.add('show');
            error.classList.remove('show');
            const formData = new FormData();
            formData.append('file', selectedFile);
            try {
                const response = await fetch('/convert', { method: 'POST', body: formData });
                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'Ошибка конвертации');
                }
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = selectedFile.name.replace(/\\.[^.]+$/, '') + '.json';
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
            } catch (err) {
                showError(err.message);
            } finally {
                loader.classList.remove('show');
                convertBtn.disabled = false;
            }
        });
    </script>
</body>
</html>
    """
    return html


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не выбран'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Разрешены только .xlsx и .xls файлы'}), 400

    try:
        data = convert_excel_to_json_data(file)

        json_str = '[\n'
        for i, row in enumerate(data):
            json_line = json.dumps(row, ensure_ascii=False, separators=(', ', ': '))
            json_str += f'  {json_line}'
            if i < len(data) - 1:
                json_str += ',\n'
            else:
                json_str += '\n'
        json_str += ']'

        output = io.BytesIO()
        output.write(json_str.encode('utf-8'))
        output.seek(0)

        return send_file(
            output,
            mimetype='application/json',
            as_attachment=True,
            download_name=file.filename.rsplit('.', 1)[0] + '.json'
        )
    except Exception as e:
        return jsonify({'error': f'Ошибка конвертации: {str(e)}'}), 500


def open_browser():
    """Открывает браузер через 1.5 секунды после запуска сервера"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:5000')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))

    # Запускаем поток для открытия браузера
    threading.Thread(target=open_browser, daemon=True).start()

    # Запускаем сервер
    app.run(debug=False, host='0.0.0.0', port=port)