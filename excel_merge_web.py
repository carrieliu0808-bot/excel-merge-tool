#!/usr/bin/env python3
"""
Excel Sheet 合并工具 - Web 版
上传 Excel 文件，合并所有 sheet 后下载
"""

from flask import Flask, render_template_string, request, send_file
import pandas as pd
from pathlib import Path
from datetime import datetime
import tempfile
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 限制

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Sheet 合并工具</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            max-width: 600px;
            margin: 50px auto;
            padding: 20px;
            background: #f5f5f5;
        }
        .container {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            margin: 0 0 10px 0;
            color: #333;
        }
        .subtitle {
            color: #666;
            margin-bottom: 30px;
        }
        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 8px;
            padding: 40px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            margin-bottom: 20px;
        }
        .upload-area:hover, .upload-area.dragover {
            border-color: #4CAF50;
            background: #f9fff9;
        }
        .upload-area input {
            display: none;
        }
        .upload-icon {
            font-size: 48px;
            margin-bottom: 10px;
        }
        .file-list {
            margin: 20px 0;
            text-align: left;
        }
        .file-item {
            padding: 10px 15px;
            background: #f0f0f0;
            border-radius: 6px;
            margin: 5px 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .file-item .remove {
            color: #e74c3c;
            cursor: pointer;
            font-weight: bold;
        }
        .options {
            margin: 20px 0;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 8px;
        }
        .options label {
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
        }
        button {
            width: 100%;
            padding: 15px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            cursor: pointer;
            transition: background 0.3s;
        }
        button:hover { background: #45a049; }
        button:disabled {
            background: #ccc;
            cursor: not-allowed;
        }
        .message {
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .message.success { background: #d4edda; color: #155724; }
        .message.error { background: #f8d7da; color: #721c24; }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        .loading.show { display: block; }
        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #4CAF50;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel Sheet 合并工具</h1>
        <p class="subtitle">上传 Excel 文件，自动合并所有 Sheet</p>
        
        {% if message %}
        <div class="message {{ message_type }}">{{ message }}</div>
        {% endif %}
        
        <form id="uploadForm" action="/merge" method="post" enctype="multipart/form-data">
            <div class="upload-area" id="uploadArea">
                <input type="file" name="files" id="fileInput" multiple accept=".xlsx,.xls">
                <div class="upload-icon">📁</div>
                <div>点击或拖拽上传 Excel 文件</div>
                <div style="color: #999; font-size: 14px; margin-top: 5px;">支持多个文件，.xlsx / .xls 格式</div>
            </div>
            
            <div class="file-list" id="fileList"></div>
            
            <div class="options">
                <label>
                    <input type="checkbox" name="keep_source" checked>
                    保留来源信息（文件名、Sheet名）
                </label>
            </div>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <div>正在处理中...</div>
            </div>
            
            <button type="submit" id="submitBtn" disabled>合并并下载</button>
        </form>
    </div>
    
    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const submitBtn = document.getElementById('submitBtn');
        const uploadForm = document.getElementById('uploadForm');
        const loading = document.getElementById('loading');
        
        let selectedFiles = [];
        
        uploadArea.addEventListener('click', () => fileInput.click());
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });
        
        fileInput.addEventListener('change', () => {
            handleFiles(fileInput.files);
        });
        
        function handleFiles(files) {
            for (let file of files) {
                if (file.name.match(/\.xlsx?$/i) && !selectedFiles.find(f => f.name === file.name)) {
                    selectedFiles.push(file);
                }
            }
            updateFileList();
        }
        
        function updateFileList() {
            fileList.innerHTML = selectedFiles.map((file, index) => `
                <div class="file-item">
                    <span>📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>
                    <span class="remove" onclick="removeFile(${index})">✕</span>
                </div>
            `).join('');
            submitBtn.disabled = selectedFiles.length === 0;
        }
        
        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
        }
        
        uploadForm.addEventListener('submit', (e) => {
            e.preventDefault();
            
            const formData = new FormData();
            selectedFiles.forEach(file => formData.append('files', file));
            formData.append('keep_source', document.querySelector('[name="keep_source"]').checked);
            
            loading.classList.add('show');
            submitBtn.disabled = true;
            
            fetch('/merge', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (response.ok) {
                    return response.blob();
                }
                return response.text().then(text => { throw new Error(text); });
            })
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'merged_' + new Date().toISOString().slice(0,10) + '.xlsx';
                a.click();
                window.URL.revokeObjectURL(url);
            })
            .catch(error => {
                alert('处理失败: ' + error.message);
            })
            .finally(() => {
                loading.classList.remove('show');
                submitBtn.disabled = false;
            });
        });
    </script>
</body>
</html>
'''


def merge_excel_files(files, keep_source=True):
    """合并上传的 Excel 文件"""
    all_data = []
    
    for file in files:
        try:
            xlsx = pd.ExcelFile(file)
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name=sheet_name)
                if df.empty:
                    continue
                if keep_source:
                    df.insert(0, '_来源文件', file.filename)
                    df.insert(1, '_来源Sheet', sheet_name)
                all_data.append(df)
        except Exception as e:
            raise ValueError(f"读取 {file.filename} 失败: {str(e)}")
    
    if not all_data:
        raise ValueError("没有找到有效数据")
    
    return pd.concat(all_data, ignore_index=True, sort=False)


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist('files')
    
    if not files or all(f.filename == '' for f in files):
        return "请上传至少一个 Excel 文件", 400
    
    keep_source = request.form.get('keep_source') == 'true'
    
    try:
        merged_df = merge_excel_files(files, keep_source)
        
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            merged_df.to_excel(tmp.name, index=False, engine='openpyxl')
            tmp_path = tmp.name
        
        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=f'merged_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except ValueError as e:
        return str(e), 400
    except Exception as e:
        return f"处理失败: {str(e)}", 500


if __name__ == '__main__':
    print("🐾 Excel 合并工具启动中...")
    print("📍 访问 http://localhost:5050")
    app.run(host='0.0.0.0', port=5050, debug=True)
