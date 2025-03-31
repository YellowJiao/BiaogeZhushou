from flask import Flask, request, jsonify, send_file
import pandas as pd
import os
import io
import uuid
import csv
import chardet
from werkzeug.utils import secure_filename
import requests
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB文件大小限制

@app.errorhandler(413)
def request_entity_too_large(error):
    app.logger.error(f'File upload too large: {error}')
    return jsonify({'error': '文件大小超过100MB限制'}), 413

# 根路径路由处理
@app.route('/')
def index():
    return send_file('index.html')

# 从环境变量读取 API 密钥
deepseek_api_key = os.getenv("DEEPSEEK_API_KEY") # 你需要从 DeepSeek 获取 API 密钥

# 检查环境变量是否已设置
if not deepseek_api_key:
    raise ValueError("请确保设置了 DEEPSEEK_API_KEY 环境变量。")

# 全局变量存储会话历史
conversation_history = {}

# 上传文件保存的目录
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# API 选择
API_CHOICES = ['deepseek']  # 仅支持 DeepSeek API

def call_deepseek_api(messages, api_key):
    """调用 DeepSeek-R1 API (需要替换为实际的 API 调用方式)"""
    try:
        url = "https://api.deepseek.com/v1/chat/completions"  # 替换为 DeepSeek 的 API 端点

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"  # 或者其他认证方式
        }

        data = {
            "model": "deepseek-chat",
            "messages": messages,
            "temperature": 0.7,
            "max_tokens": 2000
        }

        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()  # 检查HTTP状态码
            response_json = response.json()
            return response_json['choices'][0]['message']['content']
        except requests.exceptions.Timeout:
            app.logger.error('API请求超时')
            raise Exception('API请求超时，请稍后重试')
        except requests.exceptions.RequestException as e:
            app.logger.error(f'API请求错误: {str(e)}')
            raise Exception(f'API请求失败: {str(e)}')
        except (KeyError, ValueError, TypeError) as e:
            app.logger.error(f'API响应解析错误: {str(e)}')
            raise Exception('API响应格式错误')
    except Exception as e:
        app.logger.error(f'DeepSeek API调用错误: {str(e)}', exc_info=True)
        raise Exception(f'API调用失败: {str(e)}')



@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传，保存文件，并初始化会话"""
    # 确保uploads目录存在
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        try:
            os.makedirs(app.config['UPLOAD_FOLDER'])
            app.logger.info(f"Created uploads directory at {app.config['UPLOAD_FOLDER']}")
        except Exception as e:
            app.logger.error(f"Failed to create uploads directory: {str(e)}")
            return jsonify({'error': '服务器配置错误，请联系管理员'}), 500

    if 'file' not in request.files:
        app.logger.warning('No file part in the request')
        return jsonify({'error': '请选择要上传的文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400

    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传Excel文件（.xlsx或.xls格式）'}), 400

    api_choice = request.form.get('api', 'deepseek')  # 获取用户选择的 API, 默认 deepseek
    if api_choice not in API_CHOICES:
        return jsonify({'error': '无效的API选择'}), 400

    try:
        # 创建新的会话 ID
        session_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        file.save(file_path)
        
        # 验证文件是否保存成功
        if not os.path.exists(file_path):
            return jsonify({'error': '文件保存失败，请重试'}), 500

        try:
            # 读取Excel文件内容
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            
            # 获取文件基本信息
            preview_info = []
            preview_info.append(f"文件名：{filename}")
            preview_info.append(f"工作表数量：{len(sheet_names)}")
            preview_info.append("")
            
            # 遍历每个工作表并获取信息
            dfs = {}
            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                dfs[sheet_name] = df
                
                preview_info.append(f"工作表：{sheet_name}")
                preview_info.append(f"总行数：{len(df)}")
                preview_info.append(f"列名：{', '.join(df.columns.tolist())}")
                preview_info.append("数据预览（前5行）：")
                
                # 格式化数据预览
                sample_data = df.head().to_string(index=False)
                preview_info.append(sample_data)
                preview_info.append("")
            
            # 构建系统消息
            file_info = "\n".join(preview_info)
            conversation_history[session_id] = {
                'api': api_choice,
                'history': [
                    {"role": "system", "content": f"用户上传了一个Excel文件，以下是文件信息：\n{file_info}"}
                ],
                'file_uploaded': True
            }

            return jsonify({
                'message': '文件上传成功',
                'session_id': session_id,
                'api': api_choice,
                'preview': file_info
            }), 200

        except Exception as e:
            # 清理失败的文件
            if os.path.exists(file_path):
                os.remove(file_path)
            app.logger.error(f'文件读取错误: {str(e)}')
            return jsonify({'error': f'文件读取失败: {str(e)}'}), 400

    except PermissionError as e:
        app.logger.error(f'文件权限错误: {str(e)}')
        return jsonify({'error': '服务器文件权限错误，请联系管理员'}), 500
    except Exception as e:
        app.logger.error(f'文件上传错误: {str(e)}', exc_info=True)
        return jsonify({'error': '服务器内部错误，请稍后重试'}), 500

@app.route('/chat', methods=['POST'])
def chat():
    """处理用户消息，调用 AI API，返回结果"""
    try:
        session_id = request.json.get('session_id')
        user_message = request.json.get('message')
    except Exception as e:
        app.logger.error(f'Invalid request format: {str(e)}', exc_info=True)
        return jsonify({'error': '无效的请求格式'}), 400

    if not session_id or not user_message:
        return jsonify({'error': 'Missing session_id or message'}), 400

    if session_id not in conversation_history or not conversation_history[session_id].get('file_uploaded'):
        return jsonify({'error': '无效会话ID或文件未上传'}), 400

    api_choice = conversation_history[session_id]['api']  # 获取 API 选择

    try:
        # 读取Excel文件并转换为CSV格式
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        if not os.path.exists(filepath):
            return jsonify({'error': '文件不存在，请重新上传'}), 404

        excel_file = pd.ExcelFile(filepath, engine='openpyxl')
        sheet_names = excel_file.sheet_names
        csv_data = {}

        for sheet_name in sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            csv_data[sheet_name] = csv_buffer.getvalue()

        # 构建包含CSV数据的系统消息
        system_message = "你是一个擅长处理Excel数据的助手。你的任务是直接处理数据并返回处理后的结果，而不是提供处理代码。\n\n请按以下格式返回：\n1. 首先用自然语言描述你的处理思路和结果\n2. 然后在 <CSV_DATA> 标签中返回处理后的CSV格式数据\n\n以下是用户上传的Excel文件（CSV格式）的内容：\n"
        for sheet_name, data in csv_data.items():
            system_message += f"\n工作表：{sheet_name}\n{data}\n"
        system_message += "\n请根据以上数据内容和用户需求，直接处理数据并返回处理结果。记住要在回复中包含 <CSV_DATA> 标签，标签内放置处理后的CSV格式数据。"
       
        # 将用户消息添加到会话历史
        conversation_history[session_id]['history'].append({"role": "user", "content": user_message})

        # 构建 AI API 请求
        messages = [
            {"role": "system", "content": system_message}
        ] + conversation_history[session_id]['history']

        ai_message = call_deepseek_api(messages, deepseek_api_key)

        # 将 AI 消息添加到会话历史
        conversation_history[session_id]['history'].append({"role": "assistant", "content": ai_message})

        # 解析AI返回的消息，提取CSV数据
        csv_data_start = ai_message.find('<CSV_DATA>')
        csv_data_end = ai_message.find('</CSV_DATA>')
        
        if csv_data_start != -1 and csv_data_end != -1:
            csv_data = ai_message[csv_data_start + len('<CSV_DATA>'):csv_data_end].strip()
            
            # 保存CSV数据到文件，使用UTF-8-BOM编码
            csv_filename = f"{session_id}_result.csv"
            csv_filepath = os.path.join(app.config['UPLOAD_FOLDER'], csv_filename)
            
            with open(csv_filepath, 'w', newline='', encoding='utf-8-sig') as f:
                f.write(csv_data)
            
            # 将CSV转换为XLSX
            df = pd.read_csv(csv_filepath, encoding='utf-8-sig')
            xlsx_filename = f"{session_id}_result.xlsx"
            xlsx_filepath = os.path.join(app.config['UPLOAD_FOLDER'], xlsx_filename)
            
            # 使用ExcelWriter保存为XLSX格式
            with pd.ExcelWriter(xlsx_filepath, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                
                # 自动调整列宽
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
            
            # 返回处理结果和文件下载链接
            return jsonify({
                'message': ai_message[:csv_data_start].strip(),
                'csv_data': csv_data,
                'download_url': f'/download/{session_id}'
            }), 200
        else:
            return jsonify({'message': ai_message}), 200

    except FileNotFoundError as e:
        app.logger.error(f'File not found error in chat: {str(e)}')
        return jsonify({'error': '文件不存在，请重新上传'}), 404
    except Exception as e:
        app.logger.error(f'Chat processing error: {str(e)}', exc_info=True)
        return jsonify({'error': '服务器内部错误，请稍后重试'}), 500

@app.route('/download/<session_id>')
def download_file(session_id):
    """处理XLSX文件下载请求"""
    try:
        xlsx_filename = f"{session_id}_result.xlsx"
        xlsx_filepath = os.path.join(app.config['UPLOAD_FOLDER'], xlsx_filename)
        
        if os.path.exists(xlsx_filepath):
            return send_file(
                xlsx_filepath,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='processed_data.xlsx'
            )
        else:
            return jsonify({'error': '文件不存在'}), 404
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def cleanup_temp_files(session_id):
    """清理会话相关的临时文件"""
    try:
        # 清理Excel文件
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        if os.path.exists(excel_path):
            os.remove(excel_path)
            
        # 清理CSV和XLSX文件
        for file in os.listdir(app.config['UPLOAD_FOLDER']):
            if file.startswith(session_id) and (file.endswith('.csv') or file.endswith('.xlsx')):
                os.remove(os.path.join(app.config['UPLOAD_FOLDER'], file))
    except Exception as e:
        app.logger.error(f'Error cleaning up temp files for session {session_id}: {str(e)}')

@app.route('/process', methods=['POST'])
def process_data():
    """执行数据处理 (Pandas + AI)，生成 Excel 文件"""
    session_id = request.json.get('session_id')

    if not session_id:
        return jsonify({'error': 'Missing session_id'}), 400

    if session_id not in conversation_history or not conversation_history[session_id].get('file_uploaded'):
        return jsonify({'error': '无效会话ID或文件未上传'}), 400

    try:
        # 从会话历史中获取所有的 AI 消息
        all_ai_messages = [msg['content'] for msg in conversation_history[session_id]['history'] if msg['role'] == 'assistant']
        pandas_code = all_ai_messages[-1] if all_ai_messages else ""  # 获取最后一条 AI 生成的代码

        # 读取 Excel 文件（支持多种格式）
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        
        if not os.path.exists(filepath):
            return jsonify({'error': '文件不存在，请重新上传'}), 404
            
        try:
            # 读取所有工作表并转换为CSV
            excel_file = pd.ExcelFile(filepath, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            dfs = {}
            csv_files = {}
            
            for sheet_name in sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
                dfs[sheet_name] = df
                
                # 为每个工作表创建一个临时CSV文件，并进行编码处理
                csv_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{sheet_name}.csv")
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                
                # 读取CSV文件并检测编码
                with open(csv_path, 'rb') as f:
                    raw_data = f.read()
                    detected = chardet.detect(raw_data)
                    encoding = detected['encoding']
                
                # 使用检测到的编码重新读取CSV文件
                df = pd.read_csv(csv_path, encoding=encoding)
                dfs[sheet_name] = df
                csv_files[sheet_name] = csv_path
        except Exception as e:
            app.logger.error(f'Session {session_id} file read error: {str(e)}')
            return jsonify({'error': f'文件读取失败: {str(e)}'}), 400

        # 执行 Pandas 代码 (使用 try-except 捕获错误)
        try:
            # 创建一个本地变量环境
            local_vars = {'dfs': dfs, 'pd': pd, 'csv_files': csv_files}
            app.logger.info(f'Processing session {session_id}\nExecuting code:\n{pandas_code}')
            exec(pandas_code, {}, local_vars)
            dfs = local_vars['dfs']
            
            # 清理临时CSV文件
            for csv_path in csv_files.values():
                if os.path.exists(csv_path):
                    os.remove(csv_path)

            # 将处理后的数据保存为CSV格式
            processed_csv_files = {}
            for sheet_name, df in dfs.items():
                csv_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_processed_{sheet_name}.csv")
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                processed_csv_files[sheet_name] = csv_path

            # 从CSV读取数据并转换为XLSX格式
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, csv_path in processed_csv_files.items():
                    df = pd.read_csv(csv_path, encoding='utf-8-sig')
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    
                    # 自动调整列宽
                    for idx, col in enumerate(df.columns):
                        max_length = max(
                            df[col].astype(str).apply(len).max(),
                            len(str(col))
                        )
                        worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

            # 清理处理后的CSV文件
            for csv_path in processed_csv_files.values():
                if os.path.exists(csv_path):
                    os.remove(csv_path)

            output.seek(0)
        except Exception as e:
            app.logger.error(f'Session {session_id} failed:\nCode:\n{pandas_code}\nError: {str(e)}', exc_info=True)
            return jsonify({'error': f'Error executing Pandas code: {str(e)}'}), 500

        # 返回文件给客户端
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name='processed_data.xlsx',
            as_attachment=True
        )

    except FileNotFoundError:
        return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        print(f"Error processing data: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)