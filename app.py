import os
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import xlrd
import xlwt
from xlutils.copy import copy
import numpy as np

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx'}

# 检查文件扩展名是否合法
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# 判断位置类型
def get_position_type(pos):
    try:
        if pos.startswith('+'):
            pos = pos[1:]
        value = float(pos[1:])
        if value < 10.0:
            return 'roof'
        elif 10.0 <= value < 60.0:
            return 'in'
        else:
            return 'bottom'
    except (ValueError, IndexError):
        return 'unknown'

# 根据位置类型和条件筛选数据
def filter_data(data, start_type, end_type, mvb_condition):
    return data[
        (data['起始位置'].apply(lambda x: get_position_type(x) == start_type)) &
        (data['终止位置'].apply(lambda x: get_position_type(x) == end_type)) &
        (data['点位1'] != 'free') &
        (data['点位2'] != 'FREE') &
        (data['线径'] != mvb_condition)
    ]

# 处理 Excel 文件
def process_file(file_path, mvb_condition):
    try:
        # 使用 xlrd 读取上传的 .xls 文件
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = workbook.sheet_by_index(0)  # 默认读取第一个工作表
        
        # 将数据转换为 DataFrame
        data = pd.DataFrame(sheet._cell_values[1:], columns=sheet.row_values(0))
        
        # 检查必要列是否存在
        required_columns = ['起始位置', '终止位置', '连接点1', '连接点2', '点位1', '点位2', '线径', '说明1', '说明2']
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            return None, f"文件缺少以下必要列: {', '.join(missing_columns)}，请检查文件并重新上传。"
        
        # 调换连接点位置
        mask = (data['连接点1'].str.contains('=99-XT', na=False)) & \
               (data['连接点2'].str.contains('=99-XT', na=False))
        data.loc[mask, ['起始位置', '连接点1', '点位1', '说明1', '终止位置', '连接点2', '点位2', '说明2']] = \
            data.loc[mask, ['终止位置', '连接点2', '点位2', '说明2', '起始位置', '连接点1', '点位1', '说明1']].values
        
        # 使用模板文件创建副本
        template_path = '校线表模板.xls'  # 模板文件位于与app.py相同的目录下
        template_workbook = xlrd.open_workbook(template_path, formatting_info=True)
        writable_workbook = copy(template_workbook)
        writable_sheet = writable_workbook.get_sheet(0)
        
        # 获取模板的表头
        template_headers = template_workbook.sheet_by_index(0).row_values(0)
        
        # 确保数据列与模板表头对齐
        data = data.reindex(columns=template_headers)
        
        # 定义位置类型对及其标题
        position_pairs = [
            ('roof', 'roof', '车顶对车顶'),
            ('roof', 'in', '车顶对车内'),
            ('roof', 'bottom', '车顶对车下'),
            ('in', 'in', '车内对车内'),
            ('in', 'bottom', '车内对车下'),
            ('bottom', 'bottom', '车下对车下')
        ]
        
        # 筛选并整理数据
        filtered_data = []
        optional_columns = ['接线工位1', '记录', '接线工位2']
        available_columns = [col for col in optional_columns if col in data.columns]
        columns = available_columns + ['起始位置', '连接点1', '点位1', '线号', '线径', '颜色', 
                                       '线束号', '终止位置', '连接点2', '点位2', '说明1', '说明2', '备注']
        
        for start_type, end_type, title in position_pairs:
            temp_data = filter_data(data, start_type, end_type, mvb_condition)
            temp_data = temp_data[columns].sort_values(
                by=['起始位置', '终止位置', '连接点1', '连接点2', '点位1', '点位2'])
            if not temp_data.empty:
                # 插入标题行，确保标题在第0列
                insert_row = pd.DataFrame({columns[0]: [title]}, columns=columns)
                filtered_data.append(insert_row)
                filtered_data.append(temp_data)
        
        # 合并数据
        if filtered_data:
            final_data = pd.concat(filtered_data, ignore_index=True)
        else:
            final_data = pd.DataFrame(columns=columns)
        
        # 设置字体样式为仿宋
        font_style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = '仿宋'
        font_style.font = font
        
        # 写入数据到副本，从第二行开始（第一行是表头）
        for r_idx, row in final_data.iterrows():
            for c_idx, value in enumerate(row):
                # 如果值是NaN或空，写入空字符串
                writable_sheet.write(r_idx + 1, c_idx, '' if pd.isna(value) else value, font_style)
        
        # 保存副本
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], '校线表_更新.xls')
        writable_workbook.save(save_path)
        
        return save_path, None
    except ValueError as e:
        return None, "文件格式错误，请确保上传的是正确的 Excel 文件。"
    except Exception as e:
        return None, f"处理文件时发生错误: {str(e)}，请检查文件并重新上传。"

# 主页面路由
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # 检查文件和 mvb_condition 是否提供
        if 'file' not in request.files or 'mvb_condition' not in request.form:
            return render_template('upload.html', error='请上传文件并输入 mvb_condition')
        
        file = request.files['file']
        mvb_condition = request.form['mvb_condition']
        
        if file.filename == '' or mvb_condition == '':
            return render_template('upload.html', error='请上传文件并输入 mvb_condition')
        
        # 检查文件类型
        if not allowed_file(file.filename):
            return render_template('upload.html', error='文件类型错误，请上传 Excel 文件（.xls 或 .xlsx）')
        
        # 保存文件
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # 处理文件
        save_path, error = process_file(file_path, mvb_condition)
        if error:
            return render_template('upload.html', error=error)
        return redirect(url_for('download_file', filename='校线表_更新.xls'))
    
    return render_template('upload.html')

# 下载文件路由
@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == "__main__":
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=False, host="0.0.0.0", port=5001)