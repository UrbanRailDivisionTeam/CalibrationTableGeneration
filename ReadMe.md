# 自动校线表生成软件

## 项目简介

自动校线表生成软件是一个基于 Flask 的 Web 应用程序，旨在帮助用户上传 Excel 文件并生成校线表。用户可以通过简单的界面选择文件并输入 MVB 线型，系统将处理数据并生成更新后的校线表。

## 功能

- 上传 Excel 文件（支持 .xls 和 .xlsx 格式）
- 输入 MVB 线型
- 自动处理数据并生成校线表
- 下载生成的校线表

## 技术栈

- Python
- Flask
- Pandas
- xlrd, xlwt, xlutils
- HTML/CSS

## 安装与运行

1. 克隆项目到本地：

   ```bash
   git clone <项目地址>
   cd <项目目录>
   ```

2. 创建并激活虚拟环境（可选）：

   ```bash
   conda create -n tempapp2 python=3.8
   conda activate tempapp2
   ```

3. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```

4. 运行应用：

   ```bash
   python app.py
   ```

5. 打开浏览器，访问 `http://10.24.5.54:5001`。

## 使用说明

1. 在主页面，点击"选择 Excel 文件"按钮，上传需要处理的 Excel 文件。
2. 在输入框中输入 MVB 线型。
3. 点击"提交"按钮，系统将处理文件并生成校线表。
4. 处理完成后，您将被重定向到下载链接，点击下载生成的校线表。

## 贡献

欢迎任何形式的贡献！请提交问题或拉取请求。

## 许可证

本项目采用 GNU 通用公共许可证（GPL）进行许可。有关详细信息，请参阅 [LICENSE](LICENSE) 文件。
