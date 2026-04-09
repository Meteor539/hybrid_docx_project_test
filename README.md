# Word文档格式检查与修改软件

## 项目简介
本软件是一个基于Python的Word文档格式检查与修改工具，主要用于检查论文格式并自动修正。支持检查的内容包括封面、摘要、正文、各级标题、参考文献、参考文献引用格式等。

## 功能特点
- 支持docx文件的选择和拖拽
- 可自定义设置各部分格式要求
- 自动检查文档格式并显示检查结果
- 支持一键修改不符合要求的格式
- 自动保存修改后的文档

## 安装说明
1. 确保已安装Python 3.8或更高版本
2. 安装依赖包：
```bash
pip install -r requirements.txt
```

## 使用说明
1. 运行程序：
```bash
python main.py
```
2. 在界面中选择或拖拽需要检查的docx文件
3. 设置各部分格式要求
4. 点击"开始检查"按钮进行格式检查
5. 查看检查结果
6. 点击"确认修改"按钮进行格式修正

## 项目结构
```
.
├── main.py             # 程序入口
├── gui/                # GUI相关模块
	├──main_window.py
├── model/              # 业务功能模块
	├──document_parser.py
	├──format_checker.py
	├──format_modifier.py
└── requirements.txt    # 项目依赖列表(仅作记录)
```

## 技术栈
- Python
- python-docx
- PyQt6

## 开发环境
- Windows 11
- Python 3.13.2
- VSCode
