import markdown
import os

# 读取 Markdown 文件
md_file = '使用说明.md'
html_file = '使用说明.html'

with open(md_file, 'r', encoding='utf-8') as f:
    md_content = f.read()

# 转换 Markdown 为 HTML
html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code', 'toc'])

# 添加完整的 HTML 结构和 CSS 样式
full_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DGSS 区域地质调查野外记录簿一键整理工具 - 使用说明</title>
    <style>
        @media print {{
            @page {{
                size: A4;
                margin: 2cm;
            }}
            body {{
                font-size: 10pt;
            }}
        }}
        
        body {{
            font-family: "Microsoft YaHei", "SimSun", "PingFang SC", sans-serif;
            line-height: 1.8;
            color: #333;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            background-color: #fff;
        }}
        
        h1 {{
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 12px;
            margin-top: 40px;
            margin-bottom: 20px;
            font-size: 28px;
            font-weight: bold;
        }}
        
        h1:first-child {{
            margin-top: 0;
        }}
        
        h2 {{
            color: #34495e;
            border-bottom: 2px solid #95a5a6;
            padding-bottom: 10px;
            margin-top: 35px;
            margin-bottom: 15px;
            font-size: 22px;
            font-weight: bold;
        }}
        
        h3 {{
            color: #555;
            margin-top: 25px;
            margin-bottom: 12px;
            font-size: 18px;
            font-weight: bold;
        }}
        
        h4 {{
            color: #666;
            margin-top: 20px;
            margin-bottom: 10px;
            font-size: 16px;
            font-weight: bold;
        }}
        
        p {{
            margin: 10px 0;
            text-align: justify;
        }}
        
        code {{
            background-color: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: "Consolas", "Courier New", "Monaco", monospace;
            font-size: 0.9em;
            color: #e74c3c;
        }}
        
        pre {{
            background-color: #f8f8f8;
            padding: 15px;
            border-left: 4px solid #3498db;
            overflow-x: auto;
            border-radius: 4px;
            margin: 15px 0;
        }}
        
        pre code {{
            background-color: transparent;
            padding: 0;
            color: #333;
        }}
        
        ul, ol {{
            margin: 12px 0;
            padding-left: 35px;
        }}
        
        li {{
            margin: 8px 0;
            line-height: 1.6;
        }}
        
        strong {{
            color: #e74c3c;
            font-weight: bold;
        }}
        
        em {{
            color: #3498db;
        }}
        
        hr {{
            border: none;
            border-top: 2px solid #ecf0f1;
            margin: 40px 0;
        }}
        
        blockquote {{
            border-left: 4px solid #f39c12;
            padding-left: 15px;
            color: #666;
            background-color: #fef9f3;
            margin: 15px 0;
            padding: 12px 15px;
            font-style: italic;
        }}
        
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }}
        
        table, th, td {{
            border: 1px solid #ddd;
        }}
        
        th {{
            background-color: #3498db;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }}
        
        td {{
            padding: 10px;
        }}
        
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        
        /* 特殊标记样式 */
        .warning {{
            color: #e74c3c;
            font-weight: bold;
        }}
        
        /* 打印按钮 */
        .no-print {{
            text-align: center;
            margin: 20px 0;
            padding: 20px;
            background-color: #ecf0f1;
            border-radius: 8px;
        }}
        
        .print-btn {{
            background-color: #3498db;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            font-family: "Microsoft YaHei", sans-serif;
        }}
        
        .print-btn:hover {{
            background-color: #2980b9;
        }}
        
        @media print {{
            .no-print {{
                display: none;
            }}
        }}
    </style>
</head>
<body>
    <div class="no-print">
        <button class="print-btn" onclick="window.print()">🖨️ 打印或保存为 PDF</button>
        <p style="margin-top: 10px; color: #666; font-size: 14px;">
            点击按钮后，在打印对话框中选择"另存为 PDF"即可保存
        </p>
    </div>
    
    {html_content}
    
    <div class="no-print" style="margin-top: 40px;">
        <button class="print-btn" onclick="window.print()">🖨️ 打印或保存为 PDF</button>
    </div>
</body>
</html>"""

# 写入 HTML 文件
with open(html_file, 'w', encoding='utf-8') as f:
    f.write(full_html)

print(f"✅ HTML 文件已生成：{html_file}")
print(f"📌 请用浏览器打开此文件，然后使用 Ctrl+P 打印为 PDF")
print(f"   或直接点击页面中的\"打印或保存为 PDF\"按钮")

# 尝试用默认浏览器打开
os.startfile(html_file)
