#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
准备法律援助结案模板 - 在原文档中添加占位符
"""

import zipfile
import os
import shutil
import re

def unpack_docx(docx_path, extract_dir):
    """解压docx文件"""
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def pack_docx(extract_dir, output_path):
    """打包为docx文件"""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zipf.write(file_path, arcname)

def add_placeholders_to_xml(xml_path):
    """在XML中添加占位符"""
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 定义需要替换的文本模式 -> 占位符
    patterns = [
        # 承办情况通报表头
        (r'(<w:t[^>]*>)(受援人：)(</w:t>)', r'\g<1>受援人：weitr\g<3>'),
        (r'(<w:t[^>]*>)(承办人：陈伟  案件编号：)(</w:t>)', r'\g<1>承办人：陈伟  案件编号：bianh\g<3>'),
        
        # 表格中的空白单元格添加占位符
        # 序号1-9行
        (r'(<w:t[^>]*>)(1)(</w:t>)(.*?<w:t[^>]*>)(</w:t>)', r'\g<1>1\g<3>\g<4>\g<1>gcsj1\g<3>'),
    ]
    
    # 更简单的方法：直接在特定位置插入占位符文本
    # 找到表格结构，在空白单元格中添加占位符
    
    # 保存修改后的内容
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return content

def create_template_with_placeholders(input_docx, output_docx):
    """创建带占位符的模板"""
    work_dir = os.path.join(os.path.dirname(output_docx), 'temp_template')
    
    # 解压
    print(f"解压: {input_docx}")
    unpack_docx(input_docx, work_dir)
    
    # 修改document.xml
    doc_xml = os.path.join(work_dir, 'word', 'document.xml')
    if os.path.exists(doc_xml):
        print("添加占位符...")
        add_placeholders_to_xml(doc_xml)
    
    # 打包
    print(f"生成模板: {output_docx}")
    pack_docx(work_dir, output_docx)
    
    # 清理
    shutil.rmtree(work_dir)
    print("完成!")

def main():
    input_dir = r"d:\trae\法律援助文书"
    input_docx = os.path.join(input_dir, "8 9 法援结案.docx")
    output_docx = os.path.join(input_dir, "法援结案_模板.docx")
    
    create_template_with_placeholders(input_docx, output_docx)

if __name__ == "__main__":
    main()
