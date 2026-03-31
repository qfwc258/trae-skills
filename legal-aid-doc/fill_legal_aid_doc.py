#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
法律援助结案文档填充工具 - 保持原格式
读取原文档XML，直接替换占位符，保持所有格式不变
"""

import zipfile
import os
import shutil
import re
from pathlib import Path

def parse_element_file(file_path):
    """解析元素文件"""
    elements = {}
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 解析 --section-- 格式
    sections = re.split(r'\n--', content)
    for section in sections:
        section = section.strip()
        if not section:
            continue
        lines = section.split('\n')
        for line in lines:
            line = line.strip()
            if '::' in line:
                key, value = line.split('::', 1)
                elements[key.strip()] = value.strip()
    return elements

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

def replace_in_xml(xml_path, replacements):
    """在XML中替换文本，保持格式"""
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 替换所有占位符
    for key, value in replacements.items():
        # 匹配 <w:t>key</w:t> 或 <w:t xml:space="preserve">key</w:t>
        pattern = rf'(<w:t[^>]*>)({re.escape(key)})(</w:t>)'
        replacement = rf'\g<1>{value}\g<3>'
        content = re.sub(pattern, replacement, content)
    
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(content)

def fill_legal_aid_doc(template_path, element_file, output_path):
    """填充法律援助结案文档"""
    
    # 1. 解析元素文件
    print(f"读取元素文件: {element_file}")
    elements = parse_element_file(element_file)
    print(f"解析到 {len(elements)} 个元素")
    
    # 2. 准备替换映射
    replacements = {
        'bianh': elements.get('bianh', ''),
        'leib': elements.get('leib', '民事'),
        'anyou': elements.get('anyou', ''),
        'weitr': elements.get('weitr', ''),
        'dfdsr': elements.get('dfdsr', ''),
        'badw': elements.get('badw', ''),
        'jied': elements.get('jied', '一审'),
        'zprq': elements.get('zprq', ''),
        'jarq': elements.get('jarq', ''),
        'cbxj': elements.get('cbxj', ''),
        'ljsm': elements.get('ljsm', ''),
        'gdrq': elements.get('gdrq', ''),
        'yjrq': elements.get('yjrq', ''),
    }
    
    # 添加办案过程数据
    for i in range(1, 10):
        replacements[f'gcsj{i}'] = elements.get(f'gcsj{i}', '')
        replacements[f'gcfs{i}'] = elements.get(f'gcfs{i}', '')
        replacements[f'gcnr{i}'] = elements.get(f'gcnr{i}', '')
    
    # 3. 解压原文档
    work_dir = os.path.join(os.path.dirname(output_path), 'temp_work')
    print(f"解压模板: {template_path}")
    unpack_docx(template_path, work_dir)
    
    # 4. 替换document.xml中的内容
    doc_xml = os.path.join(work_dir, 'word', 'document.xml')
    if os.path.exists(doc_xml):
        print("替换文档内容...")
        replace_in_xml(doc_xml, replacements)
    
    # 5. 打包为新文档
    print(f"生成文档: {output_path}")
    pack_docx(work_dir, output_path)
    
    # 6. 清理临时文件
    shutil.rmtree(work_dir)
    
    print("完成!")
    return output_path

def main():
    input_dir = r"d:\trae\法律援助文书"
    template_path = os.path.join(input_dir, "8 9 法援结案.docx")
    element_file = os.path.join(input_dir, "民事_元素.txt")
    output_path = os.path.join(input_dir, "法援结案_填充后.docx")
    
    fill_legal_aid_doc(template_path, element_file, output_path)

if __name__ == "__main__":
    main()
