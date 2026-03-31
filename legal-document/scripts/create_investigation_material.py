#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成调证材料文档
"""

from docx import Document
import shutil
import os


def create_investigation_material():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    template_path = os.path.join(project_dir, 'legal_templates', '2 9 调证材料.docx')
    output_dir = os.getcwd()
    output_path = os.path.join(output_dir, '调证材料_张三委托陈伟律师_宁乡市公安局.docx')
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            output_path = os.path.join(output_dir, '调证材料_张三委托陈伟律师_宁乡市公安局_new.docx')

    shutil.copy(template_path, output_path)
    doc = Document(output_path)

    for i, para in enumerate(doc.paragraphs):
        full_text = para.text

        if 'weitr' in full_text:
            for j, run in enumerate(para.runs):
                if 'weitr' in run.text:
                    run.text = run.text.replace('weitr', '张三')
                if 'wtrsfz' in run.text:
                    run.text = run.text.replace('wtrsfz', '430124198203045123')
                if 'dfdsr' in run.text:
                    run.text = run.text.replace('dfdsr', '李四')
                if 'anyou' in run.text:
                    run.text = run.text.replace('anyou', '户籍信息查询')

        if 'dcdw' in full_text or ('dc' in full_text and 'dw' in full_text):
            for j, run in enumerate(para.runs):
                if 'dc' in run.text:
                    run.text = run.text.replace('dc', '宁乡市公安局')
                if 'dw' in run.text and '宁乡市公安局' not in run.text:
                    run.text = run.text.replace('dw', '')

        if 'bdqxx' in full_text:
            for j, run in enumerate(para.runs):
                if 'bdqxx' in run.text:
                    run.text = run.text.replace('bdqxx', '李四的户籍信息：430282195607084321')

    doc.save(output_path)
    print(f'调证材料已生成: {output_path}')


if __name__ == '__main__':
    create_investigation_material()
