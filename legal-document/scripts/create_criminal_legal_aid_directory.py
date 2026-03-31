#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成刑事法律援助案卷归档目录
"""

from docx import Document
import shutil
import os


def create_criminal_legal_aid_directory(city_name="宁乡市"):
    """
    生成刑事法律援助案卷归档目录

    Args:
        city_name: 城市名称，默认"宁乡市"
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    template_path = os.path.join(project_dir, 'legal_templates', '8 湖南省法律援助案卷归档目录（刑事）.docx')
    output_dir = os.getcwd()
    output_path = os.path.join(output_dir, f'法律援助案卷归档目录（刑事）_{city_name}.docx')

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            output_path = os.path.join(output_dir, f'法律援助案卷归档目录（刑事）_{city_name}_new.docx')

    shutil.copy(template_path, output_path)
    doc = Document(output_path)

    for para in doc.paragraphs:
        if para.text.strip():
            for run in para.runs:
                if "宁乡市" in run.text:
                    run.text = run.text.replace("宁乡市", city_name)

    doc.save(output_path)
    print(f'✅ 刑事法律援助案卷归档目录已生成: {output_path}')
    return output_path


if __name__ == '__main__':
    import sys

    if len(sys.argv) > 1:
        city = sys.argv[1]
        create_criminal_legal_aid_directory(city)
    else:
        create_criminal_legal_aid_directory()
