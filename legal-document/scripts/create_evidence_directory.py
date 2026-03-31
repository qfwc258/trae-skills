#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证据目录生成器

用法:
    python create_evidence_directory.py [提交人]

示例:
    python create_evidence_directory.py 张三
"""

import sys
from docx import Document
import shutil
import os


def create_evidence_directory(weitr):
    """
    生成证据目录

    Args:
        weitr: 提交人姓名
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    template_path = os.path.join(project_dir, 'legal_templates', '3 证据目录 .docx')
    output_dir = os.getcwd()

    safe_name = weitr.replace('/', '_').replace('\\', '_')
    output_path = os.path.join(output_dir, f'证据目录_{safe_name}.docx')

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            output_path = os.path.join(output_dir, f'证据目录_{safe_name}_new.docx')

    shutil.copy(template_path, output_path)
    doc = Document(output_path)

    for para in doc.paragraphs:
        full_text = para.text
        if 'weitr' in full_text:
            for run in para.runs:
                if 'weitr' in run.text:
                    run.text = run.text.replace('weitr', '        ')

    doc.save(output_path)
    print(f'✅ 证据目录已生成: {output_path}')
    return output_path


if __name__ == '__main__':
    if len(sys.argv) > 1:
        weitr = sys.argv[1]
        create_evidence_directory(weitr)
    else:
        print(__doc__)
        print("\n📝 使用默认数据生成...")
        print("-" * 50)

        weitr = "张三"

        print(f"\n使用数据:\n提交人: {weitr}")
        print("-" * 50)

        create_evidence_directory(weitr)
