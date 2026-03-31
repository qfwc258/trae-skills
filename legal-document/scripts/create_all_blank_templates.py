#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一空白模板生成器 - 一键生成所有空白模板

用法:
    python create_all_blank_templates.py [模板名称]

示例:
    python create_all_blank_templates.py all          # 生成所有空白模板
    python create_all_blank_templates.py criminal     # 生成刑事相关空白模板
    python create_all_blank_templates.py civil        # 生成民事相关空白模板
    python create_all_blank_templates.py meeting      # 生成会见前打印空白模板
"""

import os
import sys
from docx import Document
from document_base import (
    get_output_dir,
    get_template_dir,
    copy_template_to_output,
    replace_placeholders_in_doc,
    save_doc,
    BLANK_PLACEHOLDER
)


TEMPLATE_CONFIG = {
    'meeting': {
        'name': '刑事法援会见前打印',
        'template': '8 刑事法援会见前打印.docx',
        'output': '刑事法援会见前打印_空白模板.docx',
        'placeholders': {
            'weitr': '', 'wtrsfz': '', 'wtrzz': '', 'anyou': '',
            'jied': '', 'wtxm': '', 'quanx': '', 'hjdd': ''
        }
    },
    'investigation': {
        'name': '调证材料',
        'template': '2 9 调证材料.docx',
        'output': '调证材料_空白模板.docx',
        'placeholders': {
            'weitr': '', 'wtrsfz': '', 'dfdsr': '', 'anyou': '',
            'dcdw': '', 'bdqxx': ''
        }
    },
    'evidence': {
        'name': '证据目录',
        'template': '3 证据目录 .docx',
        'output': '证据目录_空白模板.docx',
        'placeholders': {'weitr': ''}
    },
    'civil_directory': {
        'name': '民事法律援助案卷归档目录',
        'template': '9 湖南省法律援助案卷归档目录（民事）.docx',
        'output': '法律援助案卷归档目录（民事）_空白模板.docx',
        'placeholders': {'宁乡市': ''}
    },
    'criminal_directory': {
        'name': '刑事法律援助案卷归档目录',
        'template': '8 湖南省法律援助案卷归档目录（刑事）.docx',
        'output': '法律援助案卷归档目录（刑事）_空白模板.docx',
        'placeholders': {'宁乡市': ''}
    }
}

CATEGORY_MAP = {
    'criminal': ['meeting', 'criminal_directory'],
    'civil': ['civil_directory', 'evidence', 'investigation'],
    'all': list(TEMPLATE_CONFIG.keys())
}


def create_blank_template(config_key):
    """生成单个空白模板"""
    config = TEMPLATE_CONFIG[config_key]
    template_path = os.path.join(get_template_dir(), config['template'])

    if not os.path.exists(template_path):
        print(f'❌ 模板文件不存在: {template_path}')
        return None

    output_path = copy_template_to_output(config['template'], config['output'])
    doc = Document(output_path)
    replace_placeholders_in_doc(doc, config['placeholders'])
    save_doc(doc, output_path)

    print(f'✅ {config["name"]}_空白模板已生成')
    return output_path


def create_category_blanks(category):
    """按类别生成空白模板"""
    keys = CATEGORY_MAP.get(category, [])
    if not keys:
        print(f'❌ 未知类别: {category}')
        return

    print(f'\n📄 生成 {category} 类空白模板...\n')
    for key in keys:
        create_blank_template(key)
    print(f'\n✅ {category} 类空白模板全部生成完成!')


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n可用类别: all, criminal, civil")
        print("\n按类别生成:")
        for cat, keys in CATEGORY_MAP.items():
            if cat != 'all':
                templates = [TEMPLATE_CONFIG[k]['name'] for k in keys]
                print(f"  {cat}: {', '.join(templates)}")
        return

    category = sys.argv[1].lower()
    create_category_blanks(category)


if __name__ == '__main__':
    main()