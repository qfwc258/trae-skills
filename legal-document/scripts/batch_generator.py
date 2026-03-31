#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量文档生成器

支持 JSON/Excel 格式批量导入数据，一次性生成多份文书

用法:
    python batch_generator.py <template_key> <data_file>

示例:
    python batch_generator.py criminal_meeting data.json
    python batch_generator.py investigation data.xlsx
"""

import sys
import os
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from document_base import (
    create_simple_document,
    setup_logging,
    get_output_dir
)


TEMPLATE_CONFIG = {
    'criminal_meeting': {
        'name': '刑事法援会见前打印',
        'template': '8 刑事法援会见前打印.docx',
        'output_suffix': '刑事法援会见前打印'
    },
    'investigation': {
        'name': '调证材料',
        'template': '2 9 调证材料.docx',
        'output_suffix': '调证材料'
    },
    'evidence': {
        'name': '证据目录',
        'template': '3 证据目录 .docx',
        'output_suffix': '证据目录'
    }
}


def load_json_data(file_path):
    """加载 JSON 格式数据"""
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if isinstance(data, list):
        return data
    elif isinstance(data, dict) and 'items' in data:
        return data['items']
    else:
        return [data]


def load_excel_data(file_path):
    """加载 Excel 格式数据（需要 openpyxl）"""
    try:
        import openpyxl
    except ImportError:
        print('❌ 需要安装 openpyxl: pip install openpyxl')
        return []

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    headers = []
    data = []

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(h) if h else f'col_{j}' for j, h in enumerate(row)]
        else:
            row_data = {}
            for j, value in enumerate(row):
                if j < len(headers):
                    row_data[headers[j]] = value
            if any(row_data.values()):
                data.append(row_data)

    return data


def batch_generate(template_key, data_list, output_base_name=None):
    """批量生成文档"""
    if template_key not in TEMPLATE_CONFIG:
        print(f'❌ 未知模板类型: {template_key}')
        return []

    config = TEMPLATE_CONFIG[template_key]
    if output_base_name is None:
        output_base_name = config['output_suffix']

    output_dir = get_output_dir()
    results = []

    print(f'\n📄 开始批量生成: {config["name"]}')
    print(f'   数据条数: {len(data_list)}')
    print(f'   输出目录: {output_dir}')
    print()

    for i, data in enumerate(data_list):
        safe_name = str(data.get('weitr', f'item_{i+1}')).replace('/', '_').replace('\\', '_')
        output_name = f'{output_base_name}_{safe_name}.docx'

        try:
            output_path = create_simple_document(
                config['template'],
                output_name,
                data
            )
            results.append({'status': 'success', 'path': output_path, 'data': data})
            print(f'  ✅ [{i+1}/{len(data_list)}] {output_name}')
        except Exception as e:
            results.append({'status': 'error', 'error': str(e), 'data': data})
            print(f'  ❌ [{i+1}/{len(data_list)}] {output_name}: {e}')

    success_count = len([r for r in results if r['status'] == 'success'])
    print(f'\n✅ 批量生成完成: {success_count}/{len(data_list)} 成功')

    return results


def main():
    parser = argparse.ArgumentParser(description='批量生成法律文书')
    parser.add_argument('template_key', help='模板类型 (如 criminal_meeting)')
    parser.add_argument('data_file', help='数据文件 (JSON 或 Excel 格式)')
    parser.add_argument('-o', '--output', help='输出文件基础名')

    args = parser.parse_args()

    if not os.path.exists(args.data_file):
        print(f'❌ 数据文件不存在: {args.data_file}')
        return

    if args.data_file.endswith('.json'):
        data_list = load_json_data(args.data_file)
    elif args.data_file.endswith(('.xlsx', '.xls')):
        data_list = load_excel_data(args.data_file)
    else:
        print('❌ 仅支持 JSON 或 Excel 格式')
        return

    if not data_list:
        print('❌ 数据为空')
        return

    setup_logging()
    batch_generate(args.template_key, data_list, args.output)


if __name__ == '__main__':
    main()