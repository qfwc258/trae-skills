#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交互式法律文书生成器

提供向导式问答，依次引导用户输入必要信息生成文书

用法:
    python interactive_generator.py [文书类型]

示例:
    python interactive_generator.py criminal_meeting
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from document_base import (
    create_simple_document,
    get_template_dir,
    load_config,
    setup_logging
)


TEMPLATE_INFO = {
    'criminal_meeting': {
        'name': '刑事法援会见前打印',
        'template': '8 刑事法援会见前打印.docx',
        'output_suffix': '刑事法援会见前打印',
        'placeholders': {
            'weitr': {'prompt': '受援人姓名', 'required': True},
            'wtrsfz': {'prompt': '受援人身份证号码', 'required': True},
            'wtrzz': {'prompt': '住所/羁押地', 'required': True},
            'anyou': {'prompt': '涉嫌罪名', 'required': True},
            'jied': {'prompt': '案件阶段（侦查/审查起诉/审判）', 'required': True},
            'wtxm': {'prompt': '委托项目编号（默认：一）', 'required': False, 'default': '一'},
            'quanx': {'prompt': '代理权限（默认：一般代理）', 'required': False, 'default': '一般代理'},
            'hjdd': {'prompt': '会见地点', 'required': False}
        }
    },
    'investigation': {
        'name': '调证材料',
        'template': '2 9 调证材料.docx',
        'output_suffix': '调证材料',
        'placeholders': {
            'weitr': {'prompt': '委托人姓名', 'required': True},
            'wtrsfz': {'prompt': '委托人身份证号码', 'required': True},
            'dfdsr': {'prompt': '对方当事人', 'required': True},
            'anyou': {'prompt': '案由/事项', 'required': True},
            'dcdw': {'prompt': '调取单位', 'required': True},
            'bdqxx': {'prompt': '被调取信息内容', 'required': True}
        }
    },
    'evidence': {
        'name': '证据目录',
        'template': '3 证据目录 .docx',
        'output_suffix': '证据目录',
        'placeholders': {
            'weitr': {'prompt': '提交人姓名', 'required': True}
        }
    },
    'civil_directory': {
        'name': '民事法律援助案卷归档目录',
        'template': '9 湖南省法律援助案卷归档目录（民事）.docx',
        'output_suffix': '民事法援案卷归档目录',
        'placeholders': {
            '宁乡市': {'prompt': '城市名称（默认：宁乡市）', 'required': False, 'default': '宁乡市'}
        }
    },
    'criminal_directory': {
        'name': '刑事法律援助案卷归档目录',
        'template': '8 湖南省法律援助案卷归档目录（刑事）.docx',
        'output_suffix': '刑事法援案卷归档目录',
        'placeholders': {
            '宁乡市': {'prompt': '城市名称（默认：宁乡市）', 'required': False, 'default': '宁乡市'}
        }
    }
}


def print_banner():
    print('=' * 60)
    print('          法律文书交互式生成器')
    print('=' * 60)
    print()


def print_template_list():
    print('可选文书类型：')
    print('-' * 40)
    for key, info in TEMPLATE_INFO.items():
        print(f'  {key:20s} - {info["name"]}')
    print()


def get_input_with_default(prompt, default=None):
    """获取用户输入，支持默认值"""
    if default:
        user_input = input(f'{prompt} [{default}]: ').strip()
        return user_input if user_input else default
    else:
        return input(f'{prompt}: ').strip()


def collect_placeholders(placeholders_config):
    """收集占位符信息"""
    data = {}
    for placeholder, config in placeholders_config.items():
        value = get_input_with_default(
            config['prompt'],
            config.get('default')
        )
        if config.get('required') and not value:
            print(f'❌ 此项为必填项，请重新输入')
            return collect_placeholders(placeholders_config)
        data[placeholder] = value
    return data


def generate_document(template_key, data):
    """生成文档"""
    config = TEMPLATE_INFO[template_key]

    output_name = f'{config["output_suffix"]}_{data.get("weitr", data.get("宁乡市", "文档"))}.docx'
    output_name = output_name.replace('/', '_').replace('\\', '_')

    try:
        output_path = create_simple_document(
            config['template'],
            output_name,
            data
        )
        return output_path
    except Exception as e:
        raise e


def main():
    print_banner()
    setup_logging()

    if len(sys.argv) < 2:
        print_template_list()
        template_key = input('请选择文书类型: ').strip()
    else:
        template_key = sys.argv[1].strip()

    if template_key not in TEMPLATE_INFO:
        print(f'❌ 未知文书类型: {template_key}')
        print_template_list()
        return

    config = TEMPLATE_INFO[template_key]
    print(f'\n📝 开始生成: {config["name"]}')
    print('-' * 40)

    data = collect_placeholders(config['placeholders'])

    print('\n📋 请确认信息:')
    for key, value in data.items():
        print(f'  {key}: {value}')
    print()

    confirm = input('确认生成? (y/n): ').strip().lower()
    if confirm != 'y':
        print('已取消生成')
        return

    try:
        output_path = generate_document(template_key, data)
        print(f'\n✅ 文档已生成: {output_path}')
    except Exception as e:
        print(f'\n❌ 生成失败: {e}')


if __name__ == '__main__':
    main()