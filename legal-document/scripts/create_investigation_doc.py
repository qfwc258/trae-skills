#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调证材料生成器 - 快速生成律师调查取证材料

用法:
    python create_investigation_doc.py [委托人] [身份证号] [对方当事人] [调取单位] [调取内容]

示例:
    python create_investigation_doc.py 张三 430124198203045123 李四 宁乡市公安局 "李四的户籍信息：430282195607084321"
"""

import sys
from docx import Document
import shutil
import os


def create_investigation_doc(wtr_name, wtr_sfz, dfr_name, dcdw_name, bdqxx_content):
    """
    生成调证材料文档

    Args:
        wtr_name: 委托人姓名
        wtr_sfz: 委托人身份证号
        dfr_name: 对方当事人
        dcdw_name: 调取单位
        bdqxx_content: 被调取信息内容
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    template_path = os.path.join(project_dir, 'legal_templates', '2 9 调证材料.docx')
    output_dir = os.getcwd()

    safe_wtr = wtr_name.replace('/', '_').replace('\\', '_')
    safe_dcdw = dcdw_name.replace('/', '_').replace('\\', '_')
    output_path = os.path.join(output_dir, f'调证材料_{safe_wtr}_{safe_dcdw}.docx')

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            output_path = os.path.join(output_dir, f'调证材料_{safe_wtr}_{safe_dcdw}_new.docx')

    shutil.copy(template_path, output_path)
    doc = Document(output_path)

    anyou = "调查取证"

    for i, para in enumerate(doc.paragraphs):
        full_text = para.text

        if 'weitr' in full_text:
            for j, run in enumerate(para.runs):
                if 'weitr' in run.text:
                    run.text = run.text.replace('weitr', wtr_name)
                if 'wtrsfz' in run.text:
                    run.text = run.text.replace('wtrsfz', wtr_sfz)
                if 'dfdsr' in run.text:
                    run.text = run.text.replace('dfdsr', dfr_name)
                if 'anyou' in run.text:
                    run.text = run.text.replace('anyou', anyou)

        if 'dcdw' in full_text:
            for j, run in enumerate(para.runs):
                if 'dcdw' in run.text:
                    run.text = run.text.replace('dcdw', dcdw_name)

        if 'bdqxx' in full_text:
            for j, run in enumerate(para.runs):
                if 'bdqxx' in run.text:
                    run.text = run.text.replace('bdqxx', bdqxx_content)

    doc.save(output_path)
    print(f'✅ 调证材料已生成: {output_path}')
    return output_path


if __name__ == '__main__':
    if len(sys.argv) >= 6:
        _, wtr, sfz, dfr, dcdw, bdqxx = sys.argv
        create_investigation_doc(wtr, sfz, dfr, dcdw, bdqxx)
    else:
        print(__doc__)
        print("\n📝 使用默认数据生成调证材料...")
        print("-" * 50)

        wtr = "张三"
        sfz = "430124198203045123"
        dfr = "李四"
        dcdw = "宁乡市公安局"
        bdqxx = "李四的户籍信息：430282195607084321"

        print(f"\n使用数据:\n委托人: {wtr}\n身份证: {sfz}\n对方: {dfr}\n调取单位: {dcdw}\n调取内容: {bdqxx}")
        print("-" * 50)

        create_investigation_doc(wtr, sfz, dfr, dcdw, bdqxx)
