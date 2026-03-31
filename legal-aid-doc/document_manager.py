#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档管理模块 - 优化5、6、7
- 增量更新（只填充空白处）
- 输出组织（按案件类型分类）
- 日志与回滚
"""

import os
import json
import shutil
from datetime import datetime
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
from pathlib import Path


@dataclass
class FillLogEntry:
    """填充日志条目"""
    timestamp: str
    template_name: str
    output_name: str
    elements_used: Dict[str, str]
    fields_filled: List[str]
    fields_skipped: List[str]
    success: bool
    error_message: Optional[str] = None


class DocumentManager:
    """文档管理器 - 优化5、6、7"""
    
    def __init__(self, base_dir: str):
        """
        Args:
            base_dir: 基础目录
        """
        self.base_dir = base_dir
        self.output_dir = os.path.join(base_dir, "output")
        self.backup_dir = os.path.join(base_dir, "backup")
        self.log_dir = os.path.join(base_dir, "logs")
        
        # 确保目录存在
        for dir_path in [self.output_dir, self.backup_dir, self.log_dir]:
            os.makedirs(dir_path, exist_ok=True)
    
    # ==================== 优化5: 增量更新 ====================
    
    def is_field_empty(self, text: str) -> bool:
        """
        检查字段是否为空（只填充空白处）
        
        空白包括：
        - 空字符串
        - 只有占位符
        - 只有空白字符
        - 只有下划线
        - 只有"无"、"暂无"等占位文本
        """
        if not text:
            return True
        
        text = text.strip()
        
        # 空字符串
        if not text:
            return True
        
        # 只有占位符
        placeholder_patterns = [
            r'^\{[^}]+\}$',  # {key}
            r'^\[\[[^\]]+\]\]$',  # [[key]]
            r'^【[^】]+】$',  # 【key】
            r'^<[^>]+>$',  # <key>
            r'^#[^#]+#$',  # #key#
            r'^\$[^$]+\$$',  # $key$
            r'^%[^%]+%$',  # %key%
        ]
        for pattern in placeholder_patterns:
            if re.match(pattern, text):
                return True
        
        # 只有空白字符
        if text.replace(' ', '').replace('\t', '').replace('\n', '') == '':
            return True
        
        # 只有下划线
        if set(text) <= set('_'):
            return True
        
        # 占位文本
        placeholder_texts = ['无', '暂无', '待填写', '待补充', '待定', '未填写', '_____', '______']
        if text in placeholder_texts:
            return True
        
        return False
    
    def should_fill_field(self, current_value: str, new_value: str, 
                          incremental: bool = True) -> bool:
        """
        判断是否应该填充字段
        
        Args:
            current_value: 当前值
            new_value: 新值
            incremental: 是否增量模式
        
        Returns:
            是否应该填充
        """
        if not incremental:
            # 非增量模式：总是填充
            return True
        
        # 增量模式：只填充空白处
        if self.is_field_empty(current_value):
            return True
        
        # 当前值不为空，不填充
        return False
    
    # ==================== 优化6: 输出组织 ====================
    
    def get_output_path(self, template_path: str, case_type: str = "民事", 
                       suffix: str = "_已填充") -> str:
        """
        获取输出路径（按案件类型分类）
        
        输出结构：
        output/
          民事/
            法援结案_已填充.docx
            阅卷笔录_已填充.docx
          刑事/
            ...
        """
        # 创建案件类型目录
        case_dir = os.path.join(self.output_dir, case_type)
        os.makedirs(case_dir, exist_ok=True)
        
        # 生成输出文件名
        template_name = os.path.basename(template_path)
        name, ext = os.path.splitext(template_name)
        output_name = f"{name}{suffix}{ext}"
        
        return os.path.join(case_dir, output_name)
    
    def organize_output(self, source_files: List[str], case_type: str = "民事"):
        """
        组织输出文件
        
        Args:
            source_files: 源文件列表
            case_type: 案件类型
        """
        case_dir = os.path.join(self.output_dir, case_type)
        os.makedirs(case_dir, exist_ok=True)
        
        for source_file in source_files:
            if os.path.exists(source_file):
                dest_name = os.path.basename(source_file)
                dest_path = os.path.join(case_dir, dest_name)
                shutil.copy2(source_file, dest_path)
                print(f"  组织到 {case_type}/: {dest_name}")
    
    def list_outputs(self, case_type: Optional[str] = None) -> Dict[str, List[str]]:
        """
        列出输出文件
        
        Returns:
            {案件类型: [文件列表]}
        """
        result = {}
        
        if case_type:
            # 列出特定类型
            case_dir = os.path.join(self.output_dir, case_type)
            if os.path.exists(case_dir):
                result[case_type] = [
                    os.path.join(case_dir, f) 
                    for f in os.listdir(case_dir) 
                    if f.endswith('.docx')
                ]
        else:
            # 列出所有类型
            if os.path.exists(self.output_dir):
                for ctype in os.listdir(self.output_dir):
                    case_dir = os.path.join(self.output_dir, ctype)
                    if os.path.isdir(case_dir):
                        result[ctype] = [
                            os.path.join(case_dir, f)
                            for f in os.listdir(case_dir)
                            if f.endswith('.docx')
                        ]
        
        return result
    
    # ==================== 优化7: 日志与回滚 ====================
    
    def create_backup(self, file_path: str) -> Optional[str]:
        """
        创建文件备份
        
        Returns:
            备份文件路径
        """
        if not os.path.exists(file_path):
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = os.path.basename(file_path)
        backup_name = f"{file_name}.{timestamp}.bak"
        backup_path = os.path.join(self.backup_dir, backup_name)
        
        shutil.copy2(file_path, backup_path)
        return backup_path
    
    def rollback(self, file_path: str, backup_path: Optional[str] = None) -> bool:
        """
        回滚到备份版本
        
        Args:
            file_path: 原文件路径
            backup_path: 指定备份路径，如果不指定则使用最新的备份
        
        Returns:
            是否成功
        """
        if backup_path and os.path.exists(backup_path):
            shutil.copy2(backup_path, file_path)
            return True
        
        # 查找最新的备份
        file_name = os.path.basename(file_path)
        backups = []
        
        for f in os.listdir(self.backup_dir):
            if f.startswith(file_name) and f.endswith('.bak'):
                backup_file = os.path.join(self.backup_dir, f)
                backups.append((backup_file, os.path.getmtime(backup_file)))
        
        if backups:
            # 按修改时间排序，取最新的
            backups.sort(key=lambda x: x[1], reverse=True)
            latest_backup = backups[0][0]
            shutil.copy2(latest_backup, file_path)
            return True
        
        return False
    
    def log_fill_operation(self, entry: FillLogEntry):
        """记录填充操作"""
        log_file = os.path.join(self.log_dir, f"fill_{datetime.now().strftime('%Y%m')}.jsonl")
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(json.dumps(asdict(entry), ensure_ascii=False) + '\n')
    
    def get_fill_history(self, template_name: Optional[str] = None,
                        limit: int = 100) -> List[FillLogEntry]:
        """
        获取填充历史
        
        Args:
            template_name: 筛选特定模板
            limit: 最大返回数量
        """
        history = []
        
        # 读取所有日志文件
        for log_file in os.listdir(self.log_dir):
            if log_file.startswith('fill_') and log_file.endswith('.jsonl'):
                log_path = os.path.join(self.log_dir, log_file)
                with open(log_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.strip():
                            data = json.loads(line)
                            entry = FillLogEntry(**data)
                            if template_name is None or entry.template_name == template_name:
                                history.append(entry)
        
        # 按时间排序，取最新的
        history.sort(key=lambda x: x.timestamp, reverse=True)
        return history[:limit]
    
    def generate_report(self, case_type: str, output_file: Optional[str] = None) -> str:
        """
        生成填充报告
        
        Returns:
            报告文件路径
        """
        if output_file is None:
            output_file = os.path.join(self.log_dir, 
                                      f"report_{case_type}_{datetime.now().strftime('%Y%m%d')}.txt")
        
        outputs = self.list_outputs(case_type)
        history = self.get_fill_history()
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"=" * 60 + '\n')
            f.write(f"法律援助文档填充报告\n")
            f.write(f"案件类型: {case_type}\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"=" * 60 + '\n\n')
            
            # 输出文件统计
            f.write("【输出文件】\n")
            if case_type in outputs:
                for file_path in outputs[case_type]:
                    f.write(f"  - {os.path.basename(file_path)}\n")
            f.write(f"\n总计: {len(outputs.get(case_type, []))} 个文件\n\n")
            
            # 操作历史
            f.write("【最近操作】\n")
            for entry in history[:10]:
                status = "✓" if entry.success else "✗"
                f.write(f"  {status} {entry.timestamp} - {entry.output_name}\n")
                if entry.fields_filled:
                    f.write(f"    填充字段: {', '.join(entry.fields_filled[:5])}\n")
        
        return output_file


# ==================== 便捷函数 ====================

def get_manager(base_dir: Optional[str] = None) -> DocumentManager:
    """获取文档管理器实例"""
    if base_dir is None:
        base_dir = r"d:\trae\法律援助文书"
    return DocumentManager(base_dir)


# ==================== 测试 ====================

if __name__ == "__main__":
    import re
    
    print("=" * 60)
    print("文档管理模块测试")
    print("=" * 60)
    
    # 创建临时测试目录
    test_dir = os.path.join(os.path.expanduser("~"), "temp_legal_test")
    manager = DocumentManager(test_dir)
    
    print(f"\n基础目录: {test_dir}")
    print(f"输出目录: {manager.output_dir}")
    print(f"备份目录: {manager.backup_dir}")
    print(f"日志目录: {manager.log_dir}")
    
    # 测试优化5: 增量更新
    print("\n【优化5: 增量更新测试】")
    empty_values = ['', '   ', '{bianh}', '[[anyou]]', '【weitr】', '无', '_____', '正常值']
    for val in empty_values:
        is_empty = manager.is_field_empty(val)
        print(f"  '{val}' -> {'空' if is_empty else '非空'}")
    
    # 测试优化6: 输出组织
    print("\n【优化6: 输出组织测试】")
    test_template = "法援结案.docx"
    output_path = manager.get_output_path(test_template, "民事")
    print(f"  模板: {test_template}")
    print(f"  输出路径: {output_path}")
    
    # 测试优化7: 日志与回滚
    print("\n【优化7: 日志与回滚测试】")
    entry = FillLogEntry(
        timestamp=datetime.now().isoformat(),
        template_name="法援结案.docx",
        output_name="法援结案_已填充.docx",
        elements_used={"bianh": "2024-001", "anyou": "劳务合同纠纷"},
        fields_filled=["bianh", "anyou", "weitr"],
        fields_skipped=[],
        success=True
    )
    manager.log_fill_operation(entry)
    print(f"  日志已记录")
    
    # 生成报告
    report_path = manager.generate_report("民事")
    print(f"  报告已生成: {report_path}")
    
    # 清理
    import shutil
    shutil.rmtree(test_dir)
    print(f"\n✓ 测试完成，已清理临时文件")
