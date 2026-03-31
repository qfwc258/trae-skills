#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
案件类型管理模块 - 优化2
支持民事、刑事、行政、国家赔偿等多种案件类型
自动检测案件类型，提供类型特定的字段映射
"""

import os
import re
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field
from enum import Enum

class CaseCategory(Enum):
    """案件类别枚举"""
    CIVIL = "民事"
    CRIMINAL = "刑事"
    ADMINISTRATIVE = "行政"
    STATE_COMPENSATION = "国家赔偿"
    LABOR = "劳动仲裁"
    OTHER = "其他"

@dataclass
class CaseTypeConfig:
    """案件类型配置"""
    category: CaseCategory
    element_filename: str
    display_name: str
    description: str
    # 该类型特有的字段
    specific_fields: Dict[str, str] = field(default_factory=dict)
    # 必填字段
    required_fields: List[str] = field(default_factory=list)
    # 可选字段
    optional_fields: List[str] = field(default_factory=list)


class CaseTypeManager:
    """案件类型管理器"""
    
    # 默认元素文件名模式
    ELEMENT_FILE_PATTERNS = {
        CaseCategory.CIVIL: ["元素_民事.txt", "民事_元素.txt", "civil_elements.txt"],
        CaseCategory.CRIMINAL: ["元素_刑事.txt", "刑事_元素.txt", "criminal_elements.txt"],
        CaseCategory.ADMINISTRATIVE: ["元素_行政.txt", "行政_元素.txt", "admin_elements.txt"],
        CaseCategory.STATE_COMPENSATION: ["元素_国赔.txt", "国家赔偿_元素.txt", "compensation_elements.txt"],
        CaseCategory.LABOR: ["元素_劳动.txt", "劳动_元素.txt", "labor_elements.txt"],
    }
    
    # 案件类型配置
    CASE_TYPE_CONFIGS = {
        CaseCategory.CIVIL: CaseTypeConfig(
            category=CaseCategory.CIVIL,
            element_filename="元素_民事.txt",
            display_name="民事案件",
            description="民事法律援助案件，包括合同纠纷、侵权责任、婚姻家庭等",
            specific_fields={
                'dfdsr': '对方当事人',
                'badw': '办案机关/法院',
                'jied': '所处阶段',
                'ssqq': '诉讼请求',
                'ajqk': '案件情况',
            },
            required_fields=['leib', 'anyou', 'weitr', 'badw', 'jied'],
            optional_fields=['dfdsr', 'wtrsfz', 'wtrdh', 'ssqq']
        ),
        
        CaseCategory.CRIMINAL: CaseTypeConfig(
            category=CaseCategory.CRIMINAL,
            element_filename="元素_刑事.txt",
            display_name="刑事案件",
            description="刑事法律援助案件，包括侦查、审查起诉、审判阶段",
            specific_fields={
                'xyr': '犯罪嫌疑人/被告人',
                'zmr': '罪名',
                'badw': '办案机关',
                'jied': '所处阶段',  # 侦查/审查起诉/一审/二审
                'zcr': '侦查机关',
                'jcjg': '检察机关',
                'fy': '法院',
                'shyj': '辩护/代理意见',
                'bhsr': '被害人',
            },
            required_fields=['leib', 'anyou', 'weitr', 'zmr', 'badw', 'jied'],
            optional_fields=['xyr', 'zcr', 'jcjg', 'fy', 'bhsr']
        ),
        
        CaseCategory.ADMINISTRATIVE: CaseTypeConfig(
            category=CaseCategory.ADMINISTRATIVE,
            element_filename="元素_行政.txt",
            display_name="行政案件",
            description="行政法律援助案件，包括行政复议、行政诉讼等",
            specific_fields={
                'xzzw': '行政职务',
                'xzjg': '行政机关',
                'xzfs': '行政方式',  # 复议/诉讼
                'sxqq': '诉求请求',
                'xzjd': '行政决定',
            },
            required_fields=['leib', 'anyou', 'weitr', 'xzjg', 'xzfs'],
            optional_fields=['xzzw', 'sxqq', 'xzjd']
        ),
        
        CaseCategory.STATE_COMPENSATION: CaseTypeConfig(
            category=CaseCategory.STATE_COMPENSATION,
            element_filename="元素_国赔.txt",
            display_name="国家赔偿案件",
            description="国家赔偿法律援助案件",
            specific_fields={
                'pcjg': '赔偿机关',
                'pclx': '赔偿类型',  # 行政赔偿/刑事赔偿
                'pcje': '赔偿金额',
                'pcls': '赔偿理由',
            },
            required_fields=['leib', 'anyou', 'weitr', 'pcjg', 'pclx'],
            optional_fields=['pcje', 'pcls']
        ),
        
        CaseCategory.LABOR: CaseTypeConfig(
            category=CaseCategory.LABOR,
            element_filename="元素_劳动.txt",
            display_name="劳动仲裁案件",
            description="劳动仲裁法律援助案件",
            specific_fields={
                'yjdw': '用人单位',
                'ldgx': '劳动关系',
                'zyqq': '仲裁请求',
                'ldzy': '劳动仲裁委',
            },
            required_fields=['leib', 'anyou', 'weitr', 'yjdw'],
            optional_fields=['ldgx', 'zyqq', 'ldzy']
        ),
    }
    
    def __init__(self):
        self.configs = self.CASE_TYPE_CONFIGS
    
    def detect_case_type(self, input_dir: str, element_file: Optional[str] = None) -> Tuple[CaseCategory, str]:
        """
        自动检测案件类型
        
        Returns:
            (案件类别, 元素文件路径)
        """
        # 如果指定了元素文件，尝试从文件名推断类型
        if element_file:
            category = self._infer_from_filename(element_file)
            if category:
                return category, element_file
        
        # 在目录中搜索已知的元素文件
        for category, filenames in self.ELEMENT_FILE_PATTERNS.items():
            for filename in filenames:
                filepath = os.path.join(input_dir, filename)
                if os.path.exists(filepath):
                    return category, filepath
        
        # 默认返回民事案件
        default_file = os.path.join(input_dir, "元素_民事.txt")
        return CaseCategory.CIVIL, default_file
    
    def _infer_from_filename(self, filename: str) -> Optional[CaseCategory]:
        """从文件名推断案件类型"""
        basename = os.path.basename(filename).lower()
        
        if any(kw in basename for kw in ['刑事', 'criminal', 'xing']):
            return CaseCategory.CRIMINAL
        elif any(kw in basename for kw in ['行政', 'admin', 'xz']):
            return CaseCategory.ADMINISTRATIVE
        elif any(kw in basename for kw in ['国赔', '赔偿', 'compensation', 'pcb']):
            return CaseCategory.STATE_COMPENSATION
        elif any(kw in basename for kw in ['劳动', 'labor', 'ld', 'arbitration']):
            return CaseCategory.LABOR
        elif any(kw in basename for kw in ['民事', 'civil', 'ms']):
            return CaseCategory.CIVIL
        
        return None
    
    def get_config(self, category: CaseCategory) -> CaseTypeConfig:
        """获取案件类型配置"""
        return self.configs.get(category, self.configs[CaseCategory.CIVIL])
    
    def get_all_element_filenames(self) -> List[str]:
        """获取所有可能的元素文件名"""
        filenames = []
        for category_filenames in self.ELEMENT_FILE_PATTERNS.values():
            filenames.extend(category_filenames)
        return filenames
    
    def validate_elements(self, elements: Dict[str, str], category: CaseCategory) -> Tuple[bool, List[str]]:
        """
        验证元素是否完整
        
        Returns:
            (是否完整, 缺失的必填字段列表)
        """
        config = self.get_config(category)
        missing = []
        
        for field in config.required_fields:
            if field not in elements or not elements[field].strip():
                missing.append(field)
        
        return len(missing) == 0, missing
    
    def build_mapping_for_category(self, elements: Dict[str, str], category: CaseCategory) -> Dict[str, str]:
        """
        根据案件类型构建替换映射
        
        包含通用字段 + 类型特定字段
        """
        mapping = {}
        config = self.get_config(category)
        
        # 通用字段（所有类型共有）
        common_fields = ['bianh', 'leib', 'anyou', 'weitr', 'zprq', 'jarq', 
                        'wtrsfz', 'wtrdh', 'wtrxb', 'wtrcs', 'wtrzz',
                        'cbxj', 'ljsm', 'gdrq', 'yjrq']
        
        for field in common_fields:
            if field in elements:
                mapping[field] = elements[field]
        
        # 类型特定字段
        for field in config.specific_fields.keys():
            if field in elements:
                mapping[field] = elements[field]
        
        # 办案过程字段（通用）
        for i in range(1, 10):
            for prefix in ['gcsj', 'gcfs', 'gcnr', 'thsj', 'thdd']:
                key = f"{prefix}{i}"
                if key in elements:
                    mapping[key] = elements[key]
        
        return mapping
    
    def get_case_type_description(self, category: CaseCategory) -> str:
        """获取案件类型描述"""
        config = self.get_config(category)
        return f"{config.display_name} - {config.description}"
    
    def list_available_types(self) -> List[Tuple[CaseCategory, str]]:
        """列出所有可用的案件类型"""
        return [(cat, config.display_name) for cat, config in self.configs.items()]


# ==================== 便捷函数 ====================

def detect_case_type(input_dir: str, element_file: Optional[str] = None) -> Tuple[CaseCategory, str]:
    """便捷函数：检测案件类型"""
    manager = CaseTypeManager()
    return manager.detect_case_type(input_dir, element_file)

def get_case_config(category: CaseCategory) -> CaseTypeConfig:
    """便捷函数：获取案件配置"""
    manager = CaseTypeManager()
    return manager.get_config(category)


# ==================== 测试 ====================

if __name__ == "__main__":
    manager = CaseTypeManager()
    
    print("=" * 60)
    print("案件类型管理器测试")
    print("=" * 60)
    
    print("\n可用案件类型：")
    for category, name in manager.list_available_types():
        config = manager.get_config(category)
        print(f"  {category.value}: {name}")
        print(f"    元素文件: {config.element_filename}")
        print(f"    描述: {config.description}")
        print(f"    特定字段: {list(config.specific_fields.keys())}")
        print()
    
    # 测试文件名推断
    test_files = [
        "元素_民事.txt",
        "元素_刑事.txt",
        "元素_行政.txt",
        "元素_国赔.txt",
        "元素_劳动.txt",
    ]
    
    print("文件名推断测试：")
    for filename in test_files:
        category = manager._infer_from_filename(filename)
        print(f"  {filename} -> {category.value if category else '未知'}")
