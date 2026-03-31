#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板配置管理模块 - 优化3
支持JSON配置文件定义填充规则
"""

import json
import os
import re
from typing import Dict, List, Optional, Any, Callable
from dataclasses import dataclass, field, asdict
from enum import Enum


class TransformType(Enum):
    """字段转换类型"""
    UPPER = "upper"           # 转大写
    LOWER = "lower"           # 转小写
    TITLE = "title"           # 首字母大写
    DATE_FORMAT = "date"      # 日期格式化
    NUMBER_FORMAT = "number"  # 数字格式化
    SUBSTRING = "substr"      # 子字符串
    REPLACE = "replace"       # 替换
    CUSTOM = "custom"         # 自定义函数


@dataclass
class FieldTransform:
    """字段转换规则"""
    transform_type: TransformType
    params: Dict[str, Any] = field(default_factory=dict)
    
    def apply(self, value: str) -> str:
        """应用转换"""
        if not value:
            return value
        
        if self.transform_type == TransformType.UPPER:
            return value.upper()
        elif self.transform_type == TransformType.LOWER:
            return value.lower()
        elif self.transform_type == TransformType.TITLE:
            return value.title()
        elif self.transform_type == TransformType.DATE_FORMAT:
            # 日期格式化
            from datetime import datetime
            fmt = self.params.get('format', '%Y-%m-%d')
            try:
                dt = datetime.strptime(value, '%Y-%m-%d')
                return dt.strftime(fmt)
            except:
                return value
        elif self.transform_type == TransformType.NUMBER_FORMAT:
            # 数字格式化
            try:
                num = float(value)
                fmt = self.params.get('format', ',.2f')
                return format(num, fmt)
            except:
                return value
        elif self.transform_type == TransformType.SUBSTRING:
            # 子字符串
            start = self.params.get('start', 0)
            end = self.params.get('end', None)
            return value[start:end]
        elif self.transform_type == TransformType.REPLACE:
            # 替换
            old = self.params.get('old', '')
            new = self.params.get('new', '')
            return value.replace(old, new)
        else:
            return value


@dataclass
class FieldMapping:
    """字段映射规则"""
    source_key: str                    # 源字段名（来自元素文件）
    target_key: str                    # 目标字段名（模板中的占位符）
    default_value: Optional[str] = None  # 默认值
    transforms: List[FieldTransform] = field(default_factory=list)  # 转换规则
    condition: Optional[str] = None    # 条件表达式
    
    def get_value(self, elements: Dict[str, str]) -> Optional[str]:
        """获取转换后的值"""
        # 检查条件
        if self.condition and not self._evaluate_condition(elements):
            return None
        
        # 获取原始值
        value = elements.get(self.source_key, self.default_value)
        if value is None:
            return None
        
        # 应用转换
        for transform in self.transforms:
            value = transform.apply(value)
        
        return value
    
    def _evaluate_condition(self, elements: Dict[str, str]) -> bool:
        """评估条件表达式"""
        if not self.condition:
            return True
        
        # 简单的条件表达式支持
        # 格式: "key==value" 或 "key!=value" 或 "key exists"
        try:
            if '==' in self.condition:
                key, val = self.condition.split('==', 1)
                return elements.get(key.strip()) == val.strip()
            elif '!=' in self.condition:
                key, val = self.condition.split('!=', 1)
                return elements.get(key.strip()) != val.strip()
            elif 'exists' in self.condition:
                key = self.condition.replace('exists', '').strip()
                return key in elements and elements[key]
        except:
            pass
        
        return True


@dataclass
class TemplateRule:
    """单个模板的填充规则"""
    template_name: str                 # 模板文件名（支持通配符）
    description: Optional[str] = None  # 描述
    field_mappings: List[FieldMapping] = field(default_factory=list)  # 字段映射
    include_common: bool = True        # 是否包含通用字段
    output_suffix: str = "_已填充"      # 输出文件后缀
    
    def applies_to(self, template_path: str) -> bool:
        """检查规则是否适用于指定模板"""
        template_filename = os.path.basename(template_path)
        # 支持通配符匹配
        pattern = self.template_name.replace('*', '.*').replace('?', '.')
        return re.match(pattern, template_filename) is not None


@dataclass
class TemplateConfig:
    """模板配置"""
    version: str = "1.0"
    description: str = ""
    case_type: str = "民事"             # 案件类型
    common_mappings: List[FieldMapping] = field(default_factory=list)  # 通用映射
    template_rules: List[TemplateRule] = field(default_factory=list)   # 模板特定规则
    global_settings: Dict[str, Any] = field(default_factory=dict)      # 全局设置
    
    @classmethod
    def from_json(cls, json_path: str) -> 'TemplateConfig':
        """从JSON文件加载配置"""
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        return cls._from_dict(data)
    
    @classmethod
    def _from_dict(cls, data: Dict) -> 'TemplateConfig':
        """从字典创建配置"""
        config = cls(
            version=data.get('version', '1.0'),
            description=data.get('description', ''),
            case_type=data.get('case_type', '民事'),
            global_settings=data.get('global_settings', {})
        )
        
        # 解析通用映射
        for mapping_data in data.get('common_mappings', []):
            config.common_mappings.append(cls._parse_field_mapping(mapping_data))
        
        # 解析模板规则
        for rule_data in data.get('template_rules', []):
            rule = TemplateRule(
                template_name=rule_data.get('template_name', ''),
                description=rule_data.get('description'),
                include_common=rule_data.get('include_common', True),
                output_suffix=rule_data.get('output_suffix', '_已填充')
            )
            
            for mapping_data in rule_data.get('field_mappings', []):
                rule.field_mappings.append(cls._parse_field_mapping(mapping_data))
            
            config.template_rules.append(rule)
        
        return config
    
    @classmethod
    def _parse_field_mapping(cls, data: Dict) -> FieldMapping:
        """解析字段映射"""
        transforms = []
        for transform_data in data.get('transforms', []):
            transform_type = TransformType(transform_data.get('type', 'custom'))
            params = transform_data.get('params', {})
            transforms.append(FieldTransform(transform_type, params))
        
        return FieldMapping(
            source_key=data.get('source_key', ''),
            target_key=data.get('target_key', ''),
            default_value=data.get('default_value'),
            transforms=transforms,
            condition=data.get('condition')
        )
    
    def to_json(self, json_path: str):
        """保存配置到JSON文件"""
        data = asdict(self)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def get_rule_for_template(self, template_path: str) -> Optional[TemplateRule]:
        """获取适用于指定模板的规则"""
        for rule in self.template_rules:
            if rule.applies_to(template_path):
                return rule
        return None
    
    def build_mapping(self, elements: Dict[str, str], template_path: str) -> Dict[str, str]:
        """
        根据配置构建替换映射
        
        Args:
            elements: 元素文件解析结果
            template_path: 模板文件路径
        
        Returns:
            占位符到值的映射
        """
        mapping = {}
        
        # 应用通用映射
        for field_mapping in self.common_mappings:
            value = field_mapping.get_value(elements)
            if value is not None:
                mapping[field_mapping.target_key] = value
        
        # 应用模板特定规则
        rule = self.get_rule_for_template(template_path)
        if rule:
            if rule.include_common:
                # 已经包含了通用映射
                pass
            else:
                # 不包含通用映射，清空
                mapping = {}
            
            for field_mapping in rule.field_mappings:
                value = field_mapping.get_value(elements)
                if value is not None:
                    mapping[field_mapping.target_key] = value
        
        return mapping


class TemplateConfigManager:
    """模板配置管理器"""
    
    DEFAULT_CONFIG_FILENAME = "template_config.json"
    
    def __init__(self, config_dir: str = None):
        self.config_dir = config_dir or os.getcwd()
        self.configs: Dict[str, TemplateConfig] = {}
    
    def load_config(self, config_path: str) -> TemplateConfig:
        """加载配置文件"""
        config = TemplateConfig.from_json(config_path)
        self.configs[config_path] = config
        return config
    
    def get_or_create_config(self, input_dir: str) -> TemplateConfig:
        """获取或创建配置"""
        config_path = os.path.join(input_dir, self.DEFAULT_CONFIG_FILENAME)
        
        if os.path.exists(config_path):
            return self.load_config(config_path)
        
        # 创建默认配置
        config = self._create_default_config()
        return config
    
    def _create_default_config(self) -> TemplateConfig:
        """创建默认配置"""
        return TemplateConfig(
            version="1.0",
            description="默认模板配置",
            case_type="民事",
            common_mappings=[
                FieldMapping('bianh', 'bianh'),
                FieldMapping('leib', 'leib'),
                FieldMapping('anyou', 'anyou'),
                FieldMapping('weitr', 'weitr'),
                FieldMapping('zprq', 'zprq'),
                FieldMapping('jarq', 'jarq'),
            ],
            global_settings={
                'keep_unmatched': True,
                'smart_fill': False,
                'use_enhanced': False,
            }
        )
    
    def save_config(self, config: TemplateConfig, config_path: str):
        """保存配置"""
        config.to_json(config_path)
        self.configs[config_path] = config


# ==================== 便捷函数 ====================

def load_template_config(config_path: str) -> TemplateConfig:
    """便捷函数：加载模板配置"""
    return TemplateConfig.from_json(config_path)

def create_default_config(input_dir: str) -> str:
    """便捷函数：创建默认配置文件"""
    manager = TemplateConfigManager(input_dir)
    config = manager._create_default_config()
    config_path = os.path.join(input_dir, TemplateConfigManager.DEFAULT_CONFIG_FILENAME)
    manager.save_config(config, config_path)
    return config_path


# ==================== 示例配置生成 ====================

def generate_example_config() -> Dict:
    """生成示例配置"""
    return {
        "version": "1.0",
        "description": "法律援助文档填充配置示例",
        "case_type": "民事",
        "global_settings": {
            "keep_unmatched": True,
            "smart_fill": False,
            "use_enhanced": True
        },
        "common_mappings": [
            {
                "source_key": "bianh",
                "target_key": "bianh",
                "default_value": "未编号"
            },
            {
                "source_key": "anyou",
                "target_key": "anyou",
                "transforms": [
                    {"type": "title", "params": {}}
                ]
            },
            {
                "source_key": "zprq",
                "target_key": "zprq",
                "transforms": [
                    {"type": "date", "params": {"format": "%Y年%m月%d日"}}
                ]
            }
        ],
        "template_rules": [
            {
                "template_name": "法援结案.docx",
                "description": "法援结案报告模板",
                "include_common": True,
                "output_suffix": "_已填充",
                "field_mappings": [
                    {
                        "source_key": "cbxj",
                        "target_key": "cbxj",
                        "default_value": "（暂无承办小结）"
                    }
                ]
            },
            {
                "template_name": "阅卷笔录.docx",
                "description": "阅卷笔录模板",
                "include_common": True,
                "field_mappings": [
                    {
                        "source_key": "yjrq",
                        "target_key": "yjrq",
                        "transforms": [
                            {"type": "date", "params": {"format": "%Y年%m月%d日"}}
                        ]
                    }
                ]
            }
        ]
    }


# ==================== 测试 ====================

if __name__ == "__main__":
    print("=" * 60)
    print("模板配置管理器测试")
    print("=" * 60)
    
    # 生成示例配置
    example = generate_example_config()
    print("\n示例配置：")
    print(json.dumps(example, ensure_ascii=False, indent=2))
    
    # 创建临时配置文件
    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        json.dump(example, f, ensure_ascii=False, indent=2)
        temp_path = f.name
    
    print(f"\n临时配置文件: {temp_path}")
    
    # 加载配置
    config = load_template_config(temp_path)
    print(f"\n加载的配置:")
    print(f"  版本: {config.version}")
    print(f"  描述: {config.description}")
    print(f"  案件类型: {config.case_type}")
    print(f"  通用映射: {len(config.common_mappings)} 个")
    print(f"  模板规则: {len(config.template_rules)} 个")
    
    # 测试构建映射
    test_elements = {
        'bianh': '2024-001',
        'anyou': '劳务合同纠纷',
        'weitr': '张三',
        'zprq': '2024-01-15',
        'cbxj': '本案已顺利结案',
        'yjrq': '2024-02-01'
    }
    
    print(f"\n测试构建映射:")
    mapping = config.build_mapping(test_elements, "法援结案.docx")
    for key, value in mapping.items():
        print(f"  {key} -> {value}")
    
    # 清理临时文件
    os.unlink(temp_path)
    print("\n✓ 测试完成")
