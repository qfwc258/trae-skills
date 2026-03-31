#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交互式填充模块 - 优化4
缺失元素时提示用户输入
"""

import sys
from typing import Dict, List, Optional, Any, Callable
from dataclasses import dataclass
from enum import Enum


class InputType(Enum):
    """输入类型"""
    TEXT = "text"           # 普通文本
    DATE = "date"           # 日期
    NUMBER = "number"       # 数字
    CHOICE = "choice"       # 选择
    BOOLEAN = "boolean"     # 是/否
    MULTILINE = "multiline" # 多行文本


@dataclass
class FieldPrompt:
    """字段提示信息"""
    key: str
    display_name: str
    description: str = ""
    input_type: InputType = InputType.TEXT
    required: bool = True
    default_value: Optional[str] = None
    choices: Optional[List[str]] = None  # 用于CHOICE类型
    validation: Optional[Callable[[str], bool]] = None


class InteractiveFiller:
    """交互式填充器"""
    
    def __init__(self, auto_fill: bool = False):
        """
        Args:
            auto_fill: 是否自动填充（非交互模式）
        """
        self.auto_fill = auto_fill
        self.cache: Dict[str, str] = {}  # 缓存用户输入
    
    def prompt_for_field(self, prompt: FieldPrompt) -> Optional[str]:
        """
        提示用户输入单个字段
        
        Returns:
            用户输入的值，如果跳过则返回None
        """
        if self.auto_fill:
            return prompt.default_value
        
        # 检查缓存
        if prompt.key in self.cache:
            return self.cache[prompt.key]
        
        # 构建提示文本
        header = f"\n{'='*50}"
        title = f"\n【{prompt.display_name}】"
        desc = f"\n{prompt.description}" if prompt.description else ""
        req = " (必填)" if prompt.required else " (可选)"
        default = f"\n默认值: {prompt.default_value}" if prompt.default_value else ""
        
        print(header)
        print(title + req)
        print(desc)
        print(default)
        print(f"{'='*50}")
        
        # 根据输入类型获取值
        value = self._get_input_by_type(prompt)
        
        # 验证
        if value and prompt.validation and not prompt.validation(value):
            print("⚠️ 输入格式不正确，请重新输入")
            return self.prompt_for_field(prompt)
        
        # 缓存
        if value:
            self.cache[prompt.key] = value
        
        return value
    
    def _get_input_by_type(self, prompt: FieldPrompt) -> Optional[str]:
        """根据输入类型获取用户输入"""
        
        if prompt.input_type == InputType.CHOICE and prompt.choices:
            return self._prompt_choice(prompt.display_name, prompt.choices, prompt.default_value)
        
        elif prompt.input_type == InputType.BOOLEAN:
            return self._prompt_boolean(prompt.display_name, prompt.default_value)
        
        elif prompt.input_type == InputType.MULTILINE:
            return self._prompt_multiline(prompt.display_name, prompt.default_value)
        
        else:
            # 普通文本输入
            prompt_text = f"请输入{prompt.display_name}"
            if prompt.default_value:
                prompt_text += f" (直接回车使用默认值: {prompt.default_value})"
            prompt_text += ": "
            
            try:
                value = input(prompt_text).strip()
                if not value and prompt.default_value:
                    value = prompt.default_value
                return value if value else None
            except (EOFError, KeyboardInterrupt):
                print("\n用户取消输入")
                return None
    
    def _prompt_choice(self, name: str, choices: List[str], default: Optional[str] = None) -> Optional[str]:
        """提示选择"""
        print(f"\n请选择{name}:")
        for i, choice in enumerate(choices, 1):
            marker = " (默认)" if choice == default else ""
            print(f"  {i}. {choice}{marker}")
        
        try:
            value = input(f"请输入选项编号 (1-{len(choices)}): ").strip()
            if not value and default:
                return default
            
            idx = int(value) - 1
            if 0 <= idx < len(choices):
                return choices[idx]
            else:
                print("⚠️ 无效的选项")
                return self._prompt_choice(name, choices, default)
        except (ValueError, EOFError, KeyboardInterrupt):
            if default:
                return default
            return None
    
    def _prompt_boolean(self, name: str, default: Optional[str] = None) -> Optional[str]:
        """提示是/否"""
        default_hint = ""
        if default:
            default_hint = "Y" if default.lower() in ['yes', 'true', 'y', '是'] else "N"
        
        prompt_text = f"{name}? (Y/N"
        if default_hint:
            prompt_text += f", 默认{default_hint}"
        prompt_text += "): "
        
        try:
            value = input(prompt_text).strip().lower()
            if not value and default:
                return "是" if default.lower() in ['yes', 'true', 'y', '是'] else "否"
            
            if value in ['y', 'yes', '是', 'true', '1']:
                return "是"
            elif value in ['n', 'no', '否', 'false', '0']:
                return "否"
            else:
                print("⚠️ 请输入 Y 或 N")
                return self._prompt_boolean(name, default)
        except (EOFError, KeyboardInterrupt):
            return None
    
    def _prompt_multiline(self, name: str, default: Optional[str] = None) -> Optional[str]:
        """提示多行输入"""
        print(f"\n请输入{name} (多行，输入空行结束):")
        if default:
            print(f"默认值:\n{default}")
            print("直接输入新内容覆盖，或输入空行使用默认值")
        
        lines = []
        try:
            while True:
                line = input()
                if not line:
                    break
                lines.append(line)
            
            if not lines and default:
                return default
            
            return '\n'.join(lines) if lines else None
        except (EOFError, KeyboardInterrupt):
            return None
    
    def fill_missing_elements(self, elements: Dict[str, str], 
                             missing_fields: List[str],
                             field_prompts: Optional[Dict[str, FieldPrompt]] = None) -> Dict[str, str]:
        """
        填充缺失的元素
        
        Args:
            elements: 现有元素
            missing_fields: 缺失的字段列表
            field_prompts: 字段提示配置
        
        Returns:
            更新后的元素字典
        """
        if not missing_fields:
            return elements
        
        result = elements.copy()
        
        print(f"\n{'='*60}")
        print(f"发现 {len(missing_fields)} 个缺失字段，请补充输入")
        print(f"{'='*60}")
        
        for field in missing_fields:
            # 获取字段提示
            prompt = self._get_default_prompt(field)
            if field_prompts and field in field_prompts:
                prompt = field_prompts[field]
            
            value = self.prompt_for_field(prompt)
            if value:
                result[field] = value
        
        return result
    
    def _get_default_prompt(self, field: str) -> FieldPrompt:
        """获取默认字段提示"""
        # 常见字段的默认配置
        common_prompts = {
            'bianh': FieldPrompt('bianh', '案件编号', '案件的编号，如：(2024)湘援0124民1号'),
            'anyou': FieldPrompt('anyou', '案由', '案件的案由，如：劳务合同纠纷'),
            'weitr': FieldPrompt('weitr', '受援人', '受援人姓名'),
            'dfdsr': FieldPrompt('dfdsr', '对方当事人', '对方当事人姓名或单位'),
            'badw': FieldPrompt('badw', '办案机关', '办案机关或法院名称'),
            'jied': FieldPrompt('jied', '所处阶段', '案件所处阶段', 
                               InputType.CHOICE, 
                               choices=['侦查', '审查起诉', '一审', '二审', '再审', '执行']),
            'zprq': FieldPrompt('zprq', '指派日期', '法律援助指派日期', InputType.DATE),
            'jarq': FieldPrompt('jarq', '结案日期', '案件结案日期', InputType.DATE),
            'wtrsfz': FieldPrompt('wtrsfz', '身份证号', '受援人身份证号'),
            'wtrdh': FieldPrompt('wtrdh', '联系电话', '受援人联系电话'),
            'wtrxb': FieldPrompt('wtrxb', '性别', '受援人性别', 
                                InputType.CHOICE, 
                                choices=['男', '女']),
            'cbxj': FieldPrompt('cbxj', '承办小结', '案件承办情况小结', InputType.MULTILINE),
            'ljsm': FieldPrompt('ljsm', '立卷说明', '立卷说明', InputType.MULTILINE),
        }
        
        if field in common_prompts:
            return common_prompts[field]
        
        # 默认配置
        return FieldPrompt(field, field, f"请输入{field}")
    
    def confirm_batch_operation(self, operation: str, count: int) -> bool:
        """确认批量操作"""
        if self.auto_fill:
            return True
        
        print(f"\n即将{operation}，共 {count} 项")
        try:
            value = input("确认执行? (Y/N): ").strip().lower()
            return value in ['y', 'yes', '是', 'true']
        except (EOFError, KeyboardInterrupt):
            return False
    
    def show_summary(self, elements: Dict[str, str], title: str = "元素摘要"):
        """显示元素摘要"""
        print(f"\n{'='*60}")
        print(f"{title}")
        print(f"{'='*60}")
        
        # 分组显示
        groups = {
            '基本信息': ['bianh', 'leib', 'anyou', 'weitr', 'dfdsr'],
            '办案信息': ['badw', 'jied', 'zprq', 'jarq'],
            '当事人信息': ['wtrsfz', 'wtrdh', 'wtrxb', 'wtrcs', 'wtrzz'],
        }
        
        for group_name, fields in groups.items():
            print(f"\n【{group_name}】")
            for field in fields:
                if field in elements:
                    value = elements[field]
                    if len(value) > 30:
                        value = value[:30] + "..."
                    print(f"  {field}: {value}")
        
        # 其他字段
        other_fields = [k for k in elements.keys() if not any(k in g for g in groups.values())]
        if other_fields:
            print(f"\n【其他】")
            for field in other_fields[:10]:  # 最多显示10个
                value = elements[field]
                if len(value) > 30:
                    value = value[:30] + "..."
                print(f"  {field}: {value}")
            if len(other_fields) > 10:
                print(f"  ... 还有 {len(other_fields) - 10} 个字段")


# ==================== 便捷函数 ====================

def prompt_for_missing(elements: Dict[str, str], 
                       missing_fields: List[str],
                       auto_fill: bool = False) -> Dict[str, str]:
    """便捷函数：提示填充缺失字段"""
    filler = InteractiveFiller(auto_fill)
    return filler.fill_missing_elements(elements, missing_fields)


def confirm_operation(message: str) -> bool:
    """便捷函数：确认操作"""
    filler = InteractiveFiller()
    return filler.confirm_batch_operation(message, 1)


# ==================== 测试 ====================

if __name__ == "__main__":
    print("=" * 60)
    print("交互式填充模块测试")
    print("=" * 60)
    
    # 测试数据
    elements = {
        'bianh': '2024-001',
        'anyou': '劳务合同纠纷',
        'weitr': '张三',
    }
    
    missing = ['badw', 'jied', 'zprq', 'wtrdh']
    
    print("\n现有元素：")
    for k, v in elements.items():
        print(f"  {k}: {v}")
    
    print(f"\n缺失字段: {missing}")
    
    # 创建填充器（非交互模式用于测试）
    filler = InteractiveFiller(auto_fill=True)
    
    # 添加默认值
    filler.cache = {
        'badw': '宁乡市人民法院',
        'jied': '一审',
        'zprq': '2024-01-15',
        'wtrdh': '13800138000'
    }
    
    result = filler.fill_missing_elements(elements, missing)
    
    print("\n填充后的元素：")
    for k, v in result.items():
        print(f"  {k}: {v}")
    
    print("\n✓ 测试完成")
