#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能占位符检测模块 - 增强版
支持多种占位符格式、模糊匹配、多行内容、条件占位符
"""

import re
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum

class PlaceholderType(Enum):
    """占位符类型枚举"""
    BRACES = "braces"           # {key}
    DOUBLE_BRACES = "double"    # {{key}}
    SQUARE = "square"           # [key]
    DOUBLE_SQUARE = "dsquare"   # [[key]]
    CHINESE = "chinese"         # 【key】
    ANGLE = "angle"             # <key>
    DOUBLE_ANGLE = "dangle"     # <<key>>
    HASH = "hash"               # #key#
    DOLLAR = "dollar"           # $key$
    PERCENT = "percent"         # %key%
    PLAIN = "plain"             # key (纯文本)

@dataclass
class PlaceholderMatch:
    """占位符匹配结果"""
    original: str           # 原始占位符文本
    key: str               # 提取的键名
    placeholder_type: PlaceholderType
    start_pos: int         # 起始位置
    end_pos: int           # 结束位置
    default_value: Optional[str] = None  # 默认值（用于条件占位符）
    format_spec: Optional[str] = None    # 格式说明

class PlaceholderDetector:
    """智能占位符检测器"""
    
    # 各种占位符格式的正则表达式模式
    PATTERNS = {
        PlaceholderType.DOUBLE_BRACES: r'\{\{\s*([^}]+?)\s*\}\}',
        PlaceholderType.BRACES: r'\{\s*([^}]+?)\s*\}',
        PlaceholderType.DOUBLE_SQUARE: r'\[\[\s*([^\]]+?)\s*\]\]',
        PlaceholderType.SQUARE: r'\[\s*([^\]]+?)\s*\]',
        PlaceholderType.CHINESE: r'【\s*([^】]+?)\s*】',
        PlaceholderType.DOUBLE_ANGLE: r'<<\s*([^>]+?)\s*>>',
        PlaceholderType.ANGLE: r'<\s*([^>]+?)\s*>',
        PlaceholderType.HASH: r'#\s*([^#]+?)\s*#',
        PlaceholderType.DOLLAR: r'\$\s*([^$]+?)\s*\$',
        PlaceholderType.PERCENT: r'%\s*([^%]+?)\s*%',
    }
    
    def __init__(self):
        self.compiled_patterns = {
            ptype: re.compile(pattern, re.MULTILINE | re.DOTALL)
            for ptype, pattern in self.PATTERNS.items()
        }
    
    def detect_all(self, text: str, include_plain: bool = False) -> List[PlaceholderMatch]:
        """
        检测文本中的所有占位符
        
        Args:
            text: 要检测的文本
            include_plain: 是否包含纯文本占位符（需要额外配置关键词列表）
        
        Returns:
            占位符匹配结果列表（按位置排序）
        """
        all_matches = []
        
        for ptype, pattern in self.compiled_patterns.items():
            for match in pattern.finditer(text):
                content = match.group(1).strip()
                key, default, format_spec = self._parse_content(content)
                
                all_matches.append(PlaceholderMatch(
                    original=match.group(0),
                    key=key,
                    placeholder_type=ptype,
                    start_pos=match.start(),
                    end_pos=match.end(),
                    default_value=default,
                    format_spec=format_spec
                ))
        
        # 按位置排序，避免重叠
        all_matches.sort(key=lambda x: x.start_pos)
        
        # 移除重叠的匹配（优先保留更具体的模式，如 {{key}} 优先于 {key}）
        filtered_matches = self._remove_overlaps(all_matches)
        
        return filtered_matches
    
    def _parse_content(self, content: str) -> Tuple[str, Optional[str], Optional[str]]:
        """
        解析占位符内容
        支持格式：
        - key
        - key?default_value (带默认值)
        - key:format_spec (带格式说明)
        - key?default_value:format_spec (组合)
        """
        key = content
        default = None
        format_spec = None
        
        # 检查格式说明
        if ':' in content:
            parts = content.rsplit(':', 1)
            key = parts[0].strip()
            format_spec = parts[1].strip()
        
        # 检查默认值
        if '?' in key:
            parts = key.split('?', 1)
            key = parts[0].strip()
            default = parts[1].strip()
        
        return key, default, format_spec
    
    def _remove_overlaps(self, matches: List[PlaceholderMatch]) -> List[PlaceholderMatch]:
        """移除重叠的匹配结果，优先保留更长的匹配"""
        if not matches:
            return []
        
        # 按起始位置排序，相同起始位置按长度降序
        sorted_matches = sorted(matches, key=lambda x: (x.start_pos, -(x.end_pos - x.start_pos)))
        
        result = []
        last_end = -1
        
        for match in sorted_matches:
            if match.start_pos >= last_end:
                result.append(match)
                last_end = match.end_pos
        
        return result
    
    def detect_with_fuzzy_matching(self, text: str, known_keys: List[str], 
                                   threshold: float = 0.8) -> List[PlaceholderMatch]:
        """
        使用模糊匹配检测可能的占位符
        
        Args:
            text: 要检测的文本
            known_keys: 已知的键名列表
            threshold: 相似度阈值
        
        Returns:
            检测到的占位符列表
        """
        matches = self.detect_all(text)
        
        # 对未匹配的文本进行模糊匹配
        # 这里可以实现更复杂的模糊匹配逻辑
        
        return matches
    
    def replace_placeholders(self, text: str, values: Dict[str, Any], 
                           keep_unmatched: bool = False) -> str:
        """
        替换文本中的所有占位符
        
        Args:
            text: 原始文本
            values: 键值映射
            keep_unmatched: 是否保留未匹配的占位符
        
        Returns:
            替换后的文本
        """
        matches = self.detect_all(text)
        
        # 从后往前替换，避免位置偏移
        result = text
        for match in reversed(matches):
            value = values.get(match.key)
            
            # 如果没有值，使用默认值
            if value is None and match.default_value is not None:
                value = match.default_value
            
            # 应用格式说明
            if value is not None and match.format_spec:
                value = self._apply_format(value, match.format_spec)
            
            if value is not None:
                result = result[:match.start_pos] + str(value) + result[match.end_pos:]
            elif not keep_unmatched:
                # 删除未匹配的占位符
                result = result[:match.start_pos] + result[match.end_pos:]
        
        return result
    
    def _apply_format(self, value: Any, format_spec: str) -> str:
        """应用格式说明"""
        try:
            if format_spec == 'upper':
                return str(value).upper()
            elif format_spec == 'lower':
                return str(value).lower()
            elif format_spec == 'title':
                return str(value).title()
            elif format_spec.startswith('date:'):
                # 日期格式化
                from datetime import datetime
                date_format = format_spec[5:]
                if isinstance(value, str):
                    # 尝试解析日期字符串
                    try:
                        dt = datetime.strptime(value, '%Y-%m-%d')
                        return dt.strftime(date_format)
                    except:
                        pass
                return str(value)
            elif format_spec.startswith('substr:'):
                # 子字符串
                params = format_spec[7:].split(',')
                start = int(params[0]) if params[0] else 0
                end = int(params[1]) if len(params) > 1 and params[1] else None
                return str(value)[start:end]
            else:
                # 尝试使用Python格式说明
                return format(value, format_spec)
        except:
            return str(value)
    
    def extract_keys(self, text: str) -> set:
        """提取文本中所有占位符的键名"""
        matches = self.detect_all(text)
        return {match.key for match in matches}
    
    def validate_keys(self, text: str, available_keys: set) -> Tuple[set, set]:
        """
        验证文本中的占位符键名
        
        Returns:
            (有效键名集合, 缺失键名集合)
        """
        found_keys = self.extract_keys(text)
        valid = found_keys & available_keys
        missing = found_keys - available_keys
        return valid, missing


class MultiLinePlaceholderHandler:
    """多行占位符处理器"""
    
    def __init__(self):
        self.detector = PlaceholderDetector()
    
    def detect_multiline(self, text: str) -> List[PlaceholderMatch]:
        """
        检测跨越多行的占位符
        支持格式：
        {{key::
        多行内容
        ::}}
        """
        pattern = r'\{\{\s*(\w+)::\s*\n(.*?)\n\s*::\}\}'
        matches = []
        
        for match in re.finditer(pattern, text, re.MULTILINE | re.DOTALL):
            matches.append(PlaceholderMatch(
                original=match.group(0),
                key=match.group(1).strip(),
                placeholder_type=PlaceholderType.DOUBLE_BRACES,
                start_pos=match.start(),
                end_pos=match.end(),
                default_value=match.group(2).strip()
            ))
        
        return matches
    
    def wrap_multiline_value(self, value: str, width: int = 80) -> str:
        """
        将多行值格式化为适合文档的格式
        
        Args:
            value: 原始值
            width: 行宽限制
        
        Returns:
            格式化后的文本
        """
        lines = value.split('\n')
        wrapped_lines = []
        
        for line in lines:
            if len(line) <= width:
                wrapped_lines.append(line)
            else:
                # 简单折行
                for i in range(0, len(line), width):
                    wrapped_lines.append(line[i:i+width])
        
        return '\n'.join(wrapped_lines)


# ==================== 便捷函数 ====================

def detect_placeholders(text: str) -> List[PlaceholderMatch]:
    """便捷函数：检测文本中的所有占位符"""
    detector = PlaceholderDetector()
    return detector.detect_all(text)

def replace_in_text(text: str, values: Dict[str, Any], keep_unmatched: bool = False) -> str:
    """便捷函数：替换文本中的占位符"""
    detector = PlaceholderDetector()
    return detector.replace_placeholders(text, values, keep_unmatched)

def extract_placeholder_keys(text: str) -> set:
    """便捷函数：提取所有占位符键名"""
    detector = PlaceholderDetector()
    return detector.extract_keys(text)


# ==================== 测试 ====================

if __name__ == "__main__":
    # 测试代码
    test_text = """
    案件编号：{bianh}
    受援人：{{weitr}}
    案由：[anyou]
    日期：{date?2024-01-01:date:%Y年%m月%d日}
    金额：{amount:,.2f}
    """
    
    detector = PlaceholderDetector()
    matches = detector.detect_all(test_text)
    
    print("检测到的占位符：")
    for match in matches:
        print(f"  {match.original} -> key='{match.key}', default={match.default_value}, format={match.format_spec}")
    
    values = {
        'bianh': '2024-001',
        'weitr': '张三',
        'anyou': '合同纠纷',
        'amount': 1234567.89
    }
    
    result = detector.replace_placeholders(test_text, values)
    print("\n替换结果：")
    print(result)
