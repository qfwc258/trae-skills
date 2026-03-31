#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板库管理工具
支持模板的增删改查操作
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List
from datetime import datetime


class TemplateManager:
    """模板库管理器"""
    
    def __init__(self, template_library_path: str = None):
        """
        初始化模板管理器
        
        Args:
            template_library_path: 模板库JSON文件路径
        """
        if template_library_path is None:
            # 默认路径
            template_library_path = Path(__file__).parent.parent / "references" / "template_library.json"
        
        self.template_library_path = Path(template_library_path)
        self.data = self._load_library()
    
    def _load_library(self) -> Dict:
        """加载模板库"""
        try:
            with open(self.template_library_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            # 创建新的模板库
            return {
                'version': '1.0',
                'last_updated': datetime.now().strftime('%Y-%m-%d'),
                'templates': []
            }
        except json.JSONDecodeError as e:
            print(f"错误：模板库JSON格式错误: {e}")
            return {
                'version': '1.0',
                'last_updated': datetime.now().strftime('%Y-%m-%d'),
                'templates': []
            }
    
    def _save_library(self):
        """保存模板库"""
        self.data['last_updated'] = datetime.now().strftime('%Y-%m-%d')
        
        with open(self.template_library_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        
        print(f"模板库已更新: {self.template_library_path}")
    
    def list_templates(self, domain: str = None, category: str = None):
        """
        列出所有模板
        
        Args:
            domain: 筛选领域
            category: 筛选类别
        """
        templates = self.data.get('templates', [])
        
        # 筛选
        if domain:
            templates = [t for t in templates if t.get('domain') == domain]
        if category:
            templates = [t for t in templates if t.get('category') == category]
        
        print(f"\n模板库列表 (共 {len(templates)} 个模板)")
        print("=" * 100)
        
        for i, template in enumerate(templates, 1):
            print(f"{i}. {template.get('name')}")
            print(f"   ID: {template.get('id')}")
            print(f"   领域: {template.get('domain')} | 类别: {template.get('category')}")
            print(f"   关键词: {', '.join(template.get('keywords', []))}")
            print()
    
    def add_template(self, template_file: str):
        """
        添加新模板
        
        Args:
            template_file: 模板JSON文件路径
        """
        try:
            with open(template_file, 'r', encoding='utf-8') as f:
                new_template = json.load(f)
        except FileNotFoundError:
            print(f"错误：模板文件不存在: {template_file}")
            return
        except json.JSONDecodeError as e:
            print(f"错误：模板文件JSON格式错误: {e}")
            return
        
        # 验证必要字段
        required_fields = ['id', 'name', 'domain', 'category']
        for field in required_fields:
            if field not in new_template:
                print(f"错误：模板缺少必要字段: {field}")
                return
        
        # 检查ID是否已存在
        templates = self.data.get('templates', [])
        if any(t.get('id') == new_template['id'] for t in templates):
            print(f"错误：模板ID已存在: {new_template['id']}")
            return
        
        # 添加模板
        templates.append(new_template)
        self.data['templates'] = templates
        self._save_library()
        
        print(f"模板已添加: {new_template['name']} (ID: {new_template['id']})")
    
    def update_template(self, template_id: str, template_file: str):
        """
        更新模板
        
        Args:
            template_id: 模板ID
            template_file: 新模板JSON文件路径
        """
        try:
            with open(template_file, 'r', encoding='utf-8') as f:
                updated_template = json.load(f)
        except FileNotFoundError:
            print(f"错误：模板文件不存在: {template_file}")
            return
        except json.JSONDecodeError as e:
            print(f"错误：模板文件JSON格式错误: {e}")
            return
        
        # 查找并更新
        templates = self.data.get('templates', [])
        found = False
        
        for i, template in enumerate(templates):
            if template.get('id') == template_id:
                # 保留ID
                updated_template['id'] = template_id
                templates[i] = updated_template
                found = True
                break
        
        if not found:
            print(f"错误：未找到模板ID: {template_id}")
            return
        
        self.data['templates'] = templates
        self._save_library()
        
        print(f"模板已更新: {updated_template.get('name')} (ID: {template_id})")
    
    def delete_template(self, template_id: str):
        """
        删除模板
        
        Args:
            template_id: 模板ID
        """
        templates = self.data.get('templates', [])
        original_count = len(templates)
        
        # 过滤掉要删除的模板
        templates = [t for t in templates if t.get('id') != template_id]
        
        if len(templates) == original_count:
            print(f"错误：未找到模板ID: {template_id}")
            return
        
        self.data['templates'] = templates
        self._save_library()
        
        print(f"模板已删除: {template_id}")
    
    def get_template(self, template_id: str) -> Dict:
        """
        获取模板详情
        
        Args:
            template_id: 模板ID
            
        Returns:
            模板详情
        """
        templates = self.data.get('templates', [])
        
        for template in templates:
            if template.get('id') == template_id:
                return template
        
        return None
    
    def export_template(self, template_id: str, output_file: str):
        """
        导出模板为JSON文件
        
        Args:
            template_id: 模板ID
            output_file: 输出文件路径
        """
        template = self.get_template(template_id)
        
        if not template:
            print(f"错误：未找到模板ID: {template_id}")
            return
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        
        print(f"模板已导出: {output_file}")
    
    def statistics(self):
        """统计模板库信息"""
        templates = self.data.get('templates', [])
        
        # 统计各领域模板数量
        domain_count = {}
        category_count = {}
        
        for template in templates:
            domain = template.get('domain', '未分类')
            category = template.get('category', '未分类')
            
            domain_count[domain] = domain_count.get(domain, 0) + 1
            category_count[category] = category_count.get(category, 0) + 1
        
        print("\n模板库统计信息")
        print("=" * 80)
        print(f"总模板数: {len(templates)}")
        print(f"最后更新: {self.data.get('last_updated')}")
        print(f"版本: {self.data.get('version')}")
        
        print("\n按领域统计:")
        for domain, count in sorted(domain_count.items()):
            print(f"  {domain}: {count}个模板")
        
        print("\n按类别统计:")
        for category, count in sorted(category_count.items()):
            print(f"  {category}: {count}个模板")


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='模板库管理工具')
    
    parser.add_argument('--action', choices=['list', 'add', 'update', 'delete', 'get', 'export', 'stats'],
                        required=True, help='操作类型')
    parser.add_argument('--template_library', type=str, help='模板库路径')
    parser.add_argument('--template_id', type=str, help='模板ID')
    parser.add_argument('--template_file', type=str, help='模板JSON文件路径')
    parser.add_argument('--output', type=str, help='输出文件路径')
    parser.add_argument('--domain', type=str, help='筛选领域')
    parser.add_argument('--category', type=str, help='筛选类别')
    
    args = parser.parse_args()
    
    # 初始化管理器
    manager = TemplateManager(args.template_library)
    
    # 执行操作
    if args.action == 'list':
        manager.list_templates(args.domain, args.category)
    
    elif args.action == 'add':
        if not args.template_file:
            print("错误：请提供 --template_file 参数")
            sys.exit(1)
        manager.add_template(args.template_file)
    
    elif args.action == 'update':
        if not args.template_id or not args.template_file:
            print("错误：请提供 --template_id 和 --template_file 参数")
            sys.exit(1)
        manager.update_template(args.template_id, args.template_file)
    
    elif args.action == 'delete':
        if not args.template_id:
            print("错误：请提供 --template_id 参数")
            sys.exit(1)
        manager.delete_template(args.template_id)
    
    elif args.action == 'get':
        if not args.template_id:
            print("错误：请提供 --template_id 参数")
            sys.exit(1)
        template = manager.get_template(args.template_id)
        if template:
            print(json.dumps(template, ensure_ascii=False, indent=2))
        else:
            print(f"未找到模板: {args.template_id}")
    
    elif args.action == 'export':
        if not args.template_id or not args.output:
            print("错误：请提供 --template_id 和 --output 参数")
            sys.exit(1)
        manager.export_template(args.template_id, args.output)
    
    elif args.action == 'stats':
        manager.statistics()


if __name__ == '__main__':
    main()
