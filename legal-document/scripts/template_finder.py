#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能模板查找器
优先从用户模板文件夹查找，其次使用内置模板库
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import jieba
from fuzzywuzzy import fuzz

# 导入其他处理器
from word_template_processor import WordTemplateProcessor
from template_matcher import TemplateMatcher


class TemplateFinder:
    """智能模板查找器"""
    
    def __init__(self, user_template_dir: str = None, builtin_template_path: str = None):
        """
        初始化模板查找器
        
        Args:
            user_template_dir: 用户模板文件夹路径
            builtin_template_path: 内置模板库路径
        """
        self.user_template_dir = Path(user_template_dir) if user_template_dir else None
        self.builtin_template_path = builtin_template_path
        
        # 初始化处理器
        self.word_processor = WordTemplateProcessor()
        self.template_matcher = TemplateMatcher(builtin_template_path)
        
        # 用户模板索引缓存
        self.user_template_index = None
        self.user_template_metadata = []
    
    def build_user_template_index(self) -> Dict:
        """
        构建用户模板索引
        
        Returns:
            索引信息字典
        """
        if not self.user_template_dir or not self.user_template_dir.exists():
            print(f"用户模板文件夹不存在: {self.user_template_dir}")
            return {'status': 'not_found', 'templates': []}
        
        print(f"正在扫描用户模板文件夹: {self.user_template_dir}")
        
        # 扫描所有Word文件
        word_files = list(self.user_template_dir.glob('*.docx')) + \
                     list(self.user_template_dir.glob('*.doc'))
        
        if not word_files:
            print("用户模板文件夹中未找到Word文件")
            return {'status': 'empty', 'templates': []}
        
        # 提取每个文件的元数据
        self.user_template_metadata = []
        
        for word_file in word_files:
            try:
                print(f"  正在解析: {word_file.name}")
                
                # 读取模板信息
                template_info = self.word_processor.read_template(str(word_file))
                
                # 提取关键词（从文件名和标题）
                keywords = self._extract_keywords_from_filename(word_file.name)
                keywords.extend(self._extract_keywords_from_title(template_info['title']))
                
                metadata = {
                    'file_path': str(word_file),
                    'file_name': word_file.name,
                    'title': template_info['title'],
                    'keywords': list(set(keywords)),  # 去重
                    'placeholders': template_info['placeholders'],
                    'type': template_info['structure']['type'],
                    'paragraph_count': len(template_info['paragraphs']),
                    'table_count': len(template_info['tables'])
                }
                
                self.user_template_metadata.append(metadata)
                
            except Exception as e:
                print(f"  警告：解析失败 {word_file.name}: {e}")
                continue
        
        # 构建索引
        index_data = {
            'status': 'success',
            'template_count': len(self.user_template_metadata),
            'templates': self.user_template_metadata,
            'indexed_at': self._get_current_time()
        }
        
        # 保存索引
        index_file = self.user_template_dir / '.template_index.json'
        with open(index_file, 'w', encoding='utf-8') as f:
            json.dump(index_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n索引构建完成:")
        print(f"  模板数量: {len(self.user_template_metadata)}")
        print(f"  索引文件: {index_file}")
        
        self.user_template_index = index_data
        return index_data
    
    def load_user_template_index(self) -> Dict:
        """
        加载已存在的用户模板索引
        
        Returns:
            索引信息
        """
        if not self.user_template_dir:
            return {'status': 'not_configured', 'templates': []}
        
        index_file = self.user_template_dir / '.template_index.json'
        
        if index_file.exists():
            with open(index_file, 'r', encoding='utf-8') as f:
                self.user_template_index = json.load(f)
                self.user_template_metadata = self.user_template_index.get('templates', [])
                return self.user_template_index
        
        # 索引不存在，重新构建
        return self.build_user_template_index()
    
    def find_template(self, keywords: List[str] = None, case_description: str = None, 
                      top_n: int = 3) -> Tuple[Dict, str]:
        """
        查找最佳匹配模板（优先用户模板）
        
        Args:
            keywords: 关键词列表
            case_description: 案情描述
            top_n: 返回前N个结果
            
        Returns:
            (匹配结果, 来源标识)
            来源标识: 'user' 或 'builtin'
        """
        # 步骤1：在用户模板中查找
        if self.user_template_dir:
            print(f"\n步骤1：在用户模板文件夹中查找...")
            
            # 加载索引
            if not self.user_template_index:
                self.load_user_template_index()
            
            if self.user_template_metadata:
                user_results = self._search_in_user_templates(keywords, case_description, top_n)
                
                if user_results:
                    best_match = user_results[0]
                    if best_match['match_score'] > 60:  # 匹配度阈值
                        print(f"✓ 找到匹配的用户模板: {best_match['file_name']}")
                        print(f"  匹配度: {best_match['match_score']}%")
                        return best_match, 'user'
        
        # 步骤2：在内置模板库中查找
        print(f"\n步骤2：在内置模板库中查找...")
        
        if keywords:
            builtin_results = self.template_matcher.match_by_keywords(keywords, top_n)
        elif case_description:
            builtin_results = self.template_matcher.recommend_by_case(case_description, top_n)
        else:
            return None, 'none'
        
        if builtin_results:
            best_match = builtin_results[0]
            print(f"✓ 找到匹配的内置模板: {best_match['template_name']}")
            print(f"  匹配度: {best_match['similarity_score']}%")
            return best_match, 'builtin'
        
        # 步骤3：未找到匹配模板
        print(f"\n⚠ 未找到匹配模板，建议使用AI生成")
        return None, 'none'
    
    def _search_in_user_templates(self, keywords: List[str], case_description: str, 
                                   top_n: int) -> List[Dict]:
        """
        在用户模板中搜索
        
        Args:
            keywords: 关键词列表
            case_description: 案情描述
            
        Returns:
            匹配结果列表
        """
        results = []
        
        # 提取搜索关键词
        search_keywords = keywords if keywords else []
        if case_description:
            words = jieba.cut(case_description)
            search_keywords.extend([w for w in words if len(w) > 1])
        
        search_keywords = list(set(search_keywords))
        
        # 在每个用户模板中匹配
        for template in self.user_template_metadata:
            score = 0
            matched_keywords = []
            
            # 计算匹配度
            for keyword in search_keywords:
                keyword_lower = keyword.lower()
                
                # 文件名匹配
                if keyword_lower in template['file_name'].lower():
                    score += 100
                    matched_keywords.append(keyword)
                
                # 标题匹配
                elif keyword_lower in template['title'].lower():
                    score += 90
                    matched_keywords.append(keyword)
                
                # 关键词列表匹配
                elif keyword_lower in [k.lower() for k in template['keywords']]:
                    score += 80
                    matched_keywords.append(keyword)
                
                # 占位符匹配
                elif keyword_lower in [p.lower() for p in template['placeholders']]:
                    score += 60
                    matched_keywords.append(keyword)
                
                # 模糊匹配
                else:
                    for kw in template['keywords']:
                        fuzzy_score = fuzz.ratio(keyword_lower, kw.lower())
                        if fuzzy_score > 70:
                            score += fuzzy_score * 0.5
                            matched_keywords.append(keyword)
                            break
            
            # 计算平均得分
            if matched_keywords:
                avg_score = score / len(search_keywords) if search_keywords else 0
            else:
                avg_score = 0
            
            if avg_score > 0:
                results.append({
                    'file_path': template['file_path'],
                    'file_name': template['file_name'],
                    'title': template['title'],
                    'keywords': template['keywords'],
                    'placeholders': template['placeholders'],
                    'type': template['type'],
                    'match_score': round(avg_score, 2),
                    'matched_keywords': matched_keywords
                })
        
        # 按匹配度排序
        results.sort(key=lambda x: x['match_score'], reverse=True)
        
        return results[:top_n]
    
    def _extract_keywords_from_filename(self, filename: str) -> List[str]:
        """从文件名提取关键词"""
        # 移除扩展名
        name = Path(filename).stem
        
        # 使用jieba分词
        words = jieba.cut(name)
        keywords = [w for w in words if len(w) > 1]
        
        return keywords
    
    def _extract_keywords_from_title(self, title: str) -> List[str]:
        """从标题提取关键词"""
        words = jieba.cut(title)
        keywords = [w for w in words if len(w) > 1]
        return keywords
    
    def _get_current_time(self) -> str:
        """获取当前时间"""
        from datetime import datetime
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    def list_user_templates(self) -> List[Dict]:
        """
        列出所有用户模板
        
        Returns:
            模板列表
        """
        if not self.user_template_index:
            self.load_user_template_index()
        
        return self.user_template_metadata
    
    def get_template_info(self, template_path: str, source: str = 'user') -> Dict:
        """
        获取模板详细信息
        
        Args:
            template_path: 模板路径
            source: 来源（user/builtin）
            
        Returns:
            模板详细信息
        """
        if source == 'user':
            return self.word_processor.read_template(template_path)
        else:
            template_id = template_path  # 内置模板使用ID
            return self.template_matcher.get_template_by_id(template_id)
    
    def rebuild_index(self):
        """重建用户模板索引"""
        return self.build_user_template_index()


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='智能模板查找器')
    
    parser.add_argument('--action', choices=['find', 'index', 'list'], required=True,
                        help='操作类型：find=查找模板，index=构建索引，list=列出模板')
    parser.add_argument('--user_template_dir', type=str, default='./templates',
                        help='用户模板文件夹路径')
    parser.add_argument('--keywords', nargs='+', help='关键词列表')
    parser.add_argument('--case_description', type=str, help='案情描述')
    parser.add_argument('--top_n', type=int, default=3, help='返回结果数量')
    parser.add_argument('--builtin_template', type=str, help='内置模板库路径')
    
    args = parser.parse_args()
    
    # 初始化查找器
    finder = TemplateFinder(args.user_template_dir, args.builtin_template)
    
    if args.action == 'index':
        # 构建索引
        finder.build_user_template_index()
    
    elif args.action == 'list':
        # 列出模板
        templates = finder.list_user_templates()
        print(f"\n用户模板列表 (共 {len(templates)} 个):")
        print("=" * 80)
        for i, template in enumerate(templates, 1):
            print(f"{i}. {template['file_name']}")
            print(f"   标题: {template['title']}")
            print(f"   类型: {template['type']}")
            print(f"   关键词: {', '.join(template['keywords'][:5])}")
            print(f"   占位符数量: {len(template['placeholders'])}")
            print()
    
    elif args.action == 'find':
        # 查找模板
        result, source = finder.find_template(args.keywords, args.case_description, args.top_n)
        
        if result:
            print(f"\n最佳匹配模板:")
            print(f"  来源: {'用户模板' if source == 'user' else '内置模板库'}")
            
            if source == 'user':
                print(f"  文件名: {result['file_name']}")
                print(f"  标题: {result['title']}")
                print(f"  路径: {result['file_path']}")
                print(f"  匹配度: {result['match_score']}%")
                print(f"  占位符: {', '.join(result['placeholders'][:5])}")
            else:
                print(f"  模板ID: {result['template_id']}")
                print(f"  名称: {result['template_name']}")
                print(f"  领域: {result['domain']}")
                print(f"  匹配度: {result['similarity_score']}%")
        else:
            print("\n未找到匹配模板，建议使用AI生成")


if __name__ == '__main__':
    main()
