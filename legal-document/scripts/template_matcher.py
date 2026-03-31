#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模板匹配与推荐引擎
支持关键词匹配和AI案情解析两种模式
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict, Tuple
import jieba
from fuzzywuzzy import fuzz


class TemplateMatcher:
    """模板匹配引擎"""
    
    def __init__(self, template_library_path: str = None):
        """
        初始化模板匹配器
        
        Args:
            template_library_path: 模板库JSON文件路径
        """
        if template_library_path is None:
            # 默认路径
            template_library_path = Path(__file__).parent.parent / "references" / "template_library.json"
        
        self.template_library_path = Path(template_library_path)
        self.templates = self._load_templates()
    
    def _load_templates(self) -> List[Dict]:
        """加载模板库"""
        try:
            with open(self.template_library_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('templates', [])
        except FileNotFoundError:
            print(f"错误：模板库文件不存在: {self.template_library_path}")
            return []
        except json.JSONDecodeError as e:
            print(f"错误：模板库JSON格式错误: {e}")
            return []
    
    def match_by_keywords(self, keywords: List[str], top_n: int = 5) -> List[Dict]:
        """
        基于关键词匹配模板
        
        Args:
            keywords: 关键词列表
            top_n: 返回前N个匹配结果
            
        Returns:
            匹配结果列表，每项包含模板信息和相似度得分
        """
        results = []
        
        for template in self.templates:
            # 计算关键词匹配得分
            template_keywords = template.get('keywords', [])
            template_name = template.get('name', '')
            
            # 多维度匹配
            score = 0
            matched_keywords = []
            
            for keyword in keywords:
                keyword = keyword.lower()
                
                # 1. 关键词列表完全匹配
                if keyword in [k.lower() for k in template_keywords]:
                    score += 100
                    matched_keywords.append(keyword)
                
                # 2. 模板名称包含关键词
                elif keyword in template_name.lower():
                    score += 80
                    matched_keywords.append(keyword)
                
                # 3. 模糊匹配
                else:
                    for tk in template_keywords:
                        fuzzy_score = fuzz.ratio(keyword, tk.lower())
                        if fuzzy_score > 70:
                            score += fuzzy_score * 0.5
                            matched_keywords.append(keyword)
                            break
            
            # 计算平均得分
            if matched_keywords:
                avg_score = score / len(keywords)
            else:
                avg_score = 0
            
            if avg_score > 0:
                results.append({
                    'template_id': template.get('id'),
                    'template_name': template_name,
                    'domain': template.get('domain'),
                    'category': template.get('category'),
                    'similarity_score': round(avg_score, 2),
                    'matched_keywords': matched_keywords,
                    'required_elements': template.get('elements', [])
                })
        
        # 按相似度排序
        results.sort(key=lambda x: x['similarity_score'], reverse=True)
        
        return results[:top_n]
    
    def recommend_by_case(self, case_description: str, top_n: int = 3) -> List[Dict]:
        """
        基于案情描述推荐模板
        
        Args:
            case_description: 案情描述文本
            top_n: 返回前N个推荐结果
            
        Returns:
            推荐结果列表
        """
        # 分词提取关键词
        words = jieba.cut(case_description)
        keywords = [w for w in words if len(w) > 1]
        
        # 识别法律领域关键词
        domain_keywords = {
            '民事': ['借款', '欠款', '合同', '房产', '婚姻', '继承', '侵权', '赔偿'],
            '商事': ['公司', '股权', '投资', '合伙', '破产', '保险'],
            '刑事': ['犯罪', '盗窃', '诈骗', '伤害', '走私', '毒品'],
            '行政': ['政府', '处罚', '许可', '拆迁', '行政复议'],
            '知识产权': ['专利', '商标', '著作权', '侵权'],
            '劳动': ['工资', '劳动合同', '工伤', '社保', '离职'],
            '海事': ['船舶', '海上', '运输', '海难'],
            '金融': ['银行', '证券', '期货', '基金'],
            '执行': ['强制执行', '财产', '查封', '拍卖']
        }
        
        # 识别领域
        detected_domains = []
        for domain, d_keywords in domain_keywords.items():
            if any(kw in case_description for kw in d_keywords):
                detected_domains.append(domain)
        
        # 基于关键词和领域匹配
        results = []
        for template in self.templates:
            template_domain = template.get('domain', '')
            template_keywords = template.get('keywords', [])
            
            score = 0
            
            # 领域匹配加分
            if template_domain in detected_domains:
                score += 50
            
            # 关键词匹配
            for keyword in keywords:
                if keyword in template_keywords:
                    score += 20
                elif keyword in template.get('name', ''):
                    score += 15
            
            # 模糊匹配
            for tk in template_keywords:
                fuzzy_score = fuzz.partial_ratio(case_description, tk)
                if fuzzy_score > 60:
                    score += fuzzy_score * 0.3
            
            if score > 0:
                results.append({
                    'template_id': template.get('id'),
                    'template_name': template.get('name'),
                    'domain': template_domain,
                    'category': template.get('category'),
                    'recommendation_score': round(score, 2),
                    'matched_elements': self._extract_case_elements(case_description, template),
                    'required_elements': template.get('elements', [])
                })
        
        # 排序并返回
        results.sort(key=lambda x: x['recommendation_score'], reverse=True)
        
        return results[:top_n]
    
    def _extract_case_elements(self, case_description: str, template: Dict) -> Dict:
        """
        从案情描述中提取已知的要素
        
        Args:
            case_description: 案情描述
            template: 模板信息
            
        Returns:
            已提取的要素
        """
        elements = template.get('elements', [])
        extracted = {}
        
        # 简单的关键词匹配提取（实际应用中可使用更复杂的NLP技术）
        for element in elements:
            element_name = element.get('name', '')
            element_keywords = element.get('keywords', [])
            
            for keyword in element_keywords:
                if keyword in case_description:
                    extracted[element_name] = {
                        'detected': True,
                        'keyword': keyword
                    }
                    break
        
        return extracted
    
    def get_template_by_id(self, template_id: str) -> Dict:
        """
        根据模板ID获取模板详情
        
        Args:
            template_id: 模板ID
            
        Returns:
            模板详情
        """
        for template in self.templates:
            if template.get('id') == template_id:
                return template
        
        return None


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='法律文书模板匹配与推荐')
    
    parser.add_argument('--keywords', nargs='+', help='关键词列表')
    parser.add_argument('--case_description', type=str, help='案情描述')
    parser.add_argument('--mode', choices=['match', 'recommend'], default='match',
                        help='匹配模式：match=关键词匹配，recommend=AI推荐')
    parser.add_argument('--top_n', type=int, default=5, help='返回结果数量')
    parser.add_argument('--template_library', type=str, help='模板库路径')
    parser.add_argument('--template_id', type=str, help='模板ID（用于查询详情）')
    
    args = parser.parse_args()
    
    # 初始化匹配器
    matcher = TemplateMatcher(args.template_library)
    
    # 查询模板详情
    if args.template_id:
        template = matcher.get_template_by_id(args.template_id)
        if template:
            print(json.dumps(template, ensure_ascii=False, indent=2))
        else:
            print(f"未找到模板: {args.template_id}")
        return
    
    # 关键词匹配
    if args.keywords:
        results = matcher.match_by_keywords(args.keywords, args.top_n)
        print(f"\n关键词匹配结果 (关键词: {', '.join(args.keywords)}):")
        print("=" * 80)
        for i, result in enumerate(results, 1):
            print(f"\n{i}. {result['template_name']}")
            print(f"   模板ID: {result['template_id']}")
            print(f"   领域: {result['domain']} | 类别: {result['category']}")
            print(f"   匹配度: {result['similarity_score']}%")
            print(f"   匹配关键词: {', '.join(result['matched_keywords'])}")
            print(f"   需补充要素: {len(result['required_elements'])}项")
    
    # 案情推荐
    elif args.case_description:
        results = matcher.recommend_by_case(args.case_description, args.top_n)
        print(f"\nAI推荐结果 (案情: {args.case_description[:50]}...):")
        print("=" * 80)
        for i, result in enumerate(results, 1):
            print(f"\n{i}. {result['template_name']}")
            print(f"   模板ID: {result['template_id']}")
            print(f"   领域: {result['domain']} | 类别: {result['category']}")
            print(f"   推荐度: {result['recommendation_score']}")
            print(f"   已识别要素: {len(result['matched_elements'])}项")
            print(f"   需补充要素: {len(result['required_elements'])}项")
    
    else:
        print("请提供 --keywords 或 --case_description 参数")
        parser.print_help()
        sys.exit(1)


if __name__ == '__main__':
    main()
