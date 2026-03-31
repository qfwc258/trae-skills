#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
法条检索工具
根据文书类型或案情检索相关法律条文
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict
import jieba
from fuzzywuzzy import fuzz


class LawRetriever:
    """法条检索引擎"""
    
    def __init__(self, law_database_path: str = None):
        """
        初始化法条检索器
        
        Args:
            law_database_path: 法条数据库JSON文件路径
        """
        if law_database_path is None:
            # 默认路径
            law_database_path = Path(__file__).parent.parent / "references" / "law_database.json"
        
        self.law_database_path = Path(law_database_path)
        self.laws = self._load_laws()
    
    def _load_laws(self) -> List[Dict]:
        """加载法条数据库"""
        try:
            with open(self.law_database_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('laws', [])
        except FileNotFoundError:
            print(f"错误：法条数据库文件不存在: {self.law_database_path}")
            return []
        except json.JSONDecodeError as e:
            print(f"错误：法条数据库JSON格式错误: {e}")
            return []
    
    def retrieve_by_keywords(self, keywords: List[str], top_n: int = 10) -> List[Dict]:
        """
        根据关键词检索法条
        
        Args:
            keywords: 关键词列表
            top_n: 返回前N条结果
            
        Returns:
            匹配的法条列表
        """
        results = []
        
        for law in self.laws:
            law_name = law.get('name', '')
            articles = law.get('articles', [])
            
            for article in articles:
                article_number = article.get('number', '')
                article_content = article.get('content', '')
                applicable_scenarios = article.get('applicable_scenarios', [])
                
                # 计算匹配得分
                score = 0
                matched_keywords = []
                
                for keyword in keywords:
                    keyword = keyword.lower()
                    
                    # 法条内容匹配
                    if keyword in article_content.lower():
                        score += 100
                        matched_keywords.append(keyword)
                    
                    # 适用场景匹配
                    for scenario in applicable_scenarios:
                        if keyword in scenario.lower():
                            score += 80
                            matched_keywords.append(keyword)
                            break
                    
                    # 模糊匹配
                    fuzzy_score = fuzz.partial_ratio(keyword, article_content)
                    if fuzzy_score > 70:
                        score += fuzzy_score * 0.5
                
                if score > 0:
                    results.append({
                        'law_name': law_name,
                        'article_number': article_number,
                        'content': article_content,
                        'applicable_scenarios': applicable_scenarios,
                        'match_score': round(score, 2),
                        'matched_keywords': list(set(matched_keywords))
                    })
        
        # 按匹配度排序
        results.sort(key=lambda x: x['match_score'], reverse=True)
        
        return results[:top_n]
    
    def retrieve_by_document_type(self, document_type: str, top_n: int = 10) -> List[Dict]:
        """
        根据文书类型检索相关法条
        
        Args:
            document_type: 文书类型（如"民事起诉状"、"劳动合同"等）
            top_n: 返回前N条结果
            
        Returns:
            相关法条列表
        """
        # 文书类型与法条映射关系
        type_law_mapping = {
            '民事起诉状': ['民事诉讼法', '民法典'],
            '答辩状': ['民事诉讼法', '民法典'],
            '劳动合同': ['劳动合同法', '劳动法', '民法典'],
            '房屋租赁合同': ['民法典', '城市房地产管理法'],
            '借款合同': ['民法典', '民间借贷司法解释'],
            '离婚协议': ['民法典', '婚姻法'],
            '遗嘱': ['民法典', '继承法'],
            '刑事起诉书': ['刑法', '刑事诉讼法'],
            '行政复议申请书': ['行政复议法', '行政诉讼法']
        }
        
        # 识别相关法律
        related_laws = []
        for doc_type, laws in type_law_mapping.items():
            if doc_type in document_type or fuzz.ratio(doc_type, document_type) > 70:
                related_laws = laws
                break
        
        # 检索相关法条
        results = []
        for law_name in related_laws:
            for law in self.laws:
                if law_name in law.get('name', ''):
                    articles = law.get('articles', [])
                    for article in articles:
                        # 优先返回重点条文
                        if article.get('is_important', False):
                            results.append({
                                'law_name': law.get('name'),
                                'article_number': article.get('number'),
                                'content': article.get('content'),
                                'applicable_scenarios': article.get('applicable_scenarios', []),
                                'is_important': True
                            })
        
        return results[:top_n]
    
    def retrieve_by_case(self, case_description: str, top_n: int = 10) -> List[Dict]:
        """
        根据案情描述检索相关法条
        
        Args:
            case_description: 案情描述
            top_n: 返回前N条结果
            
        Returns:
            相关法条列表
        """
        # 分词提取关键词
        words = jieba.cut(case_description)
        keywords = [w for w in words if len(w) > 1]
        
        # 调用关键词检索
        return self.retrieve_by_keywords(keywords, top_n)
    
    def get_article_by_id(self, law_name: str, article_number: str) -> Dict:
        """
        获取特定法条详情
        
        Args:
            law_name: 法律名称
            article_number: 法条编号
            
        Returns:
            法条详情
        """
        for law in self.laws:
            if law_name in law.get('name', ''):
                for article in law.get('articles', []):
                    if article.get('number') == article_number:
                        return {
                            'law_name': law.get('name'),
                            'article_number': article.get('number'),
                            'content': article.get('content'),
                            'applicable_scenarios': article.get('applicable_scenarios', [])
                        }
        
        return None


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='法条检索工具')
    
    parser.add_argument('--keywords', nargs='+', help='关键词列表')
    parser.add_argument('--document_type', type=str, help='文书类型')
    parser.add_argument('--case_description', type=str, help='案情描述')
    parser.add_argument('--top_n', type=int, default=10, help='返回结果数量')
    parser.add_argument('--law_database', type=str, help='法条数据库路径')
    parser.add_argument('--output', type=str, help='输出文件路径（JSON格式）')
    
    args = parser.parse_args()
    
    # 初始化检索器
    retriever = LawRetriever(args.law_database)
    
    results = []
    
    # 关键词检索
    if args.keywords:
        results = retriever.retrieve_by_keywords(args.keywords, args.top_n)
        print(f"\n关键词检索结果 (关键词: {', '.join(args.keywords)}):")
        print("=" * 80)
    
    # 文书类型检索
    elif args.document_type:
        results = retriever.retrieve_by_document_type(args.document_type, args.top_n)
        print(f"\n文书类型检索结果 (文书类型: {args.document_type}):")
        print("=" * 80)
    
    # 案情检索
    elif args.case_description:
        results = retriever.retrieve_by_case(args.case_description, args.top_n)
        print(f"\n案情检索结果 (案情: {args.case_description[:50]}...):")
        print("=" * 80)
    
    else:
        print("请提供 --keywords、--document_type 或 --case_description 参数")
        parser.print_help()
        sys.exit(1)
    
    # 输出结果
    for i, result in enumerate(results, 1):
        print(f"\n{i}. 《{result['law_name']}》{result['article_number']}")
        print(f"   内容: {result['content'][:100]}...")
        if result.get('applicable_scenarios'):
            print(f"   适用场景: {', '.join(result['applicable_scenarios'][:3])}")
        if result.get('matched_keywords'):
            print(f"   匹配关键词: {', '.join(result['matched_keywords'])}")
    
    # 保存到文件
    if args.output:
        output_data = {
            'query_type': 'keywords' if args.keywords else ('document_type' if args.document_type else 'case'),
            'query_content': args.keywords or args.document_type or args.case_description,
            'total_results': len(results),
            'results': results
        }
        
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n结果已保存到: {args.output}")


if __name__ == '__main__':
    main()
