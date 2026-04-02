---
name: contract-review
description: Contract review skill with comment-based annotations, three-layer review (basic, business, legal), risk levels, and Mermaid flowchart generation.
---

# Contract Review Skill

## 功能
- 仅添加注释批注，不修改原文
- 三层审查：基础、业务、法律
- 风险等级评估（高/中/低）
- 生成业务流程序列图

## 审查标准
- Layer 1 (Basic): 数字、日期、术语准确性
- Layer 2 (Business): 范围、交付、定价、付款
- Layer 3 (Legal): 效力、终止、责任、争议解决

## 输出
- 标注版合同：{ContractName}_审核版.docx
- 审核报告：审核报告.txt
- 业务流程图（Mermaid）