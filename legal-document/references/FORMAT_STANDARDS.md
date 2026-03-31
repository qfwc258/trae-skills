# 法律文书格式规范总览

## 概述

本文档汇总所有法律文书的格式规范，包括字体、段落结构、占位符定义等。

## 通用规范

### 占位符替换规则
| 场景 | 替换内容 |
|------|----------|
| 正式生成文书 | 占位符替换为实际内容 |
| 生成空白模板 | 占位符替换为8个空格（`        `） |

### 字体规范
| 元素 | 字体 | 字号 |
|------|------|------|
| 主标题 | 黑体/方正小标宋简体 | 二号（203200 EMU） |
| 正文 | 仿宋_GB2312 | 三号/四号 |
| 标签加粗 | 仿宋_GB2312 | 三号，粗体 |

---

## 文书类型索引

### 刑事类
| 文书 | 规范文档 | 模板文件 |
|------|----------|----------|
| 刑事法援会见前打印 | [刑事法援会见前打印格式规范.md](./刑事法援会见前打印格式规范.md) | `8 刑事法援会见前打印.docx` |
| 刑事法律援助案卷归档目录 | [法律援助案卷归档目录格式规范.md](./法律援助案卷归档目录格式规范.md) | `8 湖南省法律援助案卷归档目录（刑事）.docx` |

### 民事类
| 文书 | 规范文档 | 模板文件 |
|------|----------|----------|
| 调证材料 | [调证材料格式规范.md](./调证材料格式规范.md) | `2 9 调证材料.docx` |
| 证据目录 | [证据目录格式规范.md](./证据目录格式规范.md) | `3 证据目录 .docx` |
| 民事法律援助案卷归档目录 | [法律援助案卷归档目录格式规范.md](./法律援助案卷归档目录格式规范.md) | `9 湖南省法律援助案卷归档目录（民事）.docx` |

---

## 通用占位符列表

| 占位符 | 含义 | 适用文书 |
|--------|------|----------|
| weitr | 受援人/委托人姓名 | 所有文书 |
| wtrsfz | 受援人/委托人身份证号 | 委托代理类文书 |
| wtrzz | 受援人住所/羁押地 | 刑事法援会见前打印 |
| anyou | 涉嫌罪名/案由 | 刑事类文书 |
| jied | 案件阶段（侦查/审查起诉/审判） | 刑事类文书 |
| wtxm | 委托项目编号 | 委托代理/辩护协议 |
| quanx | 代理权限 | 委托代理类文书 |
| hjdd | 会见地点 | 刑事法援会见前打印 |
| dfdsr | 对方当事人 | 调证材料 |
| dcdw | 调取单位 | 调证材料 |
| bdqxx | 被调取信息内容 | 调证材料 |

---

## 快速生成命令

### 空白模板生成
```bash
# 生成所有空白模板
python scripts/create_all_blank_templates.py all

# 按类别生成
python scripts/create_all_blank_templates.py criminal  # 刑事类
python scripts/create_all_blank_templates.py civil     # 民事类
```

### 单个空白模板生成
```bash
python scripts/create_blank_criminal_legal_aid_meeting.py
python scripts/create_blank_investigation_doc.py
python scripts/create_blank_evidence_directory.py
python scripts/create_blank_civil_legal_aid_directory.py
python scripts/create_blank_criminal_legal_aid_directory.py
```

---

## 更新日志

| 日期 | 版本 | 更新内容 |
|------|------|----------|
| 2025-03-22 | v1.1 | 统一占位符为空8格空格，规范空白模板生成 |
| 2025-03-22 | v1.0 | 初始版本 |