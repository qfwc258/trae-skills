---
name: "legal-aid-doc"
description: "Generates legal aid (法援) closure documents from element files. Invoke when user has 元素_*.txt and needs to fill in multiple Word templates while preserving original formatting."
---

# Legal Aid Document Generator (法律援助文书生成器)

## Purpose

根据《元素_*.txt》中的案件数据，自动批量填充多个Word文档模板，**保持原文档格式不变**。

## Input Files

### Required: 元素文件 (Element File)
位置: `d:\trae\法律援助文书\元素_民事.txt` (或其他 `元素_*.txt`)

格式: Key-value pairs separated by `::`
```txt
--委托授权--
leib::民事
anyou::劳务合同纠纷
weitr::刘新林
dfdsr::资建文、湖南泰荣建设工程有限公司
bianh::(2026)湘援0124民2号
badw::宁乡市人民法院
jied::一审
zprq::2025年10月13日

--办案过程--
gcsj1::2026年1月5日
gcfs1::直接
gcnr1::接受宁乡市法律援助中心的指派，领取案卷。
...
```

### Required: 文档模板 (Document Templates)
位置: `d:\trae\法律援助文书\*.docx`

支持多个模板文件，例如：
- `法援结案.docx`
- `阅卷笔录.docx`
- `承办情况小结.docx`
- `立卷说明.docx`
- `举证提纲.docx`
- `代理意见.docx`

**注意**: 模板中需要包含与元素文件对应的占位符（如 `bianh`, `weitr`, `gcsj1` 等）

## Output

生成填充后的文档（每个模板生成一个对应的已填充版本）：
- `法援结案_已填充.docx`
- `阅卷笔录_已填充.docx`
- `承办情况小结_已填充.docx`
- ...

## 支持的占位符格式

### 标准占位符格式
工具支持多种占位符格式，模板可以使用以下任意一种：

| 格式 | 示例 | 说明 |
|------|------|------|
| `{key}` | `{bianh}` | 花括号格式 |
| `{{key}}` | `{{weitr}}` | 双花括号格式（推荐） |
| `[key]` | `[anyou]` | 方括号格式 |
| `[[key]]` | `[[badw]]` | 双方括号格式 |
| `【key】` | `【jied】` | 中文方头括号 |
| `<key>` | `<zprq>` | 尖括号格式 |
| `<<key>>` | `<<jarq>>` | 双尖括号格式 |
| `#key#` | `#wtrsfz#` | 井号格式 |
| `$key$` | `$wtrdh$` | 美元符号格式 |

### 高级占位符特性

#### 1. 默认值
当元素文件中缺少某个键时，可以使用默认值：
```
{date?2024-01-01}      → 如果date未定义，使用2024-01-01
{{anyou?合同纠纷}}      → 如果anyou未定义，使用"合同纠纷"
```

#### 2. 格式说明
支持对值进行格式化：
```
{amount:,.2f}          → 数字格式化：1,234,567.89
{name:upper}           → 转大写：ZHANG SAN
{date:date:%Y年%m月%d日} → 日期格式化：2024年01月15日
{title:title}          → 首字母大写
```

#### 3. 条件占位符
组合使用默认值和格式：
```
{jarq?未结案:date:%Y年%m月%d日}
```

### 标准占位符列表

#### 基本信息
| 占位符 | 说明 |
|--------|------|
| bianh | 案件编号 |
| leib | 案件类别(民事/刑事/行政) |
| anyou | 案由 |
| weitr | 受援人 |
| dfdsr | 对方当事人 |
| badw | 办案机关/法院 |
| jied | 所处阶段 |
| zprq | 指派日期 |
| jarq | 结案日期 |

#### 委托人信息
| 占位符 | 说明 |
|--------|------|
| wtrsfz | 身份证号 |
| wtrdh | 电话 |
| wtrxb | 性别 |
| wtrcs | 出生日期 |
| wtrzz | 住址 |

#### 办案过程 (1-9)
| 占位符 | 说明 |
|--------|------|
| gcsj1-9 | 办案时间 |
| gcfs1-9 | 办案方式 |
| gcnr1-9 | 办案内容 |

#### 其他
| 占位符 | 说明 |
|--------|------|
| cbxj | 承办情况小结 |
| ljsm | 立卷说明 |
| yjrq | 阅卷日期 |
| dcdw | 调证单位 |
| bdqxx | 被告户籍信息 |

## 技术实现

基于 `python-docx` 库，完整保持原格式：
1. **字符级格式保持** - 记录每个字符的完整格式（字体、大小、颜色等）
2. **XML级字体处理** - 通过 `qn('w:rFonts')` 直接操作XML
3. **多选框保护** - 特殊处理 `□☑` 等符号
4. **表格行高设置** - 允许跨页断行
5. **从后往前替换** - 避免位置偏移

## Usage

### 基本用法（批量处理所有模板）
```bash
cd d:\trae\.trae\skills\legal-aid-doc
python fill_docx_pro.py
```

### 指定目录
```bash
python fill_docx_pro.py "d:\其他案件目录"
```

### 指定元素文件
```bash
python fill_docx_pro.py "d:\案件目录" "元素_刑事.txt"
```

### 启用增强占位符检测（优化1）
```bash
# 使用增强模式（支持多种占位符格式、默认值、格式说明）
python fill_docx_pro.py "d:\案件目录" "元素_民事.txt" "enhanced"

# 或同时启用智能填充和增强检测
python fill_docx_pro.py "d:\案件目录" "元素_民事.txt" "smart" "true"
```

### 在Python中调用
```python
from fill_docx_pro import fill_multiple_documents, fill_single_document

# 批量填充（基础模式）
fill_multiple_documents(
    input_dir="d:\\法律援助文书",
    element_file="元素_民事.txt",
    template_pattern="*.docx"
)

# 批量填充（增强模式 - 支持多种占位符格式）
fill_multiple_documents(
    input_dir="d:\\法律援助文书",
    element_file="元素_民事.txt",
    template_pattern="*.docx",
    use_enhanced=True  # 启用增强占位符检测
)

# 单个文档
fill_single_document(
    template_path="法援结案.docx",
    element_file="元素_民事.txt",
    output_path="法援结案_已填充.docx",
    use_enhanced=True
)
```

## 支持的案件类型（优化2）

工具自动识别并支持以下案件类型：

| 案件类型 | 元素文件名 | 特定字段 |
|---------|-----------|---------|
| **民事案件** | `元素_民事.txt` | dfdsr(对方当事人), badw(办案机关), ssqq(诉讼请求) |
| **刑事案件** | `元素_刑事.txt` | xyr(嫌疑人), zmr(罪名), zcr(侦查机关), jcjg(检察机关) |
| **行政案件** | `元素_行政.txt` | xzjg(行政机关), xzfs(行政方式), sxqq(诉求请求) |
| **国家赔偿** | `元素_国赔.txt` | pcjg(赔偿机关), pclx(赔偿类型), pcje(赔偿金额) |
| **劳动仲裁** | `元素_劳动.txt` | yjdw(用人单位), ldgx(劳动关系), zyqq(仲裁请求) |

### 自动检测
工具会自动根据元素文件名检测案件类型：
```bash
# 自动识别为民事案件
python fill_docx_pro.py "d:\案件" "元素_民事.txt"

# 自动识别为刑事案件
python fill_docx_pro.py "d:\案件" "元素_刑事.txt"
```

### 字段验证
每种案件类型都有必填字段验证：
- **民事**: leib, anyou, weitr, badw, jied
- **刑事**: leib, anyou, weitr, zmr, badw, jied
- **行政**: leib, anyou, weitr, xzjg, xzfs

## 工作流程

1. 准备对应类型的元素文件（如 `元素_民事.txt` 或 `元素_刑事.txt`）
2. 准备文档模板（包含占位符）
3. 运行脚本批量填充
4. 检查生成的文档

## Files

- `fill_docx_pro.py` - 主填充脚本（支持批量）
- `placeholder_detector.py` - 增强占位符检测模块（优化1）
- `case_type_manager.py` - 案件类型管理模块（优化2）
- `SKILL.md` - 本说明文档

## 优化记录

### 优化1: 增强占位符检测
- ✅ 支持多种占位符格式：`{key}`, `{{key}}`, `[key]`, `【key】` 等
- ✅ 支持默认值：`{key?默认值}`
- ✅ 支持格式说明：`{date:date:%Y年%m月%d日}`, `{name:upper}`
- ✅ 智能重叠检测：优先匹配更具体的模式
- ✅ 向后兼容：基础模式仍可正常使用

### 优化2: 多案件类型支持
- ✅ 支持民事、刑事、行政、国家赔偿、劳动仲裁案件
- ✅ 自动检测案件类型
- ✅ 类型特定的字段映射
- ✅ 必填字段验证
- ✅ 向后兼容：未指定类型时默认为民事

### 优化3: 模板配置文件
- ✅ JSON格式定义填充规则
- ✅ 支持字段转换（大小写、日期格式等）
- ✅ 条件表达式支持
- ✅ 模板特定规则覆盖
- ✅ 全局设置统一管理

### 优化4: 交互式填充
- ✅ 缺失必填字段时提示输入
- ✅ 支持多种输入类型（文本、日期、选择、多行）
- ✅ 默认值和验证功能
- ✅ 输入缓存避免重复提示
- ✅ 批量操作确认

### 优化5: 增量更新
- ✅ 只填充空白处，保留已有内容
- ✅ 智能识别空白（占位符、空白字符、占位文本）
- ✅ 支持强制覆盖模式
- ✅ 避免误覆盖已填写内容

### 优化6: 输出组织
- ✅ 按案件类型自动分类输出
- ✅ 结构化目录：`output/民事/`、`output/刑事/`等
- ✅ 支持列出和查询已生成文件
- ✅ 便于批量管理和归档

### 优化7: 日志与回滚
- ✅ 详细记录每次填充操作
- ✅ 自动创建文件备份
- ✅ 支持回滚到任意历史版本
- ✅ 生成填充报告
- ✅ 操作历史查询

## 文件结构

```
法律援助文书/
├── 元素_民事.txt          # 案件数据元素
├── 元素_刑事.txt
├── *.docx                 # 模板文件
├── config.json            # 模板配置（优化3）
├── output/                # 输出目录（优化6）
│   ├── 民事/
│   │   └── 法援结案_已填充.docx
│   └── 刑事/
├── backup/                # 备份目录（优化7）
│   └── 法援结案.docx.20240323_143052.bak
└── logs/                  # 日志目录（优化7）
    ├── fill_202403.jsonl
    └── report_民事_20240323.txt
```

## 注意事项

1. 模板文件中需要包含与元素文件对应的占位符
2. 占位符区分大小写
3. 已生成的文件（`*_已填充.docx`）不会被当作模板处理
4. 临时文件（`~$*.docx`）会被自动排除
5. 使用增强模式时，确保 `placeholder_detector.py` 文件存在
6. 使用案件类型管理时，确保 `case_type_manager.py` 文件存在
7. 使用交互式填充时，确保 `interactive_filler.py` 文件存在
8. 使用文档管理功能时，确保 `document_manager.py` 文件存在

## 新增模块文件

| 文件 | 功能 | 对应优化 |
|------|------|----------|
| `placeholder_detector.py` | 增强占位符检测 | 优化1 |
| `case_type_manager.py` | 多案件类型管理 | 优化2 |
| `template_config.py` | 模板配置管理 | 优化3 |
| `interactive_filler.py` | 交互式填充 | 优化4 |
| `document_manager.py` | 文档管理（增量更新、输出组织、日志回滚） | 优化5、6、7 |
