---
name: legal-document
description: 法律文书AI生成工具,支持关键词匹配和AI案情解析双驱动模式,内置200+官方权威模板覆盖民事商事刑事等9大领域67类场景,提供合规校验法条关联多格式导出功能;当用户需要起诉状答辩状合同协议申请书等法律文书生成时使用
dependency:
  python:
    - python-docx==0.8.11
    - reportlab==3.6.12
    - jieba==0.42.1
    - fuzzywuzzy==0.18.0
    - python-Levenshtein==0.21.1
---

# 法律文书生成工具

## 任务目标

- 本Skill用于：快速生成符合官方规范的法律文书
- 能力包含：双驱动模式（关键词匹配+AI解析）、模板库管理、合规校验、法条关联、多格式导出
- 触发条件：用户需要起草起诉状、答辩状、合同协议、申请书、判决书等法律文书

## 前置准备

### 依赖安装
```bash
pip install python-docx==0.8.11 reportlab==3.6.12 jieba==0.42.1 fuzzywuzzy==0.18.0 python-Levenshtein==0.21.1
```

### 模板文件夹配置（推荐）
用户可以将自己的法律文书Word模板统一放入一个文件夹，系统会优先从这个文件夹查找模板：

```bash
# 创建模板文件夹
mkdir -p ./legal_templates

# 将您的Word模板文件放入该文件夹
# 支持格式：.docx、.doc

# 构建模板索引（首次使用或新增模板后执行）
python scripts/template_finder.py --action index --user_template_dir ./legal_templates
```

**模板文件夹结构建议**：
```
legal_templates/
├── 民事起诉状-借款纠纷.docx
├── 房屋租赁合同.docx
├── 劳动合同模板.docx
├── 离婚协议书.docx
└── ...
```

### 核心资源
- 用户模板文件夹：./legal_templates/（用户自定义模板，优先使用）
- 内置模板库：references/template_library.json（200+份官方权威模板，备选）
- 法条库：references/law_database.json（常用法律法规）
- 合规指南：references/compliance_guide.md
- 写作规范：references/writing_standards.md

## 操作步骤

### 步骤1：接收用户输入

用户可通过三种方式输入：

**方式一：关键词直达**
- 用户提供文书类型关键词（如"民事起诉状"、"房屋租赁合同"）
- 调用 `scripts/template_matcher.py` 进行关键词匹配
- 快速定位目标模板

**方式二：案情描述**
- 用户描述具体案情或法律需求
- 智能体分析案情，提取关键要素
- 调用 `scripts/template_matcher.py` 基于案情推荐最优模板

**方式三：上传Word模板**
- 用户上传现有的Word格式模板文件（.docx/.doc）
- 调用 `scripts/word_template_processor.py` 解析模板
- 自动提取占位符和要素信息
- 支持直接基于用户模板生成文书

### 步骤2：AI解析与智能模板查找

**查找优先级（自动执行）**：
1. **优先**：在用户模板文件夹中查找匹配模板
2. **备选**：在内置模板库中查找匹配模板
3. **兜底**：AI根据需求自行生成文书

**智能体职责**：
- 分析用户输入，识别文书类型、法律关系、关键要素
- 提取案情的核心信息：当事人信息、争议焦点、法律依据等
- 判断文书适用场景和领域

**脚本调用（智能查找）**：
```python
# 智能查找模板（优先用户模板，其次内置模板库）
python scripts/template_finder.py --action find \
  --user_template_dir ./legal_templates \
  --keywords "民事 起诉状 借款" \
  --case_description "朋友借款不还，我要起诉"

# 构建用户模板索引（首次使用或新增模板后执行）
python scripts/template_finder.py --action index \
  --user_template_dir ./legal_templates

# 列出所有用户模板
python scripts/template_finder.py --action list \
  --user_template_dir ./legal_templates
```

**查找流程示例**：
```
用户输入："我要起诉借款纠纷"

步骤1：在用户模板文件夹中查找...
  ✓ 找到匹配模板：民事起诉状-借款纠纷.docx
  匹配度：85%
  使用用户模板 → 解析占位符 → 填充生成

---

如果用户模板文件夹中没有匹配项：

步骤1：在用户模板文件夹中查找...
  未找到匹配模板
  
步骤2：在内置模板库中查找...
  ✓ 找到匹配模板：民事起诉状
  匹配度：90%
  使用内置模板 → 引导填写要素 → 生成文书

---

如果都没有匹配项：

步骤2：在内置模板库中查找...
  未找到匹配模板
  
建议：使用AI根据案情描述自行生成文书
```

### 步骤3：引导补充要素信息

**智能体职责**：
- 根据模板要求，向用户提问缺失的要素
- 采用结构化方式收集信息（当事人信息、诉讼请求、事实理由等）
- 对专业法律术语进行通俗解释，适配普通用户

**情况A：模板有占位符**
- 直接识别占位符列表
- 引导用户逐一填写占位符内容
- 示例：`{{原告姓名}}` → 提示"请输入原告姓名"

**情况B：模板无占位符（参照格式生成）**
- 调用 `scripts/template_analyzer.py` 深度分析模板
- 提取模板的结构、格式、风格特征
- 生成模板分析报告（包含章节结构、格式要求、语言风格）
- 智能体根据模板分析报告，学习模板的格式和风格
- 引导用户提供必要信息，按模板格式生成文书

**要素收集示例**：
- 当事人信息：原告/被告姓名、性别、年龄、住址、联系方式
- 诉讼请求：具体的诉求内容
- 事实与理由：案情经过、证据情况
- 法律依据：引用的相关法条

**适配策略**：
- 专业用户：直接提供要素清单，快速填写
- 普通用户：分步引导，每步解释说明

### 步骤4：合规校验与法条关联

**智能体职责**：
- 依据 `references/compliance_guide.md` 进行合规校验
- 检查文书要素完整性、格式规范性、逻辑一致性
- 识别潜在的法律风险和漏洞

**合规校验清单**：
- ✅ 必备要素是否齐全
- ✅ 格式是否符合《人民法院诉讼文书样式》
- ✅ 诉讼请求是否明确具体
- ✅ 事实陈述是否清晰有据
- ✅ 法律依据是否准确适用

**脚本调用**：
```python
# 检索相关法条
python scripts/law_retriever.py --document_type "民事起诉状" --keywords "借款 民间借贷" --output related_laws.json
```

**法条关联**：
- 根据文书类型和案情自动匹配相关法条
- 提供法条编号、内容和适用说明
- 支持一键插入到文书中

### 步骤5：生成文书与多格式导出

**智能体职责**：
- 根据模板和收集的信息生成完整文书
- 确保语言规范、逻辑严谨、格式正确
- 应用 `references/writing_standards.md` 中的写作规范

**脚本调用**：
```python
# 生成Word格式
python scripts/doc_generator.py --content "文书内容" --format word --output ./output/document.docx

# 生成PDF格式
python scripts/doc_generator.py --content "文书内容" --format pdf --output ./output/document.pdf
```

**输出规范**：
- 符合最高法、司法部官方文书样式
- 可直接用于诉讼、公证、备案等场景
- 保留可编辑版本供用户调整

**批量生成**：
- 支持基于模板批量生成同类文书
- 适用于律所、法务部门等批量需求

## 资源索引

### 必要脚本
- [scripts/template_finder.py](scripts/template_finder.py) - 智能模板查找器（优先用户模板文件夹，其次内置模板库）
- [scripts/template_analyzer.py](scripts/template_analyzer.py) - 模板分析器（深度分析模板结构、格式、风格，即使无占位符也能学习）
- [scripts/word_template_processor.py](scripts/word_template_processor.py) - Word模板处理器（读取、解析、填充用户Word模板）
- [scripts/template_matcher.py](scripts/template_matcher.py) - 内置模板匹配引擎
- [scripts/law_retriever.py](scripts/law_retriever.py) - 法条检索工具
- [scripts/doc_generator.py](scripts/doc_generator.py) - 文书生成器（Word/PDF）
- [scripts/template_manager.py](scripts/template_manager.py) - 模板库管理工具

### 领域参考
- [references/template_library.json](references/template_library.json) - 200+份文书模板库（何时读取：模板匹配时）
- [references/law_database.json](references/law_database.json) - 常用法条数据库（何时读取：法条关联时）
- [references/compliance_guide.md](references/compliance_guide.md) - 合规校验指南（何时读取：文书校验时）
- [references/writing_standards.md](references/writing_standards.md) - 写作规范（何时读取：文书生成时）

### 输出资产
- 文书示例：见 assets/examples/ 目录

### 格式规范
- [references/FORMAT_STANDARDS.md](references/FORMAT_STANDARDS.md) - 格式规范总览
- [references/刑事法援会见前打印格式规范.md](references/刑事法援会见前打印格式规范.md) - 刑事会见类
- [references/调证材料格式规范.md](references/调证材料格式规范.md) - 调查取证类
- [references/证据目录格式规范.md](references/证据目录格式规范.md) - 证据目录类
- [references/法律援助案卷归档目录格式规范.md](references/法律援助案卷归档目录格式规范.md) - 案卷归档类

## 空白模板生成

系统支持一键生成各类法律文书的空白模板，占位符统一替换为8个空格以保证打印宽度。

### 统一生成命令
```bash
# 生成所有空白模板
python scripts/create_all_blank_templates.py all

# 按类别生成
python scripts/create_all_blank_templates.py criminal  # 刑事类
python scripts/create_all_blank_templates.py civil     # 民事类
```

### 单独生成命令
| 文书类型 | 脚本 | 模板文件 |
|----------|------|----------|
| 刑事法援会见前打印 | `create_blank_criminal_legal_aid_meeting.py` | `8 刑事法援会见前打印.docx` |
| 调证材料 | `create_blank_investigation_doc.py` | `2 9 调证材料.docx` |
| 证据目录 | `create_blank_evidence_directory.py` | `3 证据目录 .docx` |
| 民事法援案卷归档目录 | `create_blank_civil_legal_aid_directory.py` | `9 湖南省法律援助案卷归档目录（民事）.docx` |
| 刑事法援案卷归档目录 | `create_blank_criminal_legal_aid_directory.py` | `8 湖南省法律援助案卷归档目录（刑事）.docx` |

### 占位符替换规则
| 场景 | 替换内容 |
|------|----------|
| 正式生成文书 | 占位符替换为实际内容 |
| 生成空白模板 | 占位符替换为8个空格（`        `） |

## 交互式生成器

提供向导式问答生成文书，适合不熟悉命令行的用户。

```bash
# 启动交互式生成器
python scripts/interactive_generator.py

# 直接指定文书类型
python scripts/interactive_generator.py criminal_meeting
```

支持的文书类型：criminal_meeting, investigation, evidence, civil_directory, criminal_directory

## 批量生成器

支持 JSON/Excel 格式批量导入数据，一次性生成多份文书。

```bash
# JSON 格式批量生成
python scripts/batch_generator.py criminal_meeting data.json

# Excel 格式批量生成
python scripts/batch_generator.py investigation data.xlsx

# 指定输出文件名前缀
python scripts/batch_generator.py criminal_meeting data.json -o 我的文书
```

数据文件格式示例 (JSON)：
```json
[
  {"weitr": "张三", "wtrsfz": "430124198801011234", "anyou": "盗窃罪", "jied": "侦查"},
  {"weitr": "李四", "wtrsfz": "430124199001011234", "anyou": "诈骗罪", "jied": "审查起诉"}
]
```

## 基础模块

`document_base.py` 提供统一的基础功能：

| 函数 | 功能 |
|------|------|
| `create_simple_document()` | 通用文档生成 |
| `validate_placeholders_replaced()` | 校验占位符替换 |
| `batch_generate()` | 批量生成 |
| `get_generation_history()` | 获取生成历史 |

配置文件 `config.json`（位于 scripts 目录）：
```json
{
  "output_dir": "output",
  "blank_placeholder": "        ",
  "log_enabled": true
}
```

## 注意事项

### 权威性保障
- 所有模板基于最高法、司法部发布的官方样式
- 定期更新模板库，确保符合最新法律规范
- 文书格式严格遵循《人民法院诉讼文书样式》

### 合规性要求
- 必须完成合规校验后才能输出最终文书
- 对于重大法律事项，提示用户咨询专业律师
- 保留文书的生成记录和版本历史

### 用户体验优化
- 普通用户：简化流程，提供引导式问答
- 专业用户：提供快速模式和高级选项
- 支持文书保存、修改、再次生成

### 技术限制
- 用户模板文件夹需要先构建索引才能高效检索
- 模板匹配依赖于关键词准确性和文件命名规范
- 法条检索基于本地数据库，最新法律变更需手动更新
- 批量生成功能需要会员权限

### 最佳实践
- **文件命名**：建议使用清晰的文件名，如"民事起诉状-借款纠纷.docx"
- **模板更新**：新增模板后及时更新索引：`python scripts/template_finder.py --action index --user_template_dir ./legal_templates`
- **占位符规范**：在Word模板中使用`{{占位符名称}}`格式标记需要填充的位置
- **无占位符模板**：即使没有占位符，系统也能通过AI分析学习模板格式，参照生成文书
- **格式规范**：确保用户模板格式规范、结构清晰，有助于AI更好地学习模板特征

## 使用示例

### 示例1：使用用户模板文件夹中的模板（推荐）

**前置准备**：
```bash
# 1. 创建模板文件夹
mkdir -p ./legal_templates

# 2. 将您的Word模板放入文件夹
# 例如：民事起诉状-借款纠纷.docx、房屋租赁合同.docx 等

# 3. 构建模板索引（首次使用或新增模板后执行）
python scripts/template_finder.py --action index --user_template_dir ./legal_templates
```

**用户输入**：
```
关键词：借款纠纷、起诉状
```

**执行流程**：
1. 调用 template_finder.py 在用户模板文件夹中查找
2. ✓ 找到匹配模板：民事起诉状-借款纠纷.docx（匹配度85%）
3. 解析模板占位符：{{原告姓名}}、{{被告姓名}}、{{借款金额}}等
4. 引导用户填写所有占位符内容
5. 合规校验 + 法条关联
6. 生成民事起诉状（Word/PDF格式）

**命令示例**：
```bash
# 查找模板
python scripts/template_finder.py --action find \
  --user_template_dir ./legal_templates \
  --keywords "借款 起诉状"

# 使用找到的模板生成文书
python scripts/word_template_processor.py --action fill \
  --template ./legal_templates/民事起诉状-借款纠纷.docx \
  --data '{"原告姓名": "张三", "被告姓名": "李四", "借款金额": "10万元"}' \
  --output ./output/民事起诉状.docx
```

### 示例2：使用内置模板库（用户模板文件夹无匹配项）

**用户输入**：
```
案情描述：我要出租自己的房子，租期一年，月租金3000元，押一付三
```

**执行流程**：
1. 在用户模板文件夹中查找 → 未找到匹配项
2. 在内置模板库中查找 → ✓ 找到"房屋租赁合同"模板
3. 引导补充：房屋地址、面积、设施、维修责任、违约条款等
4. 合规校验 + 法条关联
5. 生成房屋租赁合同

### 示例3：AI自行生成（无匹配模板）

**用户输入**：
```
案情描述：我需要一份特殊的调解协议书，涉及三方当事人的债务重组
```

**执行流程**：
1. 在用户模板文件夹和内置模板库中查找 → 都未找到匹配项
2. AI根据案情描述自行分析需求
3. 提取关键要素：三方当事人信息、债务详情、重组方案等
4. 参考法律规范自行生成调解协议书
5. 合规校验
6. 输出文书

### 示例4：批量生成（专业用户）

**用户需求**：
```
为10位员工生成劳动合同
```

**前提**：用户模板文件夹中已有"劳动合同模板.docx"

**执行流程**：
1. 在用户模板文件夹中找到劳动合同模板
2. 准备员工信息列表（JSON格式）
3. 批量填充模板
4. 生成10份劳动合同文档

**批量生成命令**：
```bash
python scripts/word_template_processor.py --action batch \
  --template ./legal_templates/劳动合同模板.docx \
  --data_list employees.json \
  --output_dir ./contracts/
```

### 示例5：使用无占位符的模板（参照格式生成）

**场景**：用户提供了一份格式规范的民事起诉状模板，但没有占位符标记

**执行流程**：
1. 在用户模板文件夹中找到模板：民事起诉状-标准格式.docx
2. 调用模板分析器深度分析模板：
   ```bash
   python scripts/template_analyzer.py \
     --template ./legal_templates/民事起诉状-标准格式.docx \
     --output template_analysis.json \
     --brief
   ```
3. 分析报告输出：
   ```
   【文档类型】: 诉讼文书-起诉状
   
   【主要章节】:
     - 当事人信息
     - 诉讼请求
     - 事实与理由
     - 证据清单
   
   【关键要素】:
     - 当事人信息
     - 请求/条款
     - 事实理由
     - 证据材料
     - 落款签名
   
   【格式要求】:
     - 标题居中对齐
     - 标题字号：18磅
     - 标题加粗
   
   【语言要求】:
     - 使用法律术语：原告, 被告, 诉讼请求, 事实与理由
     - 正式程度：high
   ```
4. 智能体学习模板格式和风格
5. 引导用户提供具体内容：
   - "请提供原告信息（姓名、性别、年龄、住址、联系方式）"
   - "请提供被告信息"
   - "请描述您的诉讼请求"
   - "请陈述事实与理由"
6. 按照模板的格式和结构生成新文书

**技术说明**：
- 即使模板没有占位符，系统也能通过AI分析学习模板的：
  - ✓ 章节结构（当事人信息→诉讼请求→事实与理由→落款）
  - ✓ 格式样式（标题居中、字号18磅、章节加粗等）
  - ✓ 语言风格（正式程度、法律术语使用习惯）
  - ✓ 内容模式（典型要素、表达方式）
- 智能体根据学到的格式规范生成符合要求的新文书

### 示例6：管理用户模板文件夹

**查看所有用户模板**：
```bash
python scripts/template_finder.py --action list \
  --user_template_dir ./legal_templates
```

**输出示例**：
```
用户模板列表 (共 5 个):
================================================================================
1. 民事起诉状-借款纠纷.docx
   标题: 民事起诉状
   类型: 诉讼文书
   关键词: 民事, 起诉状, 借款, 纠纷
   占位符数量: 8

2. 房屋租赁合同.docx
   标题: 房屋租赁合同
   类型: 合同协议
   关键词: 房屋, 租赁, 合同
   占位符数量: 15

...
```

**新增模板后更新索引**：
```bash
# 将新模板放入 legal_templates/ 文件夹后
python scripts/template_finder.py --action index \
  --user_template_dir ./legal_templates
```

## 模板库更新

调用 `scripts/template_manager.py` 管理模板：

```python
# 查看所有模板
python scripts/template_manager.py --action list

# 添加新模板
python scripts/template_manager.py --action add --template_file new_template.json

# 更新模板
python scripts/template_manager.py --action update --template_id "civil-complaint-001" --template_file updated_template.json

# 删除模板
python scripts/template_manager.py --action delete --template_id "old-template"
```
