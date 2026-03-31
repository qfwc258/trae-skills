# Legal PPT Template - 法律案件分析PPT模板

## 概述

基于湖南金厚（宁乡）律师事务所风格的法律案件分析PPT模板，使用pptxgenjs构建。

## 设计风格

- **配色方案**: 深蓝色专业主题 (#1E3A5F, #2C5282) + 金色点缀 (#C9A227)
- **字体**: Microsoft YaHei (微软雅黑)
- **布局**: 16:9 宽屏格式
- **特点**:
  - 律所品牌标识展示
  - 结构化案件分析框架
  - 核心策略/抗辩分析/证据清单等专业模板

## 模板页面（共12页）

| 页码 | 页面类型 | 说明 |
|------|----------|------|
| 1 | 封面 | 标题页，支持案号显示 |
| 2 | 案件概述 | 当事人信息、案件基本信息、背景摘要 |
| 3 | 案件时间线 | 时间轴展示案件发展脉络 |
| 4 | SWOT分析 | 四象限展示双方优劣势机会威胁 |
| 5 | 法律要件分解 | 逐条分析法律构成要件及满足状态 |
| 6 | 金额计算 | 诉请金额明细与合计 |
| 7 | 诉讼进度 | 阶段性进度指示器 |
| 8 | 核心策略 | 策略分析、主要诉求列表 |
| 9 | 对方抗辩分析 | 抗辩要点、我方反驳、关键结论 |
| 10 | 关键证据清单 | 三栏分类展示证据 |
| 11 | 法律依据 | 法条引用高亮框 |
| 12 | 结束页 | 谢谢/联系方式 |

## 可配置项

```javascript
const config = {
  title: "合同纠纷案件分析",
  lawFirm: {
    name: "湖南金厚（宁乡）律师事务所",
    tel: "139 7589 2485",
    branch: "金厚"
  },
  case: {
    caseNo: "(2024)湘0124民初1234号",
    court: "宁乡市人民法院",
    type: "民事诉讼"
  },
  colors: {
    primary: "1E3A5F",
    secondary: "2C5282",
    accent: "C9A227"
  }
};
```

## 使用方法

### 运行模板生成器

```bash
node create_template.js
```

### 自定义单个页面

```javascript
const { createSwotSlide, createTimelineSlide, defaultConfig } = require('./create_template');

let pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";

// 自定义SWOT数据
const swotData = {
  strengths: ["证据充分", "合同明确"],
  weaknesses: ["标的较大"],
  opportunities: ["对方经营困难"],
  threats: ["执行难度"]
};
createSwotSlide(pres, config, swotData);

// 自定义时间线
const events = [
  { date: "2024.01", title: "合同签订", desc: "双方签订协议" },
  { date: "2024.03", title: "发货", desc: "产品交付" }
];
createTimelineSlide(pres, config, events);

pres.writeFile({ fileName: "custom.pptx" });
```

## 导出函数列表

| 函数 | 说明 |
|------|------|
| `createFullTemplate(config)` | 生成完整12页模板 |
| `createLegalSlide(pres, config)` | 封面 |
| `createCaseOverviewSlide(pres, config, data)` | 案件概述 |
| `createTimelineSlide(pres, config, events)` | 时间线 |
| `createSwotSlide(pres, config, swotData)` | SWOT分析 |
| `createLegalElementSlide(pres, config, elements)` | 法律要件分解 |
| `createAmountSlide(pres, config, amounts)` | 金额计算 |
| `createProgressSlide(pres, config, stages)` | 诉讼进度 |
| `createStrategySlide(pres, config, data)` | 核心策略 |
| `createDefenseSlide(pres, config, data)` | 抗辩分析 |
| `createEvidenceListSlide(pres, config, categories)` | 证据清单 |
| `createLawCitationSlide(pres, config, citations)` | 法律依据 |
| `createEndingSlide(pres, config)` | 结束页 |

## 文件结构

```
legal-ppt-template/
├── create_template.js           # 主模板生成脚本（含12个页面函数）
├── create_template_backup.js   # 原始简单版备份
├── 法律案件分析模板_enhanced.pptx  # 增强版示例
├── _meta.json                  # 元数据
└── README.md                   # 使用说明
```

## 示例输出

运行后将生成 `法律案件分析模板_enhanced.pptx` 文件，包含12页专业法律分析模板。
