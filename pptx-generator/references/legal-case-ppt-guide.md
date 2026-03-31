# 法律案件分析PPT制作指南

## 概述

本指南基于唐浩然案件PPT制作经验，总结法律案件分析类PPT的特定要求和最佳实践。

## 适用场景

- 机动车交通事故责任纠纷
- 民事案件分析报告
- 法援案件汇报
- 诉讼策略展示

---

## 配色方案：现代商务蓝系

### 主色调

```javascript
const theme = {
  primary: "1a365d",    // 深蓝 - 主色（标题、背景）
  secondary: "2c5282",  // 中蓝 - 辅助（副标题、强调）
  accent: "3182ce",     // 亮蓝 - 强调（装饰线、高亮）
  light: "90cdf4",      // 浅蓝 - 装饰（副标题文字）
  bg: "ebf8ff"          // 极浅蓝 - 背景（卡片背景）
};
```

### 使用规范

| 元素 | 颜色 | 用途 |
|------|------|------|
| 封面背景 | `theme.primary` | 深蓝纯色背景 |
| 封面标题 | `"FFFFFF"` | 白色 |
| 封面副标题 | `theme.light` | 浅蓝 |
| 内容页背景 | `"FFFFFF"` | 白色 |
| 页面标题 | `theme.primary` | 深蓝 |
| 卡片背景 | `theme.bg` | 极浅蓝 |
| 装饰线条 | `theme.accent` | 亮蓝 |
| 正文文字 | `theme.secondary` | 中蓝 |

---

## 字体规范

### 中文字体

| 元素 | 字体 | 字号 | 粗细 |
|------|------|------|------|
| 封面主标题 | Microsoft YaHei | 44pt | Bold |
| 封面副标题 | Microsoft YaHei | 24pt | Normal |
| 页面标题 | Microsoft YaHei | 28pt | Bold |
| 卡片标题 | Microsoft YaHei | 16pt | Bold |
| 卡片内容 | Microsoft YaHei | 13pt | Normal |
| 正文文字 | Microsoft YaHei | 14pt | Normal |
| 表格文字 | Microsoft YaHei | 12pt | Normal |

---

## 页面结构

### 1. 封面页 (Cover)

```javascript
// 布局要点
- 背景：theme.primary 纯色
- 顶部装饰条：theme.accent, 高度 0.15"
- 左侧装饰线：theme.accent, 宽度 0.08"
- 主标题：白色，44pt，左对齐
- 副标题：theme.light，24pt
- 案件信息：白色，18pt
```

### 2. 目录页 (TOC)

```javascript
// 布局要点
- 背景：白色
- 顶部装饰条：theme.primary, 高度 0.12"
- 页面标题："目录"，theme.primary，28pt
- 标题下划线：theme.accent，宽度 1.5"
- 目录项：带圆形编号，左对齐
- 编号圆圈：theme.accent 背景，白色文字
```

### 3. 内容页 (Content)

```javascript
// 布局要点
- 背景：白色
- 顶部装饰条：theme.primary, 高度 0.12"
- 页面标题格式："01  案件基本信息"
- 标题下划线：theme.accent，高度 0.05"
- 信息卡片：
  - 背景：theme.bg
  - 圆角：rectRadius: 0.1
  - 内边距：0.2"
  - 阴影效果（可选）
```

### 4. 信息卡片设计

```javascript
// 卡片样式
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.3, w: 4.3, h: 1.6,
  fill: { color: theme.bg },
  rectRadius: 0.1
});

// 卡片标题
slide.addText("原告信息", {
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: theme.primary, bold: true
});

// 卡片内容（多行）
slide.addText("姓名：唐浩然\n身份：未成年人\n诉讼地位：原告", {
  fontSize: 13, fontFace: "Microsoft YaHei",
  color: theme.secondary,
  lineSpacing: 22  // 行间距
});
```

### 5. 表格页

```javascript
// 赔偿项目表格示例
const tableData = [
  [{ text: "赔偿项目", options: { fill: theme.primary, color: "FFFFFF", bold: true } },
   { text: "金额（元）", options: { fill: theme.primary, color: "FFFFFF", bold: true } }],
  ["医疗费", "7,363.63"],
  ["后续治疗费", "8,000.00"],
  // ...
];

slide.addTable(tableData, {
  x: 0.5, y: 1.5, w: 9,
  colW: [6, 3],
  border: { type: "solid", pt: 0.5, color: theme.light },
  fill: { color: "FFFFFF" },
  fontFace: "Microsoft YaHei",
  fontSize: 12
});
```

### 6. 结束页

```javascript
// 布局要点
- 背景：theme.primary 纯色
- 居中显示"谢谢"或"Thanks"
- 白色文字，48pt
- 底部可添加律所信息
```

---

## 组件规范

### 圆形编号

```javascript
// 用于目录、步骤等
slide.addShape(pres.shapes.OVAL, {
  x: 0.8, y: 1.8, w: 0.35, h: 0.35,
  fill: { color: theme.accent }
});
slide.addText("1", {
  x: 0.8, y: 1.8, w: 0.35, h: 0.35,
  fontSize: 14, color: "FFFFFF", bold: true,
  align: "center", valign: "middle"
});
```

### 责任标签

```javascript
// 主责/次责/无责标签
const responsibilityTag = {
  "主要责任": { bg: "DC2626", color: "FFFFFF" },  // 红色
  "次要责任": { bg: "F59E0B", color: "FFFFFF" },  // 橙色
  "无责任": { bg: "10B981", color: "FFFFFF" }     // 绿色
};

slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.5, y: 2.2, w: 0.8, h: 0.3,
  fill: { color: "DC2626" },
  rectRadius: 0.05
});
slide.addText("主责", {
  x: 3.5, y: 2.2, w: 0.8, h: 0.3,
  fontSize: 11, color: "FFFFFF",
  align: "center", valign: "middle"
});
```

---

## 页面类型映射

| 页面 | 类型 | 内容 |
|------|------|------|
| 封面 | cover | 案件法律分析报告、案件类型、原告信息 |
| 目录 | toc | 五大板块导航 |
| 案件基本信息 | content | 原告、被告信息卡片 |
| 事故经过 | content | 时间线、责任认定 |
| 赔偿项目 | content | 表格形式展示 |
| 证据清单 | content | 分组展示证据 |
| 法律分析 | content | 依据、策略、结论 |
| 结束页 | final | 谢谢 |

---

## 最佳实践

### 1. 信息密度控制

- 每页不超过4个信息卡片
- 表格不超过8行（超出则分页）
- 使用卡片分组，避免信息堆砌

### 2. 视觉层次

- 使用颜色区分重要性
- 深蓝 > 中蓝 > 浅蓝
- 重要数据使用亮色强调

### 3. 对齐规范

- 所有元素左对齐或居中对齐
- 卡片间距保持一致（0.3"-0.4"）
- 页面边距：左右 0.5"，上下 0.4"

### 4. 数据展示

- 金额数字右对齐
- 使用千分位分隔符
- 保留两位小数

---

## 完整示例

```javascript
// compile.js - 法律案件分析PPT编译脚本
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// 现代商务蓝系主题
const theme = {
  primary: "1a365d",
  secondary: "2c5282",
  accent: "3182ce",
  light: "90cdf4",
  bg: "ebf8ff"
};

// 加载所有幻灯片
const slides = [
  'slide-01-cover',
  'slide-02-toc',
  'slide-03-case-info',
  'slide-04-accident',
  'slide-05-compensation',
  'slide-06-evidence',
  'slide-07-analysis',
  'slide-08-final'
];

slides.forEach((name, index) => {
  const slideModule = require(`./${name}`);
  slideModule.createSlide(pres, theme);
});

pres.writeFile({ fileName: '法律案件分析报告.pptx' });
```

---

## 注意事项

1. **隐私保护**：确保案件敏感信息已脱敏处理
2. **数据准确性**：赔偿金额、日期等需与材料一致
3. **格式统一**：同一类型元素保持样式一致
4. **打印友好**：避免使用深色背景的大面积区域
