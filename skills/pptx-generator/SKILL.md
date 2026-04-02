---
name: pptx-generator
description: Generate, edit, and read PowerPoint presentations. Create from scratch with PptxGenJS, edit existing PPTX via XML workflows, or extract text with markitdown. Includes built-in legal case analysis templates (merged from legal-ppt-template).
version: 1.1
---

# PPTX Generator & Editor

## Overview

This skill handles all PowerPoint tasks: reading/analyzing existing presentations, editing template-based decks via XML manipulation, and creating presentations from scratch using PptxGenJS. It includes a complete design system (color palettes, fonts, style recipes), detailed guidance for every slide type, and **built-in legal case analysis templates** (merged from legal-ppt-template).

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | See `references/editing.md` |
| Create from scratch | See Creating from Scratch workflow below |
| Legal case analysis | Use Legal Theme + Legal Slide Patterns |

| Item | Value |
|------|-------|
| **Dimensions** | 10" x 5.625" (LAYOUT_16x9) |
| **Chinese font** | Microsoft YaHei |
| **Theme keys** | `primary`, `secondary`, `accent`, `light`, `bg` |

## Legal Theme Colors

```javascript
const legalTheme = {
  primary: "1E3A5F",    // Deep navy blue
  secondary: "2C5282",  // Mid blue
  accent: "C9A227",      // Gold accent
  light: "E8D48A",       // Light gold
  bg: "FFFFFF"           // White background
};
```

## Legal Slide Patterns

### Case Overview (交替背景表格)
```javascript
caseInfo.forEach((info, i) => {
  const y = 1.2 + i * 0.65;
  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: y, w: 9, h: 0.55,
    fill: { color: i % 2 === 0 ? "#F9FAFB" : "#FFFFFF" },
    line: { color: "#E5E7EB", width: 1 }
  });
  slide.addText(info.label, { x: 0.7, y: y + 0.12, w: 1.8, h: 0.3, fontSize: 14, bold: true });
  slide.addText(info.value, { x: 2.6, y: y + 0.12, w: 6.7, h: 0.3, fontSize: 14 });
});
```

### Claims with numbered cards
```javascript
claims.forEach((claim, i) => {
  const y = 2.0 + i * 1.0;
  slide.addShape(pres.ShapeType.rect, { x: 0.5, y: y, w: 9, h: 0.85, fill: { color: "#F9FAFB" } });
  slide.addShape(pres.ShapeType.rect, { x: 0.5, y: y, w: 0.6, h: 0.85, fill: { color: "#1E3A5F" } });
  slide.addText(claim.num, { x: 0.5, y: y + 0.2, w: 0.6, h: 0.4, fontSize: 18, align: "center" });
  slide.addText(claim.text, { x: 1.3, y: y + 0.22, w: 8, h: 0.4, fontSize: 16 });
});
```

### Defense Points (左侧色条)
```javascript
defensePoints.forEach((point, i) => {
  const y = 1.15 + i * 1.35;
  slide.addShape(pres.ShapeType.rect, { x: 0.5, y: y, w: 9, h: 1.2 });
  slide.addShape(pres.ShapeType.rect, { x: 0.5, y: y, w: 0.1, h: 1.2, fill: { color: point.color } });
  slide.addText(point.title, { x: 0.8, y: y + 0.1, w: 8.5, h: 0.4, fontSize: 15, bold: true });
  slide.addText(point.content, { x: 0.8, y: y + 0.5, w: 8.5, h: 0.3, fontSize: 13 });
});
```

### Key Issues (三栏卡片)
```javascript
issues.forEach((issue, i) => {
  const x = 0.5 + i * 3.1;
  slide.addShape(pres.ShapeType.rect, { x: x, y: 1.5, w: 2.9, h: 3.5 });
  slide.addShape(pres.ShapeType.rect, { x: x, y: 1.5, w: 2.9, h: 0.5, fill: { color: "#1E3A5F" } });
  slide.addText(issue.focus, { x: x, y: 1.55, w: 2.9, h: 0.4, align: "center" });
});
```

## Legal Case Analysis Template (12 Pages)

| Page | Type | Content |
|------|------|---------|
| 1 | Cover | 案件标题、案号、法院 |
| 2 | Case Overview | 当事人信息、案件背景 |
| 3 | Timeline | 案件时间轴 |
| 4 | SWOT | 双方优劣势分析 |
| 5 | Legal Elements | 法律要件分解 |
| 6 | Amount Calculation | 金额计算明细 |
| 7 | Progress | 诉讼进度 |
| 8 | Core Strategy | 核心策略 |
| 9 | Defense Analysis | 对方抗辩分析 |
| 10 | Evidence List | 关键证据清单 |
| 11 | Legal Basis | 法律依据 |
| 12 | Closing | 谢谢页+联系方式 |

## Dependencies

- `pip install "markitdown[pptx]"` — text extraction
- `npm install -g pptxgenjs` — creating from scratch