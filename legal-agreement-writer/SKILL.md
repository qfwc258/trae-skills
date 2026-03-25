---
name: "legal-agreement-writer"
description: "Creates formal legal agreement documents in Chinese following standard legal document format. Invoke when user needs to create legal agreements, contracts, or settlement documents in Word format."
---

# Legal Agreement Writer (法律文书助手)

This skill creates formal Chinese legal agreement documents following the standard format used in Chinese legal documents.

## 使用场景

当用户需要创建以下类型的法律文书时，应使用此技能：
- 执行和解协议
- 合同协议
- 和解协议
- 调解协议
- 其他正式法律文件

## 文档格式规范

### 1. 页面设置
- 纸张大小：A4 (210mm x 297mm)
- 页边距：上 2.54cm，下 2.54cm，左 3.18cm，右 3.18cm
- 字体：仿宋或宋体
- 字号：
  - 标题：二号加粗（44 磅）
  - 正文：三号或四号（32 磅或 28 磅）

### 2. 文档结构

#### 标题部分
- 标题居中显示
- 使用二号加粗字体
- 示例：执行和解协议

#### 当事人信息
- 列明各方当事人基本信息
- 包括：姓名、性别、出生日期、民族、住址
- 格式示例：
  ```
  申请执行人：范玉泉，女，1969 年 7 月 14 日出生，汉族，住湖南省宁乡市朱良桥乡左家山村刘兴冲组 12 号。
  被执行人：程利英，女，1969 年 1 月 14 日出生，汉族，住湖南省宁乡市双江口镇山园村花园二组 23 号。
  ```

#### 案由说明
- 简述案件背景和来源
- 包括：案件性质、相关法院、案号等
- 格式示例：
  ```
  申请执行人 XXX 与被执行人 XXX 民间借贷纠纷一案，业经 XX 人民法院（XXXX）XX 民终 XXXX 号民事判决书终审判决，并已由 XX 人民法院以（XXXX）XX 执 XXXX 号立案执行。
  ```

#### 协议正文
- 使用数字编号条款（一、二、三、四...）
- 每个条款下可设子项
- 条款标题加粗
- 常见条款包括：
  1. 债务了结方案
  2. 支付安排
  3. 执行措施的解除与修复
  4. 违约责任与特别约定
  5. 其他

#### 签署部分
- 双方当事人签名区域
- 日期栏（年、月、日）
- 份数说明（一式 X 份）

### 3. 数字和大写金额规范

使用中文大写数字：
- 金额：壹、贰、叁、肆、伍、陆、柒、捌、玖、拾
- 单位：佰、仟、万、元、角、分
- 示例：人民币壹拾万元整（¥100,000.00）

### 4. 编号格式

使用中文数字编号：
- 一级标题：一、二、三、四、
- 二级标题：（一）、（二）、（三）
- 三级标题：1、2、3、
- 四级标题：（1）、（2）、（3）

## 创建步骤

使用 docx-js 创建法律文书：

```javascript
const { Document, Packer, Paragraph, TextRun, HeadingLevel, 
        AlignmentType, BorderStyle, WidthType } = require('docx');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "仿宋", size: 32 }, // 三号字
        paragraph: { spacing: { line: 360, after: 0 } }
      }
    },
    paragraphStyles: [
      {
        id: "Title",
        name: "标题",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { 
          size: 44, // 二号字
          bold: true,
          font: "仿宋"
        },
        paragraph: { 
          spacing: { before: 240, after: 240 },
          alignment: AlignmentType.CENTER
        }
      },
      {
        id: "Clause",
        name: "条款",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { 
          size: 32,
          bold: true,
          font: "仿宋"
        },
        paragraph: { 
          spacing: { before: 200, after: 100 }
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { 
          top: 1440, 
          right: 1800, 
          bottom: 1440, 
          left: 1800 
        }
      }
    },
    children: [
      // 标题
      new Paragraph({ 
        heading: HeadingLevel.TITLE,
        children: [new TextRun("执行和解协议")] 
      }),
      
      // 当事人信息
      new Paragraph({ 
        children: [new TextRun("申请执行人：XXX，女，XXXX 年 X 月 X 日出生，汉族，住 XXX。")] 
      }),
      new Paragraph({ 
        children: [new TextRun("被执行人：XXX，女，XXXX 年 X 月 X 日出生，汉族，住 XXX。")] 
      }),
      
      // 案由
      new Paragraph({ 
        children: [new TextRun("申请执行人 XXX 与被执行人 XXX 纠纷一案，业经 XX 人民法院...")] 
      }),
      
      // 协议条款
      new Paragraph({ 
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("一、债务了结方案")] 
      }),
      new Paragraph({ 
        children: [new TextRun("双方确认，为一次性了结...")] 
      }),
      
      // 签署部分
      new Paragraph({ children: [new TextRun("申请执行人：")] }),
      new Paragraph({ children: [new TextRun("被执行人：")] }),
      new Paragraph({ children: [new TextRun("年    月    日")] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("执行和解协议.docx", buffer);
});
```

## 关键要点

1. **字体选择**：使用仿宋或宋体，这是中国法律文书的标准字体
2. **标题格式**：标题居中、加粗、二号字
3. **条款编号**：使用中文数字编号（一、二、三...）
4. **金额大写**：所有金额必须使用中文大写数字
5. **段落间距**：适当设置段落间距，保持文档清晰
6. **签署区域**：预留足够的签名和日期空间

## 常见法律文书类型

- 执行和解协议
- 民事起诉状
- 答辩状
- 上诉状
- 合同协议
- 授权委托书
- 律师函

## 注意事项

1. 确保所有当事人信息准确完整
2. 案号、法院名称等必须准确
3. 金额必须同时使用大写和阿拉伯数字
4. 日期格式规范：XXXX 年 X 月 X 日
5. 条款逻辑清晰，无歧义
6. 签署部分预留足够空间
