const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
        TableOfContents } = require("docx");
const fs = require("fs");
const path = require("path");
const brandVI = require("./brand_vi");
const { legalDocConfig, tableStyles, border, createHeader, createFooter } = require("./legal_word_config");

const CONTENT_WIDTH = 9026;

function createTitleParagraph(text, level = 1) {
  if (level === 1) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text, bold: true, font: "微软雅黑", size: 32, color: brandVI.colors.primary.dark })]
    });
  } else if (level === 2) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, font: "微软雅黑", size: 28, color: brandVI.colors.primary.main })]
    });
  } else {
    return new Paragraph({
      heading: HeadingLevel.HEADING_3,
      children: [new TextRun({ text, bold: true, font: "微软雅黑", size: 24, color: brandVI.colors.primary.mid })]
    });
  }
}

function createBodyParagraph(text, options = {}) {
  return new Paragraph({
    spacing: { line: 360, after: 120 },
    children: [new TextRun({
      text,
      font: options.font || "宋体",
      size: options.size || 24,
      color: options.color || brandVI.colors.neutral[800],
      bold: options.bold || false
    })]
  });
}

function createInfoTable(infoData) {
  const colWidths = [2000, 3500, 2000, 3526];
  const cellBorder = {
    top: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    left: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    right: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] }
  };

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: infoData.map(row => new TableRow({
      children: row.map(cell => new TableCell({
        borders: cellBorder,
        width: { size: colWidths[row.indexOf(cell)], type: WidthType.DXA },
        shading: { fill: row.indexOf(cell) % 2 === 0 ? brandVI.colors.neutral[50] : "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: cell,
            font: "宋体",
            size: 21,
            color: brandVI.colors.neutral[800]
          })]
        })]
      }))
    }))
  });
}

function createCaseAnalysisDoc(data) {
  const caseData = data || {
    caseNo: "(2024)湘0124民初1234号",
    type: "股权转让纠纷",
    court: "宁乡市人民法院",
    plaintiff: "原告科技有限公司",
    defendant: "被告贸易有限公司",
    amount: "2,580,000",
    date: "2024年3月15日",
    claim: "判令被告支付股权转让款及违约金共计2,580,000元",
    facts: "原被告于2023年6月签订《股权转让协议》，约定原告将其持有的目标公司30%股权转让给被告，转让款为500万元。协议签订后，被告仅支付了部分款项，尚欠240万元经多次催告未果。",
    analysis: "原被告之间的股权转让关系合法有效。被告未按约定支付全部转让款，已构成违约。根据《民法典》第577条规定，当事人一方不履行合同义务或者履行合同义务不符合约定的，应当承担违约责任。",
    opinion: "建议代理方案：1.立即启动诉讼财产保全；2.主张全部未付转让款；3.按协议约定主张违约金；4.考虑同时追究担保人责任。"
  };

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "宋体", size: 24 } }
      },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "微软雅黑", color: brandVI.colors.primary.dark },
          paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "微软雅黑", color: brandVI.colors.primary.main },
          paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "微软雅黑", color: brandVI.colors.primary.mid },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: { default: createHeader() },
      footers: { default: createFooter() },
      children: [
        new Paragraph({
          alignment: "center",
          spacing: { after: 480 },
          children: [new TextRun({ text: "案件分析报告", bold: true, font: "微软雅黑", size: 44, color: brandVI.colors.primary.dark })]
        }),
        createInfoTable([
          ["案号", caseData.caseNo, "案件类型", caseData.type],
          ["审理法院", caseData.court, "立案日期", caseData.date],
          ["原告", caseData.plaintiff, "被告", caseData.defendant]
        ]),
        new Paragraph({ children: [new PageBreak()] }),
        createTitleParagraph("一、诉讼请求"),
        createBodyParagraph(caseData.claim),
        new Paragraph({ spacing: { after: 240 }, children: [] }),
        createTitleParagraph("二、案件事实"),
        createBodyParagraph(caseData.facts),
        new Paragraph({ spacing: { after: 240 }, children: [] }),
        createTitleParagraph("三、法律分析"),
        createBodyParagraph(caseData.analysis),
        new Paragraph({ spacing: { after: 240 }, children: [] }),
        createTitleParagraph("四、代理意见"),
        createBodyParagraph(caseData.opinion)
      ]
    }]
  });
  return doc;
}

function createContractReviewDoc(data) {
  const reviewData = data || {
    contractTitle: "《股权转让协议》",
    parties: ["转让方: 原告科技有限公司", "受让方: 被告贸易有限公司"],
    reviewDate: "2024年3月10日",
    issues: [
      { severity: "高", title: "付款条件约定不明", desc: "协议第3条对付款条件的约定存在歧义，建议明确具体付款时间节点和条件。" },
      { severity: "中", title: "违约金标准过高", desc: "协议第7条约定的违约金为日万分之十，可能被法院依法调减，建议调整为日万分之五。" },
      { severity: "低", title: "管辖约定", desc: "建议明确约定由宁乡市人民法院管辖，避免后续管辖权争议。" }
    ],
    suggestions: "建议与对方协商修改上述条款，或另行签订补充协议明确相关约定。"
  };

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "宋体", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "微软雅黑", color: brandVI.colors.primary.dark },
          paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "微软雅黑", color: brandVI.colors.primary.main },
          paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 1 } }
      ]
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: createHeader() },
      footers: { default: createFooter() },
      children: [
        new Paragraph({
          alignment: "center",
          spacing: { after: 480 },
          children: [new TextRun({ text: "合同审查报告", bold: true, font: "微软雅黑", size: 44, color: brandVI.colors.primary.dark })]
        }),
        createInfoTable([
          ["合同名称", reviewData.contractTitle, "审查日期", reviewData.reviewDate],
          ["当事人", reviewData.parties.join(" / "), "", ""]
        ]),
        new Paragraph({ children: [new PageBreak()] }),
        createTitleParagraph("一、审查发现问题"),
        ...reviewData.issues.map((issue, i) => new Paragraph({
          spacing: { after: 240 },
          children: [
            new TextRun({ text: `${i + 1}. `, bold: true, font: "微软雅黑", size: 24, color: brandVI.colors.primary.dark }),
            new TextRun({ text: `[${issue.severity}风险] ${issue.title}`, bold: true, font: "微软雅黑", size: 24, color: issue.severity === "高" ? brandVI.colors.semantic.danger : issue.severity === "中" ? brandVI.colors.accent.gold : brandVI.colors.semantic.success }),
            new TextRun({ text: `\n${issue.desc}`, font: "宋体", size: 24, color: brandVI.colors.neutral[700] })
          ]
        })),
        new Paragraph({ spacing: { after: 240 }, children: [] }),
        createTitleParagraph("二、修改建议"),
        createBodyParagraph(reviewData.suggestions)
      ]
    }]
  });
  return doc;
}

function createDueDiligenceDoc(data) {
  const ddData = data || {
    target: "目标公司: 湖南某某有限公司",
    date: "2024年3月5日",
    scope: "有限责任公司股权转让尽职调查",
    items: [
      { category: "公司基本信息", status: "通过", findings: "公司合法存续，股权结构清晰。" },
      { category: "财务状况", status: "关注", findings: "近两年资产负债率上升，建议进一步核实。" },
      { category: "法律诉讼", status: "通过", findings: "无未结重大诉讼。" },
      { category: "资产情况", status: "关注", findings: "部分资产权属证明待完善。" },
      { category: "合同履行", status: "通过", findings: "主要合同正常履行中。" }
    ],
    conclusion: "目标公司整体风险可控，建议在完善相关资产证明后推进交易。"
  };

  const statusColors = { "通过": brandVI.colors.semantic.success, "关注": brandVI.colors.accent.gold, "警示": brandVI.colors.semantic.danger };
  const cellBorder = {
    top: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    left: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] },
    right: { style: BorderStyle.SINGLE, size: 4, color: brandVI.colors.neutral[300] }
  };

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "宋体", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "微软雅黑", color: brandVI.colors.primary.dark },
          paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } }
      ]
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: createHeader() },
      footers: { default: createFooter() },
      children: [
        new Paragraph({
          alignment: "center",
          spacing: { after: 480 },
          children: [new TextRun({ text: "尽职调查报告", bold: true, font: "微软雅黑", size: 44, color: brandVI.colors.primary.dark })]
        }),
        createInfoTable([
          ["调查对象", ddData.target, "调查日期", ddData.date],
          ["调查范围", ddData.scope, "", ""]
        ]),
        new Paragraph({ children: [new PageBreak()] }),
        createTitleParagraph("一、调查事项清单"),
        new Table({
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          columnWidths: [2500, 1500, 5026],
          rows: [
            new TableRow({
              children: [
                new TableCell({ borders: cellBorder, width: { size: 2500, type: WidthType.DXA }, shading: { fill: brandVI.colors.primary.main, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "调查事项", bold: true, font: "微软雅黑", size: 22, color: "FFFFFF" })] })] }),
                new TableCell({ borders: cellBorder, width: { size: 1500, type: WidthType.DXA }, shading: { fill: brandVI.colors.primary.main, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "状态", bold: true, font: "微软雅黑", size: 22, color: "FFFFFF" })] })] }),
                new TableCell({ borders: cellBorder, width: { size: 5026, type: WidthType.DXA }, shading: { fill: brandVI.colors.primary.main, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "调查发现", bold: true, font: "微软雅黑", size: 22, color: "FFFFFF" })] })] })
              ]
            }),
            ...ddData.items.map((item, i) => new TableRow({
              children: [
                new TableCell({ borders: cellBorder, width: { size: 2500, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? brandVI.colors.neutral[50] : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: item.category, font: "宋体", size: 21 })] })] }),
                new TableCell({ borders: cellBorder, width: { size: 1500, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? brandVI.colors.neutral[50] : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: item.status, bold: true, font: "微软雅黑", size: 21, color: statusColors[item.status] })] })] }),
                new TableCell({ borders: cellBorder, width: { size: 5026, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? brandVI.colors.neutral[50] : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: item.findings, font: "宋体", size: 21 })] })] })
              ]
            }))
          ]
        }),
        new Paragraph({ spacing: { after: 360 }, children: [] }),
        createTitleParagraph("二、调查结论"),
        createBodyParagraph(ddData.conclusion)
      ]
    }]
  });
  return doc;
}

function createPleadingDoc(data) {
  const pleadingData = data || {
    type: "民事答辩状",
    caseNo: "(2024)湘0124民初1234号",
    court: "宁乡市人民法院",
    defendant: "被告贸易有限公司",
    plaintiff: "原告科技有限公司",
    response: "答辩人就被答辩人诉答辩人股权转让纠纷一案，提出如下答辩意见：\n\n一、答辩人并非恶意拖欠转让款。答辩人因目标公司经营困难，资金周转暂时出现问题，已积极筹措资金并陆续支付部分款项。\n\n二、被答辩人主张的违约金过高。答辩人请求法院依据《民法典》第585条的规定，以实际损失为基础，参照银行同期贷款利率予以调减。\n\n三、请求法院驳回被答辩人的诉讼请求或依法减少违约金金额。"
  };

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "宋体", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "微软雅黑", color: brandVI.colors.primary.dark },
          paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } }
      ]
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: createHeader() },
      footers: { default: createFooter() },
      children: [
        new Paragraph({
          alignment: "center",
          spacing: { after: 480 },
          children: [new TextRun({ text: pleadingData.type, bold: true, font: "微软雅黑", size: 44, color: brandVI.colors.primary.dark })]
        }),
        createInfoTable([
          ["案号", pleadingData.caseNo, "受理法院", pleadingData.court],
          ["答辩人", pleadingData.defendant, "被答辩人", pleadingData.plaintiff]
        ]),
        new Paragraph({ spacing: { after: 480 }, children: [] }),
        createTitleParagraph("答辩意见"),
        ...pleadingData.response.split("\n\n").map(p => new Paragraph({
          spacing: { line: 360, after: 240 },
          children: [new TextRun({ text: p.replace(/\n/g, ""), font: "宋体", size: 24, color: brandVI.colors.neutral[800] })]
        })),
        new Paragraph({ spacing: { after: 720 }, alignment: "right", children: [new TextRun({ text: "答辩人(签章): _____________", font: "宋体", size: 24 })] }),
        new Paragraph({ alignment: "right", children: [new TextRun({ text: `日期: ${new Date().toLocaleDateString("zh-CN")}`, font: "宋体", size: 24 })] })
      ]
    }]
  });
  return doc;
}

const templates = {
  createCaseAnalysisDoc,
  createContractReviewDoc,
  createDueDiligenceDoc,
  createPleadingDoc
};

async function generateAllTemplates() {
  const outputDir = __dirname;

  console.log("Generating Word templates...");

  templates.createCaseAnalysisDoc().then(doc => {
    Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(path.join(outputDir, "案件分析报告模板.docx"), buffer);
      console.log("Created: 案件分析报告模板.docx");
    });
  });

  templates.createContractReviewDoc().then(doc => {
    Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(path.join(outputDir, "合同审查报告模板.docx"), buffer);
      console.log("Created: 合同审查报告模板.docx");
    });
  });

  templates.createDueDiligenceDoc().then(doc => {
    Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(path.join(outputDir, "尽职调查报告模板.docx"), buffer);
      console.log("Created: 尽职调查报告模板.docx");
    });
  });

  templates.createPleadingDoc().then(doc => {
    Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(path.join(outputDir, "答辩状模板.docx"), buffer);
      console.log("Created: 答辩状模板.docx");
    });
  });

  console.log("All Word templates created!");
}

if (require.main === module) {
  generateAllTemplates();
}

module.exports = templates;