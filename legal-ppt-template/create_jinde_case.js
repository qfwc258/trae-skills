const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const brandVI = require("./brand_vi");
const master = require("./master_slide");

const colors = brandVI.colors;

function createJindeCasePPT() {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = "金得自动化诉汇和消防案件分析";
  pres.company = brandVI.lawFirm.name;

  let slide = pres.addSlide();
  master.titleSlide.apply(slide, pres, {
    mainTitle: "买卖合同纠纷案件分析",
    subTitle: "金得自动化 诉 汇和消防",
    caseNo: "(2025)湘0104民初33092号",
    court: "长沙市岳麓区人民法院"
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "案件概述",
    subtitle: "Case Overview",
    pageNum: 1,
    totalPages: 10
  });

  const caseInfo = [
    { label: "案号", value: "(2025)湘0104民初33092号" },
    { label: "案由", value: "买卖合同纠纷" },
    { label: "审理法院", value: "长沙市岳麓区人民法院" },
    { label: "原告", value: "长沙市望城区金得自动化技术装备厂" },
    { label: "被告", value: "湖南汇和消防工程有限公司" },
    { label: "诉讼标的", value: "约71,110元" }
  ];

  caseInfo.forEach((info, i) => {
    const y = 1.2 + i * 0.65;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.55,
      fill: { color: i % 2 === 0 ? colors.neutral[50] : "FFFFFF" },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addText(info.label, {
      x: 0.7, y: y + 0.12, w: 1.8, h: 0.3,
      fontSize: 14, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(info.value, {
      x: 2.6, y: y + 0.12, w: 6.7, h: 0.3,
      fontSize: 14, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "原告诉讼请求",
    subtitle: "Plaintiff's Claims",
    pageNum: 2,
    totalPages: 10
  });

  slide.addText("原告主张被告支付拖欠货款及利息", {
    x: 0.5, y: 1.3, w: 9, h: 0.5,
    fontSize: 18, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
  });

  const claims = [
    { num: "1", text: "判令被告支付货款本金 71,110元" },
    { num: "2", text: "判令被告按 1.3倍LPR 支付逾期付款利息" },
    { num: "3", text: "判令被告承担律师费" }
  ];

  claims.forEach((claim, i) => {
    const y = 2.0 + i * 1.0;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.85,
      fill: { color: colors.neutral[50] },
      line: { color: colors.primary.light, width: 1 }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.6, h: 0.85,
      fill: { color: colors.primary.main }
    });
    slide.addText(claim.num, {
      x: 0.5, y: y + 0.2, w: 0.6, h: 0.4,
      fontSize: 18, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    slide.addText(claim.text, {
      x: 1.3, y: y + 0.22, w: 8, h: 0.4,
      fontSize: 16, color: colors.neutral[800], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "被告抗辩要点",
    subtitle: "Defense Arguments",
    pageNum: 3,
    totalPages: 10
  });

  const defensePoints = [
    { title: "货款金额异议", content: "实际欠款应为31,110元，而非71,110元", detail: "金鼎公司已于2023年1月20日代为偿付40,000元", color: colors.primary.main },
    { title: "代付事实", content: "40,000元由金鼎公司代付，有对账单备注为证", detail: "原告方对账单备注载明'金鼎2023年1月20日付长沙林茂40000元（算代湖南汇和付）'", color: colors.primary.mid },
    { title: "主体混同", content: "原告与林茂公司存在人格混同", detail: "两公司注册地址相同，经办人均为金辉", color: colors.accent.gold }
  ];

  defensePoints.forEach((point, i) => {
    const y = 1.15 + i * 1.35;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 1.2,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.1, h: 1.2,
      fill: { color: point.color }
    });
    slide.addText(point.title, {
      x: 0.8, y: y + 0.1, w: 8.5, h: 0.4,
      fontSize: 15, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(point.content, {
      x: 0.8, y: y + 0.5, w: 8.5, h: 0.3,
      fontSize: 13, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(point.detail, {
      x: 0.8, y: y + 0.8, w: 8.5, h: 0.3,
      fontSize: 11, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "核心争议焦点",
    subtitle: "Key Issues",
    pageNum: 4,
    totalPages: 10
  });

  const issues = [
    { focus: "代付效力", question: "金鼎公司支付的40,000元能否认定为被告的付款？", analysis: "需审查对账单备注的效力及各方合意" },
    { focus: "主体认定", question: "原告与林茂公司是否构成人格混同？", analysis: "影响代付行为的效力认定" },
    { focus: "利息计算", question: "原告主张的1.3倍LPR利息及律师费是否应支持？", analysis: "需审查合同约定及法律规定" }
  ];

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 9, h: 4.2,
    fill: { color: colors.neutral[50] },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });

  issues.forEach((issue, i) => {
    const y = 1.25 + i * 1.35;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.7, y: y, w: 1.8, h: 0.5,
      fill: { color: colors.primary.dark }
    });
    slide.addText(`焦点${i + 1}`, {
      x: 0.7, y: y + 0.08, w: 1.8, h: 0.35,
      fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    slide.addText(issue.focus, {
      x: 2.7, y: y + 0.08, w: 6.6, h: 0.35,
      fontSize: 15, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(issue.question, {
      x: 0.9, y: y + 0.55, w: 8.4, h: 0.35,
      fontSize: 13, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(issue.analysis, {
      x: 0.9, y: y + 0.9, w: 8.4, h: 0.3,
      fontSize: 11, color: colors.accent.gold, fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "证据分析",
    subtitle: "Evidence Analysis",
    pageNum: 5,
    totalPages: 10
  });

  const evidenceGroups = [
    {
      category: "微信记录",
      items: [
        { name: "汇和公司与金辉聊天记录", weight: "证明代付合意", status: "✓" }
      ]
    },
    {
      category: "支付凭证",
      items: [
        { name: "金鼎公司付款凭证", weight: "证明代付事实", status: "✓" }
      ]
    },
    {
      category: "工商信息",
      items: [
        { name: "金得厂与林茂公司工商登记", weight: "证明主体混同", status: "关键" }
      ]
    }
  ];

  evidenceGroups.forEach((group, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 2.9, h: 3.8,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 2.9, h: 0.5,
      fill: { color: colors.primary.main }
    });
    slide.addText(group.category, {
      x: x, y: 1.2, w: 2.9, h: 0.3,
      fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    group.items.forEach((item, j) => {
      const itemY = 1.8 + j * 1.2;
      const statusColor = item.status === "关键" ? colors.accent.gold : colors.semantic.success;
      slide.addShape(pres.ShapeType.rect, {
        x: x + 0.15, y: itemY, w: 2.6, h: 1.0,
        fill: { color: colors.neutral[50] },
        line: { color: colors.neutral[200], width: 0.5 }
      });
      slide.addText(item.status, {
        x: x + 0.25, y: itemY + 0.1, w: 0.5, h: 0.4,
        fontSize: 16, bold: true, color: statusColor, fontFace: "Microsoft YaHei", margin: 0
      });
      slide.addText(item.name, {
        x: x + 0.25, y: itemY + 0.5, w: 2.4, h: 0.4,
        fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "法律依据",
    subtitle: "Legal Basis",
    pageNum: 6,
    totalPages: 10
  });

  const laws = [
    { type: "法律", name: "《民法典》第577条", articles: "违约责任", content: "当事人一方不履行合同义务或者履行合同义务不符合约定的，应当承担违约责任" },
    { type: "司法解释", name: "《买卖合同司法解释》", articles: "第18条", content: "买卖合同对付款期限作出的变更，不影响当事人关于逾期付款违约金的约定" },
    { type: "法律", name: "《民法典》第523条", articles: "第三人代为履行", content: "当事人约定由第三人向债权人履行债务，第三人不履行债务或者履行债务不符合约定的，债务人应当向债权人承担违约责任" }
  ];

  laws.forEach((law, i) => {
    const y = 1.1 + i * 1.35;
    const typeColor = law.type === "法律" ? colors.primary.main : colors.primary.mid;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 1.2,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.2, h: 1.2,
      fill: { color: typeColor }
    });
    slide.addText(law.type, {
      x: 0.5, y: y + 0.4, w: 1.2, h: 0.35,
      fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    slide.addText(law.name, {
      x: 1.9, y: y + 0.15, w: 7.4, h: 0.35,
      fontSize: 14, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(law.articles, {
      x: 1.9, y: y + 0.5, w: 7.4, h: 0.25,
      fontSize: 11, color: colors.accent.gold, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(law.content, {
      x: 1.9, y: y + 0.8, w: 7.4, h: 0.35,
      fontSize: 10, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "诉讼策略建议",
    subtitle: "Litigation Strategy",
    pageNum: 7,
    totalPages: 10
  });

  const strategies = [
    { title: "核心策略", content: "主张代付有效，货款金额应扣减40,000元", icon: "★", color: colors.primary.dark },
    { title: "证据策略", content: "重点举证对账单备注及微信记录，证明代付合意", icon: "◆", color: colors.primary.main },
    { title: "程序策略", content: "申请法院调查金鼎公司与林茂公司付款事实", icon: "●", color: colors.primary.mid },
    { title: "调解策略", content: "如调解，可考虑承担部分利息以达成和解", icon: "○", color: colors.neutral[500] }
  ];

  strategies.forEach((s, i) => {
    const y = 1.1 + i * 1.05;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.9,
      fill: { color: colors.neutral[50] },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.08, h: 0.9,
      fill: { color: s.color }
    });
    slide.addText(s.icon, {
      x: 0.75, y: y + 0.2, w: 0.5, h: 0.5,
      fontSize: 18, color: s.color, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(s.title, {
      x: 1.4, y: y + 0.15, w: 2, h: 0.35,
      fontSize: 14, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(s.content, {
      x: 1.4, y: y + 0.5, w: 7.8, h: 0.35,
      fontSize: 12, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "风险提示",
    subtitle: "Risk Warning",
    pageNum: 8,
    totalPages: 10
  });

  const risks = [
    { level: "中", title: "代付认定风险", content: "如法院不认可代付事实，将承担全部71,110元货款及利息", color: colors.accent.gold },
    { level: "低", title: "利息风险", content: "如合同未明确约定律师费，可能需原告自行承担", color: colors.semantic.success },
    { level: "高", title: "执行风险", content: "被告经营状况不明，需关注财产保全", color: colors.semantic.danger }
  ];

  risks.forEach((risk, i) => {
    const y = 1.1 + i * 1.35;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 1.2,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.0, h: 1.2,
      fill: { color: risk.color }
    });
    slide.addText(risk.level, {
      x: 0.5, y: y + 0.4, w: 1.0, h: 0.4,
      fontSize: 16, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    slide.addText(risk.title, {
      x: 1.7, y: y + 0.15, w: 7.6, h: 0.4,
      fontSize: 15, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(risk.content, {
      x: 1.7, y: y + 0.6, w: 7.6, h: 0.5,
      fontSize: 12, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide = pres.addSlide();
  master.contentSlide.apply(slide, pres, {
    title: "损失计算",
    subtitle: "Damage Calculation",
    pageNum: 9,
    totalPages: 10
  });

  const calcItems = [
    { item: "货款本金（原告主张）", amount: "71,110", days: "-", rate: "-", interest: "-" },
    { item: "实际欠款（被告主张）", amount: "31,110", days: "-", rate: "-", interest: "扣减40,000" },
    { item: "逾期利息（按LPR1.3倍估算）", amount: "31,110", days: "约365天", rate: "年化5.46%", interest: "约1,700" }
  ];

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 9, h: 3.2,
    fill: { color: colors.white },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });

  const headers = ["项目", "本金(元)", "天数", "利率", "备注"];
  const widths = [3.0, 1.5, 1.2, 1.5, 1.8];
  let xPos = 0.5;
  headers.forEach((h, i) => {
    slide.addShape(pres.ShapeType.rect, {
      x: xPos, y: 1.1, w: widths[i], h: 0.5,
      fill: { color: colors.primary.dark }
    });
    slide.addText(h, {
      x: xPos, y: 1.2, w: widths[i], h: 0.3,
      fontSize: 12, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    xPos += widths[i];
  });

  calcItems.forEach((item, i) => {
    const y = 1.7 + i * 0.7;
    const bgColor = i === 1 ? colors.neutral[100] : colors.white;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.6,
      fill: { color: bgColor },
      line: { color: colors.neutral[200], width: 0.5 }
    });
    xPos = 0.5;
    const values = [item.item, item.amount, item.days, item.rate, item.interest];
    values.forEach((val, j) => {
      slide.addText(val, {
        x: xPos, y: y + 0.15, w: widths[j], h: 0.3,
        fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", align: "center", margin: 0
      });
      xPos += widths[j];
    });
  });

  slide = pres.addSlide();
  master.endingSlide.apply(slide, pres);

  const outputPath = "D:\\0金厚\\金得汇和\\金得自动化诉汇和消防案件分析.pptx";
  pres.writeFile({ fileName: outputPath })
    .then(() => console.log(`PPT created: ${outputPath}`))
    .catch(err => console.error("Error:", err));
}

createJindeCasePPT();