const PptxGenJS = require("pptxgenjs");
const brandVI = require("./brand_vi");
const master = require("./master_slide");

const defaultConfig = { brandVI };

function resolveColor(colors, colorRef) {
  if (!colorRef) return colors.primary.main;
  if (colorRef.includes(".")) {
    const [group, key] = colorRef.split(".");
    return colors[group]?.[key] || colors.primary.main;
  }
  return colors[colorRef] || colors.primary.main;
}

function addPageHeader(slide, pres, config, title, subtitle, pageNum, totalPages) {
  const vi = { ...brandVI, ...(config?.brandVI || {}) };
  const colors = vi.colors;
  const typo = vi.typography;

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: colors.primary.deep }
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0.06, w: 10, h: 0.02,
    fill: { color: colors.accent.gold }
  });

  slide.addText(title, {
    x: vi.spacing.pageMargin, y: 0.25, w: 5, h: 0.5,
    fontSize: typo.h2.fontSize, bold: typo.h2.bold,
    color: resolveColor(colors, typo.h2.color),
    fontFace: "Microsoft YaHei", margin: 0
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 5.5, y: 0.32, w: 2.5, h: 0.35,
      fontSize: 12, color: colors.neutral[400],
      fontFace: "Microsoft YaHei", align: "right", margin: 0
    });
  }

  slide.addText(`${String(pageNum).padStart(2, "0")} / ${String(totalPages).padStart(2, "0")}`, {
    x: 8.3, y: 0.32, w: 1.1, h: 0.35,
    fontSize: 11, color: colors.neutral[500],
    fontFace: "Microsoft YaHei", align: "right", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: vi.spacing.pageMargin, y: 0.82, w: 1.5, h: 0.04,
    fill: { color: colors.accent.gold }
  });
}

function addPageFooter(slide, pres, config) {
  const vi = { ...brandVI, ...(config?.brandVI || {}) };
  const colors = vi.colors;

  slide.addShape(pres.ShapeType.line, {
    x: vi.spacing.pageMargin, y: 5.3, w: 8.8, h: 0,
    line: { color: colors.neutral[200], width: 0.5 }
  });

  slide.addText(vi.lawFirm.name, {
    x: vi.spacing.pageMargin, y: 5.35, w: 4, h: 0.22,
    fontSize: 9, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(`TEL: ${vi.lawFirm.tel}`, {
    x: 5.5, y: 5.35, w: 2, h: 0.22,
    fontSize: 9, color: colors.neutral[500], fontFace: "Microsoft YaHei", align: "center", margin: 0
  });

  slide.addText(vi.lawFirm.branch, {
    x: 7.5, y: 5.35, w: 1.8, h: 0.22,
    fontSize: 9, color: colors.neutral[500], fontFace: "Microsoft YaHei", align: "right", margin: 0
  });
}

function createTitleSlide(pres, config, data) {
  const vi = { ...brandVI, ...(config?.brandVI || {}) };
  const colors = vi.colors;

  let slide = pres.addSlide();
  slide.background = { color: colors.primary.deep };

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: colors.primary.dark, transparency: 30 }
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 2.2, w: 0.15, h: 1.2,
    fill: { color: colors.accent.gold }
  });

  slide.addText(config?.title || "法律案件分析报告", {
    x: 0.6, y: 2.2, w: 8, h: 0.8,
    fontSize: 40, bold: true, color: colors.neutral[50],
    fontFace: "Microsoft YaHei", margin: 0
  });

  if (config?.case) {
    slide.addText(`${config.case.type}  ·  ${config.case.court}`, {
      x: 0.6, y: 3.0, w: 8, h: 0.4,
      fontSize: 16, color: colors.neutral[300], fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(config.case.caseNo, {
      x: 0.6, y: 3.4, w: 8, h: 0.35,
      fontSize: 14, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
    });
  }

  slide.addText(vi.lawFirm.name, {
    x: 0.6, y: 4.6, w: 5, h: 0.35,
    fontSize: 14, color: colors.accent.goldLight, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(vi.lawFirm.tel, {
    x: 0.6, y: 4.95, w: 3, h: 0.3,
    fontSize: 11, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
  });

  return slide;
}

function createCaseOverviewSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "案件概述",
    subtitle: "Case Overview",
    pageNum: 1,
    totalPages: 10
  });

  const caseData = data || {
    caseNo: "(2024)湘0124民初1234号",
    type: "合同纠纷",
    court: "宁乡市人民法院",
    judge: "张法官",
    plaintiff: "原告科技有限公司",
    defendant: "被告贸易有限公司",
    amount: "¥2,580,000",
    date: "2024年3月15日"
  };

  const infoBlocks = [
    { label: "案号", value: caseData.caseNo, x: 0.6 },
    { label: "案件类型", value: caseData.type, x: 3.4 },
    { label: "审理法院", value: caseData.court, x: 6.2 }
  ];

  infoBlocks.forEach(block => {
    slide.addShape(pres.ShapeType.rect, {
      x: block.x, y: 1.2, w: 2.6, h: 1.0,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addText(block.label, {
      x: block.x + 0.15, y: 1.3, w: 2.3, h: 0.25,
      fontSize: 10, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(block.value, {
      x: block.x + 0.15, y: 1.55, w: 2.3, h: 0.5,
      fontSize: 13, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });
  });

  const parties = [
    { role: "原告", name: caseData.plaintiff, color: colors.semantic.success },
    { role: "被告", name: caseData.defendant, color: colors.semantic.danger }
  ];

  parties.forEach((p, i) => {
    const x = 0.6 + i * 4.7;
    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 2.5, w: 4.3, h: 1.4,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 2.5, w: 0.08, h: 1.4,
      fill: { color: p.color }
    });
    slide.addText(p.role, {
      x: x + 0.25, y: 2.6, w: 1, h: 0.3,
      fontSize: 11, color: p.color, fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(p.name, {
      x: x + 0.25, y: 2.95, w: 3.8, h: 0.8,
      fontSize: 15, bold: true, color: colors.neutral[800], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.6, y: 4.15, w: 8.8, h: 1.0,
    fill: { color: colors.primary.dark }
  });
  slide.addText("诉讼标的金额", {
    x: 0.8, y: 4.25, w: 2, h: 0.3,
    fontSize: 11, color: colors.neutral[300], fontFace: "Microsoft YaHei", margin: 0
  });
  slide.addText(caseData.amount, {
    x: 0.8, y: 4.55, w: 4, h: 0.5,
    fontSize: 28, bold: true, color: colors.accent.gold, fontFace: "Microsoft YaHei", margin: 0
  });
  slide.addText(`立案日期: ${caseData.date}`, {
    x: 6.5, y: 4.55, w: 2.5, h: 0.4,
    fontSize: 12, color: colors.neutral[300], fontFace: "Microsoft YaHei", align: "right", margin: 0
  });

  addPageFooter(slide, pres, config);
  return slide;
}

function createTimelineSlide(pres, config, events) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "案件时间线",
    subtitle: "Timeline",
    pageNum: 2,
    totalPages: 10
  });

  const timeline = events || [
    { date: "2023-08-15", title: "合同签订", desc: "双方签订采购合同", status: "completed" },
    { date: "2023-10-20", title: "违约发生", desc: "被告未按期付款", status: "completed" },
    { date: "2024-01-08", title: "律师函", desc: "发送催告函", status: "completed" },
    { date: "2024-02-28", title: "立案", desc: "向法院提交诉状", status: "current" },
    { date: "2024-04-15", title: "开庭", desc: "一审开庭审理", status: "pending" },
    { date: "2024-06-30", title: "判决", desc: "预计一审判决", status: "pending" }
  ];

  const startX = 0.6;
  const lineY = 2.0;
  const spacing = 1.5;

  slide.addShape(pres.ShapeType.line, {
    x: startX + 0.15, y: lineY, w: 8.7, h: 0,
    line: { color: colors.neutral[300], width: 2 }
  });

  timeline.forEach((item, i) => {
    const x = startX + i * spacing;
    const isCompleted = item.status === "completed";
    const isCurrent = item.status === "current";
    const dotColor = isCompleted ? colors.semantic.success : isCurrent ? colors.accent.gold : colors.neutral[300];

    slide.addShape(pres.ShapeType.ellipse, {
      x: x, y: lineY - 0.15, w: 0.3, h: 0.3,
      fill: { color: dotColor }
    });

    slide.addText(item.date, {
      x: x - 0.3, y: lineY + 0.3, w: 0.9, h: 0.25,
      fontSize: 9, color: colors.neutral[500], fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(item.title, {
      x: x - 0.4, y: lineY + 0.6, w: 1.2, h: 0.3,
      fontSize: 11, bold: true, color: isCurrent ? colors.accent.gold : colors.primary.dark,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(item.desc, {
      x: x - 0.5, y: lineY + 0.9, w: 1.4, h: 0.6,
      fontSize: 9, color: colors.neutral[600], fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
  });

  return slide;
}

function createSwotSlide(pres, config, swotData) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "SWOT分析",
    subtitle: "Strategic Analysis",
    pageNum: 3,
    totalPages: 10
  });

  const swotItems = [
    { title: "优势", subtitle: "Strengths", icon: "↑", color: colors.semantic.success, items: swotData?.strengths || ["证据链完整充分", "合同条款明确", "违约事实清楚"] },
    { title: "劣势", subtitle: "Weaknesses", icon: "↓", color: colors.semantic.danger, items: swotData?.weaknesses || ["标的金额较大", "执行周期较长"] },
    { title: "机会", subtitle: "Opportunities", icon: "→", color: colors.primary.main, items: swotData?.opportunities || ["对方经营困难", "存在和解空间"] },
    { title: "威胁", subtitle: "Threats", icon: "⚠", color: colors.primary.dark, items: swotData?.threats || ["对方拖延诉讼", "执行难度大"] }
  ];

  const positions = [
    { x: 0.5, y: 1.15 },
    { x: 5.05, y: 1.15 },
    { x: 0.5, y: 3.35 },
    { x: 5.05, y: 3.35 }
  ];

  swotItems.forEach((item, i) => {
    const pos = positions[i];

    slide.addShape(pres.ShapeType.rect, {
      x: pos.x, y: pos.y, w: 4.45, h: 2.0,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: pos.x, y: pos.y, w: 4.45, h: 0.5,
      fill: { color: item.color }
    });

    slide.addText(`${item.icon} ${item.title}`, {
      x: pos.x + 0.15, y: pos.y + 0.08, w: 2, h: 0.35,
      fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(item.subtitle, {
      x: pos.x + 3.2, y: pos.y + 0.1, w: 1.1, h: 0.3,
      fontSize: 9, color: colors.neutral[300], fontFace: "Microsoft YaHei", align: "right", margin: 0
    });

    item.items.forEach((text, j) => {
      slide.addText(`• ${text}`, {
        x: pos.x + 0.2, y: pos.y + 0.65 + j * 0.4, w: 4, h: 0.35,
        fontSize: 12, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  return slide;
}

function createLegalElementSlide(pres, config, elements) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "法律要件分析",
    subtitle: "Legal Elements",
    pageNum: 4,
    totalPages: 10
  });

  const legalElements = elements || [
    { name: "合同成立", detail: "双方意思表示真实，合同依法成立", status: "satisfied" },
    { name: "合同生效", detail: "合同符合生效要件，已实际生效", status: "satisfied" },
    { name: "履行义务", detail: "原告已履行合同约定的全部义务", status: "satisfied" },
    { name: "违约事实", detail: "被告存在逾期付款的违约行为", status: "satisfied" },
    { name: "损害结果", detail: "原告因被告违约遭受实际经济损失", status: "satisfied" },
    { name: "因果关系", detail: "被告违约与原告损失存在直接因果关系", status: "current" },
    { name: "主观过错", detail: "被告明知违约仍故意拖延履行", status: "pending" }
  ];

  legalElements.forEach((el, i) => {
    const y = 1.1 + i * 0.58;
    const statusColor = el.status === "satisfied" ? colors.semantic.success : el.status === "current" ? colors.accent.gold : colors.neutral[400];

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.52,
      fill: { color: i % 2 === 0 ? colors.neutral[50] : colors.white }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.06, h: 0.52,
      fill: { color: statusColor }
    });

    slide.addText(el.name, {
      x: 0.7, y: y + 0.08, w: 1.5, h: 0.35,
      fontSize: 13, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(el.detail, {
      x: 2.2, y: y + 0.1, w: 5.5, h: 0.35,
      fontSize: 11, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });

    const statusText = el.status === "satisfied" ? "✓ 满足" : el.status === "current" ? "◐ 审查中" : "○ 待证";
    slide.addText(statusText, {
      x: 8, y: y + 0.1, w: 1.3, h: 0.35,
      fontSize: 11, color: statusColor, fontFace: "Microsoft YaHei", align: "right", margin: 0
    });
  });

  return slide;
}

function createAmountSlide(pres, config, amounts) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "诉讼请求与金额",
    subtitle: "Claims & Amounts",
    pageNum: 5,
    totalPages: 10
  });

  const claimItems = amounts || [
    { title: "货款本金", amount: "2,180,000", rate: "84.5%" },
    { title: "逾期利息", amount: "286,400", rate: "11.1%" },
    { title: "违约金", amount: "113,600", rate: "4.4%" },
    { title: "合计", amount: "2,580,000", rate: "100%", isTotal: true }
  ];

  claimItems.forEach((item, i) => {
    const y = 1.2 + i * 0.9;
    const isTotal = item.isTotal;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 5.5, h: 0.75,
      fill: { color: isTotal ? colors.primary.dark : colors.white },
      line: { color: isTotal ? colors.primary.dark : colors.neutral[200], width: brandVI.border.width }
    });

    slide.addText(item.title, {
      x: 0.7, y: y + 0.2, w: 2, h: 0.35,
      fontSize: isTotal ? 14 : 13, bold: isTotal, color: isTotal ? colors.white : colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(`¥${item.amount}`, {
      x: 2.5, y: y + 0.15, w: 3.3, h: 0.45,
      fontSize: isTotal ? 22 : 18, bold: true, color: isTotal ? colors.accent.gold : colors.primary.dark, fontFace: "Microsoft YaHei", align: "right", margin: 0
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 6.2, y: y, w: 3.3, h: 0.75,
      fill: { color: isTotal ? colors.primary.deep : colors.neutral[100] }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 6.2, y: y, w: parseFloat(item.rate) / 100 * 3.3, h: 0.75,
      fill: { color: isTotal ? colors.accent.gold : colors.primary.mid }
    });

    slide.addText(item.rate, {
      x: 6.4, y: y + 0.2, w: 3, h: 0.35,
      fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0
    });
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fill: { color: colors.white },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });
  slide.addText("注：利息计算至实际清偿之日止，按同期LPR四倍标准", {
    x: 0.7, y: 4.9, w: 8.5, h: 0.3,
    fontSize: 10, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
  });

  addPageFooter(slide, pres, config);
  return slide;
}

function createProgressSlide(pres, config, stages) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "诉讼进度",
    subtitle: "Litigation Progress",
    pageNum: 6,
    totalPages: 10
  });

  slide.addChart(pres.ChartType.doughnut, [{
    name: "进度",
    labels: ["已完成", "进行中", "未开始"],
    values: [35, 25, 40]
  }], {
    x: 5.8, y: 1.0, w: 3.8, h: 3.5,
    chartColors: [colors.semantic.success, colors.accent.gold, colors.neutral[200]],
    showLegend: true,
    legendPos: "b",
    showValue: false,
    showPercent: true,
    showTitle: false
  });

  const progressStages = stages || [
    { name: "立案", status: "completed" },
    { name: "送达", status: "completed" },
    { name: "举证", status: "completed" },
    { name: "开庭", status: "current" },
    { name: "判决", status: "pending" },
    { name: "执行", status: "pending" }
  ];

  progressStages.forEach((stage, i) => {
    const x = 0.7 + i * 0.85;
    const isCompleted = stage.status === "completed";
    const isCurrent = stage.status === "current";

    slide.addShape(pres.ShapeType.ellipse, {
      x: x, y: 1.5, w: 0.65, h: 0.65,
      fill: { color: isCompleted ? colors.semantic.success : isCurrent ? colors.accent.gold : colors.neutral[100] },
      line: { color: isCurrent ? colors.primary.dark : colors.neutral[300], width: isCurrent ? 2 : 0.5 }
    });

    slide.addText(stage.name, {
      x: x - 0.15, y: 2.25, w: 0.95, h: 0.3,
      fontSize: 10, bold: true, color: isCurrent ? colors.primary.dark : colors.neutral[600], fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    if (i < progressStages.length - 1) {
      slide.addShape(pres.ShapeType.line, {
        x: x + 0.65, y: 1.82, w: 0.2, h: 0,
        line: { color: isCompleted ? colors.semantic.success : colors.neutral[300], width: 2 }
      });
    }
  });

  const summaryItems = [
    { label: "当前阶段", value: "一审开庭", color: colors.accent.gold },
    { label: "剩余时间", value: "约60天", color: colors.primary.main },
    { label: "胜诉概率", value: "85%", color: colors.semantic.success }
  ];

  summaryItems.forEach((item, i) => {
    const y = 3.0 + i * 0.7;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.6, y: y, w: 4.5, h: 0.55,
      fill: { color: colors.neutral[50] }
    });
    slide.addShape(pres.ShapeType.rect, {
      x: 0.6, y: y, w: 0.06, h: 0.55,
      fill: { color: item.color }
    });
    slide.addText(item.label, {
      x: 0.8, y: y + 0.12, w: 1.5, h: 0.3,
      fontSize: 11, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
    });
    slide.addText(item.value, {
      x: 2.5, y: y + 0.08, w: 2.4, h: 0.4,
      fontSize: 16, bold: true, color: item.color, fontFace: "Microsoft YaHei", align: "right", margin: 0
    });
  });

  return slide;
}

function createStrategySlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "代理策略",
    subtitle: "Litigation Strategy",
    pageNum: 7,
    totalPages: 10
  });

  const strategies = data || [
    { phase: "第一阶段", title: "财产保全", desc: "申请冻结被告银行账户，确保执行可能性", priority: "high" },
    { phase: "第二阶段", title: "证据攻势", desc: "系统整理合同履行证据，准备举证材料", priority: "high" },
    { phase: "第三阶段", title: "谈判和解", desc: "在庭审前尝试调解，争取有利结果", priority: "medium" },
    { phase: "第四阶段", title: "庭审应对", desc: "强化代理意见，积极应对庭审程序", priority: "medium" }
  ];

  const priorityColors = { high: colors.semantic.danger, medium: colors.accent.gold, low: colors.primary.main };

  strategies.forEach((item, i) => {
    const y = 1.1 + i * 1.05;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.9,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.2, h: 0.9,
      fill: { color: priorityColors[item.priority] }
    });

    slide.addText(item.phase, {
      x: 0.55, y: y + 0.28, w: 1.1, h: 0.35,
      fontSize: 12, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(item.title, {
      x: 1.9, y: y + 0.15, w: 3, h: 0.35,
      fontSize: 15, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(item.desc, {
      x: 1.9, y: y + 0.5, w: 7.3, h: 0.35,
      fontSize: 12, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });

    const priorityLabel = item.priority === "high" ? "优先" : item.priority === "medium" ? "常规" : "备选";
    slide.addText(priorityLabel, {
      x: 8.3, y: y + 0.3, w: 1, h: 0.3,
      fontSize: 10, color: priorityColors[item.priority], fontFace: "Microsoft YaHei", align: "right", margin: 0
    });
  });

  return slide;
}

function createDefenseSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "对方抗辩分析",
    subtitle: "Defense Analysis",
    pageNum: 8,
    totalPages: 10
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 4.3, h: 2.2,
    fill: { color: colors.neutral[50] },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });
  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 0.08, h: 2.2,
    fill: { color: colors.semantic.danger }
  });

  slide.addText("对方核心主张", {
    x: 0.75, y: 1.2, w: 3.8, h: 0.35,
    fontSize: 13, bold: true, color: colors.semantic.danger, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(data?.opponentClaim || "• 货物质量存在瑕疵\n• 付款条件未成就\n• 原告亦有违约行为", {
    x: 0.75, y: 1.6, w: 3.8, h: 1.5,
    fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 5.2, y: 1.1, w: 4.3, h: 2.2,
    fill: { color: colors.neutral[50] },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });
  slide.addShape(pres.ShapeType.rect, {
    x: 5.2, y: 1.1, w: 0.08, h: 2.2,
    fill: { color: colors.semantic.success }
  });

  slide.addText("我方反驳要点", {
    x: 5.45, y: 1.2, w: 3.8, h: 0.35,
    fontSize: 13, bold: true, color: colors.semantic.success, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(data?.rebuttal || "• 验收记录显示质量合格\n• 发票签收证明付款期限\n• 对方违约在先", {
    x: 5.45, y: 1.6, w: 3.8, h: 1.5,
    fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 3.55, w: 9, h: 1.5,
    fill: { color: colors.primary.dark }
  });

  slide.addText("关键结论", {
    x: 0.7, y: 3.65, w: 8.5, h: 0.35,
    fontSize: 13, bold: true, color: colors.accent.gold, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(data?.conclusion || "综合分析，我方证据链完整，对方抗辩理由缺乏事实和法律依据，预计庭审中可有效反驳对方观点，争取有利判决。", {
    x: 0.7, y: 4.05, w: 8.5, h: 0.9,
    fontSize: 12, color: colors.neutral[100], fontFace: "Microsoft YaHei", margin: 0
  });

  return slide;
}

function createEvidenceListSlide(pres, config, categories) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "关键证据清单",
    subtitle: "Evidence List",
    pageNum: 9,
    totalPages: 10
  });

  const evidenceCategories = categories || [
    { title: "合同文件", items: ["采购合同原件", "补充协议", "招标文件"] },
    { title: "履行证据", items: ["送货单据", "验收确认书", "发票清单"] },
    { title: "违约证明", items: ["催告函", "律师函", "往来函件"] },
    { title: "损失依据", items: ["财务凭证", "利息计算表", "执行申请书"] }
  ];

  evidenceCategories.forEach((cat, i) => {
    const x = 0.5 + (i % 2) * 4.7;
    const y = 1.1 + Math.floor(i / 2) * 2.15;

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: y, w: 4.4, h: 2.0,
      fill: { color: colors.neutral[50] },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: y, w: 4.4, h: 0.45,
      fill: { color: colors.primary.main }
    });

    slide.addText(cat.title, {
      x: x + 0.15, y: y + 0.08, w: 4, h: 0.3,
      fontSize: 13, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0
    });

    cat.items.forEach((item, j) => {
      slide.addText(`☐ ${item}`, {
        x: x + 0.2, y: y + 0.55 + j * 0.35, w: 4, h: 0.3,
        fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  return slide;
}

function createBackgroundSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "案件背景",
    subtitle: "Case Background",
    pageNum: 1,
    totalPages: 16
  });

  const backgrounds = data || [
    { title: "股权转让背景", content: "2022年3月，被告张三将其持有的湖南星辉科技有限公司30%股权（对应注册资本300万元）转让给原告，双方签订《股权转让协议》约定转让款为1000万元，分三期支付。" },
    { title: "股权变更完成", content: "2022年3月20日，双方完成工商变更登记，原告正式成为持有公司30%股权的股东，并开始参与公司经营管理。" },
    { title: "违约事实发生", content: "截至2023年6月30日，被告仅支付第一期款项400万元，第二期款项400万元及第三期款项200万元均未按约支付。" },
    { title: "一审诉讼经过", content: "原告于2024年1月向岳麓区人民法院提起诉讼，一审法院判决支持原告全部诉讼请求，被告不服提起上诉。" }
  ];

  backgrounds.forEach((item, i) => {
    const y = 1.1 + i * 1.05;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.95,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.08, h: 0.95,
      fill: { color: colors.primary.main }
    });

    slide.addText(item.title, {
      x: 0.75, y: y + 0.1, w: 8.5, h: 0.3,
      fontSize: 13, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(item.content, {
      x: 0.75, y: y + 0.45, w: 8.5, h: 0.45,
      fontSize: 11, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  addPageFooter(slide, pres, config);
  return slide;
}

function createPartyInfoSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "当事人信息",
    subtitle: "Party Information",
    pageNum: 2,
    totalPages: 16
  });

  const parties = data || {
    plaintiff: {
      name: "湖南星辉科技有限公司",
      representative: "李明（法定代表人）",
      role: "上诉人（原审原告）",
      addr: "长沙市高新区...",
      background: "成立于2015年，注册资本1000万元，主要从事软件开发业务"
    },
    defendant: {
      name: "张三",
      idCard: "4301051978**********",
      role: "被上诉人（原审被告）",
      addr: "长沙市岳麓区...",
      background: "原星辉公司股东，持有30%股权，现已转让全部股权"
    }
  };

  Object.entries(parties).forEach(([key, party], i) => {
    const x = 0.5 + i * 4.7;
    const isPlaintiff = key === "plaintiff";

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 4.4, h: 4.1,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 4.4, h: 0.55,
      fill: { color: isPlaintiff ? colors.primary.main : colors.semantic.danger }
    });

    slide.addText(isPlaintiff ? "原告方" : "被告方", {
      x: x + 0.15, y: 1.2, w: 4, h: 0.35,
      fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0
    });

    const fields = [
      { label: "名称/姓名", value: party.name },
      { label: "身份标识", value: party.idCard || party.representative },
      { label: "诉讼地位", value: party.role },
      { label: "住所地", value: party.addr },
      { label: "背景", value: party.background }
    ];

    fields.forEach((field, j) => {
      const fieldY = 1.8 + j * 0.6;
      slide.addText(field.label, {
        x: x + 0.15, y: fieldY, w: 1.2, h: 0.25,
        fontSize: 10, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
      });
      slide.addText(field.value, {
        x: x + 0.15, y: fieldY + 0.25, w: 4.1, h: 0.3,
        fontSize: 11, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  return slide;
}

function createDisputeFocusSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "争议焦点",
    subtitle: "Key Dispute Points",
    pageNum: 3,
    totalPages: 16
  });

  const focusPoints = data || [
    { point: "焦点一", title: "转让款支付条件是否成就", plaintiff: "被告未按约定时间支付第二、三期款项", defendant: "原告未协助办理资质变更手续", result: "被告违约" },
    { point: "焦点二", title: "违约金金额如何认定", plaintiff: "按协议约定日万分之五计算", defendant: "约定过高，请求法院调整", result: "需法院酌定" },
    { point: "焦点三", title: "利息计算的起止时间", plaintiff: "自各期应付之日起算", defendant: "应从催告后开始计算", result: "待二审认定" },
    { point: "焦点四", title: "被告主张的抗辩是否成立", plaintiff: "被告抗辩缺乏事实依据", defendant: "存在情势变更情形", result: "一审未采信" }
  ];

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 9, h: 4.15,
    fill: { color: colors.neutral[50] }
  });

  const headers = [
    { text: "焦点", x: 0.5, w: 0.8 },
    { text: "争议事项", x: 1.3, w: 2.2 },
    { text: "原告主张", x: 3.5, w: 2.0 },
    { text: "被告抗辩", x: 5.5, w: 2.0 },
    { text: "现状", x: 7.5, w: 2.0 }
  ];

  headers.forEach(h => {
    slide.addShape(pres.ShapeType.rect, {
      x: h.x, y: 1.1, w: h.w, h: 0.45,
      fill: { color: colors.primary.dark }
    });
    slide.addText(h.text, {
      x: h.x, y: 1.18, w: h.w, h: 0.3,
      fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
  });

  focusPoints.forEach((fp, i) => {
    const y = 1.55 + i * 0.9;

    slide.addText(fp.point, {
      x: 0.5, y: y + 0.25, w: 0.8, h: 0.4,
      fontSize: 10, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(fp.title, {
      x: 1.3, y: y + 0.15, w: 2.2, h: 0.6,
      fontSize: 10, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(fp.plaintiff, {
      x: 3.5, y: y + 0.15, w: 2.0, h: 0.6,
      fontSize: 10, color: colors.semantic.success, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(fp.defendant, {
      x: 5.5, y: y + 0.15, w: 2.0, h: 0.6,
      fontSize: 10, color: colors.semantic.danger, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(fp.result, {
      x: 7.5, y: y + 0.25, w: 2.0, h: 0.4,
      fontSize: 10, bold: true, color: colors.accent.gold, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    if (i < focusPoints.length - 1) {
      slide.addShape(pres.ShapeType.line, {
        x: 0.5, y: y + 0.9, w: 9, h: 0,
        line: { color: colors.neutral[200], width: 0.5 }
      });
    }
  });

  return slide;
}

function createLegalBasisSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "法律依据",
    subtitle: "Legal Basis",
    pageNum: 4,
    totalPages: 16
  });

  const laws = data || [
    { type: "法律", name: "《中华人民共和国民法典》", articles: "第509条、第563条、第577条", content: "合同当事人应当按照约定全面履行自己的义务..." },
    { type: "法律", name: "《中华人民共和国公司法》", articles: "第71条、第73条", content: "有限责任公司的股东之间可以相互转让其全部或部分股权..." },
    { type: "司法解释", name: "《最高人民法院关于适用〈民法典〉合同编通则若干问题的解释》", articles: "第20条、第22条", content: "当事人一方违反合同义务的，另一方有权请求其承担违约责任..." },
    { type: "司法解释", name: "《全国法院民商事审判工作会议纪要》", articles: "第47条", content: "关于股权转让合同履行障碍的相关规定..." }
  ];

  laws.forEach((law, i) => {
    const y = 1.1 + i * 1.0;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.9,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    const typeColor = law.type === "法律" ? colors.primary.main : colors.primary.mid;
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.2, h: 0.9,
      fill: { color: typeColor }
    });

    slide.addText(law.type, {
      x: 0.5, y: y + 0.3, w: 1.2, h: 0.3,
      fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(law.name, {
      x: 1.85, y: y + 0.1, w: 7.4, h: 0.3,
      fontSize: 12, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(law.articles, {
      x: 1.85, y: y + 0.4, w: 7.4, h: 0.2,
      fontSize: 10, color: colors.accent.gold, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(law.content, {
      x: 1.85, y: y + 0.6, w: 7.4, h: 0.25,
      fontSize: 9, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  return slide;
}

function createEvidenceAnalysisSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "证据分析",
    subtitle: "Evidence Analysis",
    pageNum: 5,
    totalPages: 16
  });

  const evidenceGroups = data || [
    {
      category: "书证",
      items: [
        { name: "股权转让协议", weight: "证明力强", status: "✓" },
        { name: "工商变更登记材料", weight: "证明力强", status: "✓" },
        { name: "银行转账凭证", weight: "证明力强", status: "✓" }
      ]
    },
    {
      category: "函件",
      items: [
        { name: "律师函及邮寄凭证", weight: "证明力中", status: "✓" },
        { name: "往来微信记录", weight: "需公证", status: "!" }
      ]
    },
    {
      category: "当事人陈述",
      items: [
        { name: "原告陈述", weight: "待质证", status: "○" },
        { name: "被告陈述", weight: "证明力弱", status: "✗" }
      ]
    }
  ];

  evidenceGroups.forEach((group, i) => {
    const x = 0.5 + i * 3.1;

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 2.9, h: 4.1,
      fill: { color: colors.neutral[50] },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: x, y: 1.1, w: 2.9, h: 0.45,
      fill: { color: colors.primary.main }
    });

    slide.addText(group.category, {
      x: x, y: 1.18, w: 2.9, h: 0.3,
      fontSize: 13, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    group.items.forEach((item, j) => {
      const itemY = 1.7 + j * 1.1;

      slide.addShape(pres.ShapeType.rect, {
        x: x + 0.1, y: itemY, w: 2.7, h: 0.95,
        fill: { color: colors.white },
        line: { color: colors.neutral[200], width: 0.5 }
      });

      const statusColor = item.status === "✓" ? colors.semantic.success : item.status === "!" ? colors.accent.gold : colors.neutral[400];

      slide.addText(item.status, {
        x: x + 0.2, y: itemY + 0.1, w: 0.4, h: 0.4,
        fontSize: 16, color: statusColor, fontFace: "Microsoft YaHei", margin: 0
      });

      slide.addText(item.name, {
        x: x + 0.6, y: itemY + 0.15, w: 2.0, h: 0.35,
        fontSize: 11, bold: true, color: colors.neutral[700], fontFace: "Microsoft YaHei", margin: 0
      });

      slide.addText(item.weight, {
        x: x + 0.6, y: itemY + 0.5, w: 2.0, h: 0.3,
        fontSize: 10, color: colors.neutral[500], fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  return slide;
}

function createDamageCalcSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "损失计算",
    subtitle: "Damage Calculation",
    pageNum: 6,
    totalPages: 16
  });

  const calcItems = data || [
    { item: "到期未付本金（第二期）", amount: "4,000,000", days: "365天", rate: "年化15.4%", interest: "615,000" },
    { item: "到期未付本金（第三期）", amount: "2,000,000", days: "180天", rate: "年化15.4%", interest: "151,000" },
    { item: "违约金（协议约定）", amount: "8,500,000", days: "-", rate: "日万分之五", interest: "1,552,500" },
    { item: "律师费（预估）", amount: "-", days: "-", rate: "政府指导价", interest: "80,000" }
  ];

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 1.1, w: 9, h: 3.4,
    fill: { color: colors.white },
    line: { color: colors.neutral[200], width: brandVI.border.width }
  });

  const headers = ["项目", "本金(元)", "天数", "利率", "利息/金额(元)"];
  const widths = [2.5, 1.5, 1.0, 1.5, 2.5];
  let xPos = 0.5;

  headers.forEach((h, i) => {
    slide.addShape(pres.ShapeType.rect, {
      x: xPos, y: 1.1, w: widths[i], h: 0.5,
      fill: { color: colors.primary.dark }
    });
    slide.addText(h, {
      x: xPos, y: 1.2, w: widths[i], h: 0.3,
      fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
    xPos += widths[i];
  });

  calcItems.forEach((item, i) => {
    const y = 1.6 + i * 0.7;
    xPos = 0.5;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.7,
      fill: { color: i % 2 === 0 ? colors.neutral[50] : colors.white }
    });

    const values = [item.item, item.amount, item.days, item.rate, item.interest];
    values.forEach((v, j) => {
      slide.addText(v, {
        x: xPos, y: y + 0.2, w: widths[j], h: 0.3,
        fontSize: 11, color: j === 4 ? colors.accent.gold : colors.neutral[700],
        fontFace: "Microsoft YaHei", align: j === 0 ? "left" : "center", margin: 0
      });
      xPos += widths[j];
    });
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 4.55, w: 9, h: 0.7,
    fill: { color: colors.primary.dark }
  });

  slide.addText("诉请合计", {
    x: 0.7, y: 4.7, w: 2, h: 0.4,
    fontSize: 14, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText("¥ 6,398,500", {
    x: 5, y: 4.65, w: 4.3, h: 0.5,
    fontSize: 24, bold: true, color: colors.accent.gold, fontFace: "Microsoft YaHei", align: "right", margin: 0
  });

  addPageFooter(slide, pres, config);
  return slide;
}

function createSecondInstanceSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "二审策略",
    subtitle: "Second Instance Strategy",
    pageNum: 7,
    totalPages: 16
  });

  const strategies = data || [
    { title: "核心策略", content: "维持一审判决，争取二审全面胜诉", icon: "★", color: colors.primary.dark },
    { title: "证据策略", content: "补充提交被告恶意转移财产的相关证据，证明其存在逃避债务的主观故意", icon: "◆", color: colors.primary.main },
    { title: "程序策略", content: "申请法院调取被告银行流水，查明其实际财务状况和资金去向", icon: "●", color: colors.primary.mid },
    { title: "和解策略", content: "做好调解预案，如被告确有和解诚意，可适当减免利息以尽快实现债权", icon: "○", color: colors.neutral[500] }
  ];

  strategies.forEach((s, i) => {
    const y = 1.1 + i * 1.05;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 0.95,
      fill: { color: colors.neutral[50] },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 0.08, h: 0.95,
      fill: { color: s.color }
    });

    slide.addText(s.icon, {
      x: 0.75, y: y + 0.2, w: 0.5, h: 0.5,
      fontSize: 20, color: s.color, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(s.title, {
      x: 1.4, y: y + 0.15, w: 2, h: 0.35,
      fontSize: 14, bold: true, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(s.content, {
      x: 1.4, y: y + 0.5, w: 7.8, h: 0.4,
      fontSize: 12, color: colors.neutral[600], fontFace: "Microsoft YaHei", margin: 0
    });
  });

  return slide;
}

function createRiskWarningSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: "风险提示",
    subtitle: "Risk Warning",
    pageNum: 8,
    totalPages: 16
  });

  const risks = data || [
    { level: "高", title: "执行风险", content: "被告名下无可供执行财产，即使胜诉也可能面临执行困难", color: colors.semantic.danger },
    { level: "中", title: "时间风险", content: "二审程序预计需要6-12个月，债权实现周期较长", color: colors.accent.gold },
    { level: "低", title: "改判风险", content: "一审事实认定清楚，法律适用正确，二审改判可能性较小", color: colors.semantic.success }
  ];

  risks.forEach((risk, i) => {
    const y = 1.1 + i * 1.4;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 9, h: 1.25,
      fill: { color: colors.white },
      line: { color: colors.neutral[200], width: brandVI.border.width }
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.0, h: 1.25,
      fill: { color: risk.color }
    });

    slide.addText(risk.level, {
      x: 0.5, y: y + 0.45, w: 1.0, h: 0.4,
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

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fill: { color: colors.accent.gold, transparency: 20 }
  });

  slide.addText("建议：综合评估风险后确定代理方案，如风险较高可考虑诉前保全或和解策略", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.2,
    fontSize: 10, color: colors.primary.dark, fontFace: "Microsoft YaHei", margin: 0
  });

  return slide;
}

function createEndingSlide(pres, config, data) {
  const colors = brandVI.colors;

  let slide = pres.addSlide();
  slide.background = { color: colors.primary.deep };

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: colors.primary.dark, transparency: 40 }
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 4.5, y: 2.0, w: 0.1, h: 1.5,
    fill: { color: colors.accent.gold }
  });

  slide.addText("感谢聆听", {
    x: 4.8, y: 2.0, w: 4.5, h: 0.7,
    fontSize: 36, bold: true, color: colors.neutral[50], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText("THANK YOU", {
    x: 4.8, y: 2.7, w: 4.5, h: 0.5,
    fontSize: 16, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.line, {
    x: 4.8, y: 3.4, w: 4, h: 0,
    line: { color: colors.neutral[600], width: 0.5 }
  });

  slide.addText(brandVI.lawFirm.name, {
    x: 4.8, y: 3.6, w: 4.5, h: 0.35,
    fontSize: 14, color: colors.accent.goldLight, fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(brandVI.lawFirm.branch, {
    x: 4.8, y: 3.95, w: 4.5, h: 0.3,
    fontSize: 11, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(`TEL: ${brandVI.lawFirm.tel}`, {
    x: 4.8, y: 4.3, w: 4.5, h: 0.3,
    fontSize: 11, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
  });

  return slide;
}

function createFullTemplate(config = {}, data = {}) {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = config.title || "法律案件分析";
  pres.company = brandVI.lawFirm.name;
  pres.subject = config.title || "法律案件分析";
  pres.category = "Legal Presentation";

  createTitleSlide(pres, config);
  createBackgroundSlide(pres, config, data.background);
  createPartyInfoSlide(pres, config, data.parties);
  createDisputeFocusSlide(pres, config, data.disputes);
  createLegalBasisSlide(pres, config, data.laws);
  createEvidenceAnalysisSlide(pres, config, data.evidenceAnalysis);
  createDamageCalcSlide(pres, config, data.damageCalc);
  createTimelineSlide(pres, config, data.timeline);
  createSwotSlide(pres, config, data.swot);
  createLegalElementSlide(pres, config, data.elements);
  createAmountSlide(pres, config, data.amounts);
  createProgressSlide(pres, config, data.progress);
  createStrategySlide(pres, config, data.strategies);
  createSecondInstanceSlide(pres, config, data.secondInstance);
  createRiskWarningSlide(pres, config, data.risks);
  createDefenseSlide(pres, config, data.defense);
  createEvidenceListSlide(pres, config, data.evidence);
  createEndingSlide(pres, config);

  return pres;
}

module.exports = {
  createTitleSlide,
  createBackgroundSlide,
  createPartyInfoSlide,
  createDisputeFocusSlide,
  createLegalBasisSlide,
  createEvidenceAnalysisSlide,
  createDamageCalcSlide,
  createTimelineSlide,
  createSwotSlide,
  createLegalElementSlide,
  createAmountSlide,
  createProgressSlide,
  createStrategySlide,
  createSecondInstanceSlide,
  createRiskWarningSlide,
  createDefenseSlide,
  createEvidenceListSlide,
  createEndingSlide,
  createFullTemplate,
  brandVI,
  defaultConfig
};

if (require.main === module) {
  const config = {
    title: "股权转让纠纷案件分析",
    case: {
      caseNo: "(2024)湘01民终4567号",
      court: "长沙市中级人民法院",
      type: "股权转让纠纷"
    }
  };

  const caseData = {
    caseNo: "(2024)湘01民终4567号",
    type: "股权转让纠纷",
    court: "长沙市中级人民法院",
    judge: "李法官",
    plaintiff: "湖南星辉科技有限公司",
    defendant: "张三（原股东）",
    amount: "¥8,500,000",
    date: "2024年5月20日"
  };

  const timelineData = [
    { date: "2022-03-15", title: "股权转让协议", desc: "签订《股权转让协议》", status: "completed" },
    { date: "2022-03-20", title: "股权变更", desc: "完成工商变更登记", status: "completed" },
    { date: "2023-06-30", title: "付款逾期", desc: "第二期款项未支付", status: "completed" },
    { date: "2023-09-01", title: "律师函", desc: "发送催告函", status: "completed" },
    { date: "2024-01-15", title: "一审立案", desc: "岳麓区人民法院立案", status: "completed" },
    { date: "2024-04-20", title: "一审判决", desc: "判决支持原告诉请", status: "current" },
    { date: "2024-06-30", title: "二审开庭", desc: "长沙市中级人民法院", status: "pending" }
  ];

  const swotData = {
    strengths: ["协议原件完整有效", "工商变更登记齐全", "催告函送达证据充分"],
    weaknesses: ["转让价格约定较高", "被告经济状况不明"],
    opportunities: ["被告有其他投资", "可申请财产保全"],
    threats: ["被告下落不明", "执行难度较大"]
  };

  const legalElements = [
    { name: "协议成立", detail: "双方签字盖章，协议依法成立", status: "satisfied" },
    { name: "协议生效", detail: "符合生效条件，协议已生效", status: "satisfied" },
    { name: "股权变更", detail: "已完成工商变更登记手续", status: "satisfied" },
    { name: "付款义务", detail: "被告未按约支付第二期转让款", status: "satisfied" },
    { name: "催告有效", detail: "律师函已有效送达被告", status: "satisfied" },
    { name: "违约责任", detail: "协议约定逾期付款违约责任", status: "satisfied" },
    { name: "损害金额", detail: "原告损失金额待二审确认", status: "current" }
  ];

  const amountsData = [
    { title: "股权转让款", amount: "6,000,000", rate: "70.6%" },
    { title: "逾期利息", amount: "1,800,000", rate: "21.2%" },
    { title: "违约金", amount: "700,000", rate: "8.2%" },
    { title: "合计", amount: "8,500,000", rate: "100%", isTotal: true }
  ];

  const progressStages = [
    { name: "一审", status: "completed" },
    { name: "上诉", status: "completed" },
    { name: "立案", status: "completed" },
    { name: "举证", status: "current" },
    { name: "开庭", status: "pending" },
    { name: "判决", status: "pending" }
  ];

  const strategies = [
    { phase: "第一阶段", title: "财产保全", desc: "申请冻结被告持有的其他公司股权", priority: "high" },
    { phase: "第二阶段", title: "补充证据", desc: "收集被告其他财产线索", priority: "high" },
    { phase: "第三阶段", title: "和解谈判", desc: "在二审中争取调解结案", priority: "medium" },
    { phase: "第四阶段", title: "强制执行", desc: "判决后申请强制执行股权", priority: "medium" }
  ];

  const defenseData = {
    opponentClaim: "• 资金周转困难请求延期\n• 原告未履行协助义务\n• 协议存在显失公平",
    rebuttal: "• 协议系双方真实意思表示\n• 原告已全部履行合同义务\n• 被告违约事实清楚",
    conclusion: "一审法院全面支持原告主张，二审中被告上诉理由缺乏事实和法律依据，预计维持原判可能性较大。"
  };

  const evidenceData = [
    { title: "协议文件", items: ["股权转让协议原件", "补充协议", "股东会决议"] },
    { title: "变更登记", items: ["工商变更登记证明", "股东名册变更", "公司章程修正案"] },
    { title: "付款凭证", items: ["第一期付款凭证", "银行转账记录", "收款收据"] },
    { title: "催告函件", items: ["律师函", "邮寄凭证", "送达回证"] }
  ];

  const secondInstanceData = [
    { title: "核心策略", content: "维持一审判决，争取二审全面胜诉", icon: "★" },
    { title: "证据策略", content: "补充提交被告恶意转移财产的相关证据，证明其存在逃避债务的主观故意", icon: "◆" },
    { title: "程序策略", content: "申请法院调取被告银行流水，查明其实际财务状况和资金去向", icon: "●" },
    { title: "和解策略", content: "做好调解预案，如被告确有和解诚意，可适当减免利息以尽快实现债权", icon: "○" }
  ];

  const riskData = [
    { level: "高", title: "执行风险", content: "被告名下无可供执行财产，即使胜诉也可能面临执行困难" },
    { level: "中", title: "时间风险", content: "二审程序预计需要6-12个月，债权实现周期较长" },
    { level: "低", title: "改判风险", content: "一审事实认定清楚，法律适用正确，二审改判可能性较小" }
  ];

  const backgroundData = [
    { title: "股权转让背景", content: "2022年3月，被告张三将其持有的湖南星辉科技有限公司30%股权（对应注册资本300万元）转让给原告，双方签订《股权转让协议》约定转让款为1000万元，分三期支付。" },
    { title: "股权变更完成", content: "2022年3月20日，双方完成工商变更登记，原告正式成为持有公司30%股权的股东，并开始参与公司经营管理。" },
    { title: "违约事实发生", content: "截至2023年6月30日，被告仅支付第一期款项400万元，第二期款项400万元及第三期款项200万元均未按约支付。" },
    { title: "一审诉讼经过", content: "原告于2024年1月向岳麓区人民法院提起诉讼，一审法院判决支持原告全部诉讼请求，被告不服提起上诉。" }
  ];

  const partiesData = {
    plaintiff: {
      name: "湖南星辉科技有限公司",
      representative: "李明（法定代表人）",
      role: "上诉人（原审原告）",
      addr: "长沙市高新区...",
      background: "成立于2015年，注册资本1000万元，主要从事软件开发业务"
    },
    defendant: {
      name: "张三",
      idCard: "4301051978**********",
      role: "被上诉人（原审被告）",
      addr: "长沙市岳麓区...",
      background: "原星辉公司股东，持有30%股权，现已转让全部股权"
    }
  };

  const disputesData = [
    { point: "焦点一", title: "转让款支付条件是否成就", plaintiff: "被告未按约定时间支付第二、三期款项", defendant: "原告未协助办理资质变更手续", result: "被告违约" },
    { point: "焦点二", title: "违约金金额如何认定", plaintiff: "按协议约定日万分之五计算", defendant: "约定过高，请求法院调整", result: "需法院酌定" },
    { point: "焦点三", title: "利息计算的起止时间", plaintiff: "自各期应付之日起算", defendant: "应从催告后开始计算", result: "待二审认定" },
    { point: "焦点四", title: "被告主张的抗辩是否成立", plaintiff: "被告抗辩缺乏事实依据", defendant: "存在情势变更情形", result: "一审未采信" }
  ];

  const lawsData = [
    { type: "法律", name: "《中华人民共和国民法典》", articles: "第509条、第563条、第577条", content: "合同当事人应当按照约定全面履行自己的义务..." },
    { type: "法律", name: "《中华人民共和国公司法》", articles: "第71条、第73条", content: "有限责任公司的股东之间可以相互转让其全部或部分股权..." },
    { type: "司法解释", name: "《最高人民法院关于适用〈民法典〉合同编通则若干问题的解释》", articles: "第20条、第22条", content: "当事人一方违反合同义务的，另一方有权请求其承担违约责任..." },
    { type: "司法解释", name: "《全国法院民商事审判工作会议纪要》", articles: "第47条", content: "关于股权转让合同履行障碍的相关规定..." }
  ];

  const evidenceAnalysisData = [
    { category: "书证", items: [
      { name: "股权转让协议", weight: "证明力强", status: "✓" },
      { name: "工商变更登记材料", weight: "证明力强", status: "✓" },
      { name: "银行转账凭证", weight: "证明力强", status: "✓" }
    ]},
    { category: "函件", items: [
      { name: "律师函及邮寄凭证", weight: "证明力中", status: "✓" },
      { name: "往来微信记录", weight: "需公证", status: "!" }
    ]},
    { category: "当事人陈述", items: [
      { name: "原告陈述", weight: "待质证", status: "○" },
      { name: "被告陈述", weight: "证明力弱", status: "✗" }
    ]}
  ];

  const damageCalcData = [
    { item: "到期未付本金（第二期）", amount: "4,000,000", days: "365天", rate: "年化15.4%", interest: "615,000" },
    { item: "到期未付本金（第三期）", amount: "2,000,000", days: "180天", rate: "年化15.4%", interest: "151,000" },
    { item: "违约金（协议约定）", amount: "8,500,000", days: "-", rate: "日万分之五", interest: "1,552,500" },
    { item: "律师费（预估）", amount: "-", days: "-", rate: "政府指导价", interest: "80,000" }
  ];

  const pres = createFullTemplate(config, {
    background: backgroundData,
    parties: partiesData,
    disputes: disputesData,
    laws: lawsData,
    evidenceAnalysis: evidenceAnalysisData,
    damageCalc: damageCalcData,
    timeline: timelineData,
    swot: swotData,
    elements: legalElements,
    amounts: amountsData,
    progress: progressStages,
    strategies: strategies,
    secondInstance: secondInstanceData,
    risks: riskData,
    defense: defenseData,
    evidence: evidenceData
  });
  pres.writeFile({ fileName: "股权转让纠纷案件分析_完整版.pptx" })
    .then(() => console.log("PPT created: 股权转让纠纷案件分析_完整版.pptx"))
    .catch(err => console.error(err));
}
