const PptxGenJS = require("pptxgenjs");
const brandVI = require("./brand_vi");
const master = require("./master_slide");

function createLectureTitleSlide(pres, config, data) {
  const slide = pres.addSlide();
  const lectureData = data || {
    title: "企业合同风险管理",
    subtitle: "从签订到履行的全流程法律实务",
    speaker: "湖南金厚（宁乡）律师事务所",
    date: "2024年6月",
    topic: "专题讲座"
  };

  master.coverSlide.apply(slide, pres, {
    style: 'default',
    decoration: 'lecture',
    accentBar: { x: 0, y: 2.2, h: 1.5 },
    type: lectureData.topic,
    title: lectureData.title,
    subtitle: lectureData.subtitle,
    titleY: 2.2,
    titleSize: 44,
    meta: { y: 4.0, text: `${lectureData.speaker}  |  ${lectureData.date}`, w: 3.5 }
  });
}

function createLectureAgendaSlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();
  const agendaItems = data || [
    { num: "01", title: "合同风险概述", desc: "企业合同管理的重要性" },
    { num: "02", title: "签订阶段风险", desc: "主体审查、条款设计" },
    { num: "03", title: "履行阶段风险", desc: "变更、转让、解除" },
    { num: "04", title: "纠纷解决路径", desc: "协商、调解、仲裁、诉讼" }
  ];

  master.contentSlide.apply(slide, pres, {
    title: "讲座大纲",
    subtitle: "Agenda",
    pageNum: pageNum,
    totalPages: 12
  });

  const colors = brandVI.colors;
  const masterContent = master.contentSlide;

  agendaItems.forEach((item, i) => {
    const x = 0.5 + (i % 2) * 4.6;
    const y = 1.2 + Math.floor(i / 2) * 2.0;

    masterContent.card(slide, pres, x, y, 4.3, 1.8, {
      fill: colors.white,
      borderColor: colors.neutral[200]
    });

    masterContent.card(slide, pres, x, y, 0.9, 1.8, {
      fill: colors.primary.main
    });

    slide.addText(item.num, {
      x: x, y: y + 0.5, w: 0.9, h: 0.5,
      fontSize: 26, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    masterContent.sectionTitle(slide, pres, x + 1.1, y + 0.4, item.title, colors.primary.dark);
    masterContent.bodyText(slide, pres, x + 1.1, y + 1.0, item.desc, 3.0);

    slide.addShape(pres.ShapeType.rect, {
      x: x + 0.05, y: y + 0.9, w: 0.8, h: 0.02,
      fill: { color: colors.accent.gold }
    });
  });

  for (let i = 0; i < 3; i++) {
    slide.addShape(pres.ShapeType.rect, {
      x: 9.2 + (i % 2) * 0.2, y: 1.5 + i * 0.5,
      w: 0.5, h: 0.4,
      fill: { color: i % 2 === 0 ? colors.accent.gold : colors.primary.mid, transparency: 40 + i * 15 }
    });
  }
}

function createLectureContentSlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();

  master.contentSlide.apply(slide, pres, {
    title: data?.title || "内容标题",
    subtitle: data?.subtitle || "",
    pageNum: pageNum,
    totalPages: 12
  });

  const colors = brandVI.colors;
  const headerColors = [colors.primary.main, colors.primary.mid, colors.primary.dark];
  const contents = data?.contents || [
    { title: "合同主体风险", points: ["公章管理不当", "授权不明", "冒名签订"] },
    { title: "合同条款风险", points: ["霸王条款", "重大误解", "显失公平"] },
    { title: "合同履行风险", points: ["迟延履行", "瑕疵履行", "拒绝履行"] }
  ];

  contents.forEach((section, i) => {
    const x = 0.5 + i * 3.1;
    const masterContent = master.contentSlide;

    masterContent.card(slide, pres, x, 1.1, 2.9, 4.1, {
      fill: colors.white,
      borderColor: colors.neutral[200]
    });

    masterContent.card(slide, pres, x, 1.1, 2.9, 0.55, {
      fill: headerColors[i]
    });

    slide.addText(section.title, {
      x: x, y: 1.2, w: 2.9, h: 0.4,
      fontSize: 14, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    section.points.forEach((point, j) => {
      const pointY = 1.85 + j * 1.1;

      slide.addShape(pres.ShapeType.rect, {
        x: x + 0.15, y: pointY, w: 2.6, h: 0.95,
        fill: { color: colors.neutral[50] }
      });

      masterContent.accentBar(slide, pres, x + 0.15, pointY, 0.95, colors.accent.gold);

      slide.addText(`${j + 1}`, {
        x: x + 0.35, y: pointY + 0.2, w: 0.4, h: 0.4,
        fontSize: 18, bold: true, color: colors.primary.main,
        fontFace: "Microsoft YaHei", margin: 0
      });

      slide.addText(point, {
        x: x + 0.85, y: pointY + 0.25, w: 1.8, h: 0.5,
        fontSize: 13, color: colors.neutral[700],
        fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });

  for (let i = 0; i < 3; i++) {
    slide.addShape(pres.ShapeType.rect, {
      x: 9.2 + (i % 2) * 0.2, y: 1.5 + i * 0.5,
      w: 0.5, h: 0.4,
      fill: { color: i % 2 === 0 ? colors.accent.gold : colors.primary.mid, transparency: 40 + i * 15 }
    });
  }
}

function createLectureCaseSlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();
  const colors = brandVI.colors;

  master.contentSlide.apply(slide, pres, {
    title: "案例分析",
    subtitle: "Case Study",
    pageNum: pageNum,
    totalPages: 12
  });

  const caseData = data || {
    title: "某设备采购合同纠纷案",
    parties: ["原告：甲公司", "被告：乙公司"],
    dispute: "设备验收合格后，乙方拖欠货款800万元",
    points: [
      { title: "争议焦点", content: "设备是否验收合格" },
      { title: "原告主张", content: "设备已正常运行超过30日" },
      { title: "被告抗辩", content: "设备存在质量问题" },
      { title: "法院认定", content: "支持原告，判决被告支付全款" }
    ]
  };

  master.contentSlide.card(slide, pres, 0.5, 1.1, 9, 1.1, { fill: colors.primary.dark });
  master.contentSlide.accentBar(slide, pres, 0.5, 1.1, 1.1, colors.accent.gold);

  slide.addText(caseData.title, {
    x: 0.8, y: 1.2, w: 8.5, h: 0.4,
    fontSize: 20, bold: true, color: colors.white,
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(caseData.parties.join("    |    "), {
    x: 0.8, y: 1.7, w: 8.5, h: 0.3,
    fontSize: 11, color: colors.neutral[300],
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.3, w: 9, h: 0.55,
    fill: { color: colors.accent.gold, transparency: 20 }
  });

  slide.addText("核心争议：" + caseData.dispute, {
    x: 0.7, y: 2.4, w: 8.6, h: 0.35,
    fontSize: 14, bold: true, color: colors.primary.dark,
    fontFace: "Microsoft YaHei", margin: 0
  });

  caseData.points.forEach((p, i) => {
    const y = 3.0 + i * 0.6;
    const pointColor = i === 3 ? colors.semantic.success : colors.primary.main;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: y, w: 1.6, h: 0.5,
      fill: { color: pointColor }
    });

    slide.addText(p.title, {
      x: 0.5, y: y + 0.1, w: 1.6, h: 0.3,
      fontSize: 11, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(p.content, {
      x: 2.3, y: y + 0.1, w: 7.0, h: 0.3,
      fontSize: 12, color: colors.neutral[700],
      fontFace: "Microsoft YaHei", margin: 0
    });
  });
}

function createDueDiligenceTitleSlide(pres, config, data) {
  const slide = pres.addSlide();
  const ddData = data || {
    title: "股权收购尽职调查报告",
    subtitle: "湖南星辉科技有限公司",
    type: "尽职调查",
    date: "2024年5月"
  };

  master.coverSlide.apply(slide, pres, {
    style: 'split',
    decoration: 'dd',
    accentBar: { x: 0, y: 1.8, h: 1.2 },
    type: ddData.type,
    title: ddData.title,
    subtitle: ddData.subtitle,
    titleY: 1.8,
    titleSize: 32,
    titleW: 5,
    meta: { y: 3.5, text: ddData.date, w: 2.5 }
  });
}

function createDueDiligenceSummarySlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();
  const colors = brandVI.colors;

  master.contentSlide.apply(slide, pres, {
    title: "尽调概览",
    subtitle: "Due Diligence",
    pageNum: pageNum,
    totalPages: 8
  });

  const summaryData = data || {
    company: "湖南星辉科技有限公司",
    target: "收购30%股权",
    amount: "人民币1000万元",
    conclusion: "建议开展正式尽调",
    items: [
      { area: "公司基本情况", status: "通过", color: colors.semantic.success },
      { area: "股权结构", status: "通过", color: colors.semantic.success },
      { area: "法律合规", status: "通过", color: colors.semantic.success },
      { area: "经营状况", status: "待核实", color: colors.accent.gold }
    ],
    chartData: [
      { label: "通过", value: 75 },
      { label: "待核实", value: 15 },
      { label: "问题", value: 10 }
    ]
  };

  master.contentSlide.card(slide, pres, 0.5, 1.1, 5.5, 4.0, {
    fill: colors.white,
    borderColor: colors.neutral[200]
  });

  slide.addText(summaryData.company, {
    x: 0.7, y: 1.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: colors.primary.dark,
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText(`收购标的：${summaryData.target}  |  交易金额：${summaryData.amount}`, {
    x: 0.7, y: 1.8, w: 5, h: 0.3,
    fontSize: 11, color: colors.neutral[500],
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addText("调查结论：" + summaryData.conclusion, {
    x: 0.7, y: 2.3, w: 5, h: 0.3,
    fontSize: 13, bold: true, color: colors.accent.gold,
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.7, y: 2.8, w: 5, h: 0.35,
    fill: { color: colors.primary.main }
  });

  slide.addText("调查项目", { x: 0.7, y: 2.85, w: 2.5, h: 0.25, fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", margin: 0 });
  slide.addText("状态", { x: 3.5, y: 2.85, w: 2, h: 0.25, fontSize: 11, bold: true, color: colors.white, fontFace: "Microsoft YaHei", align: "center", margin: 0 });

  summaryData.items.forEach((item, i) => {
    const y = 3.15 + i * 0.45;

    slide.addShape(pres.ShapeType.rect, {
      x: 0.7, y: y, w: 5, h: 0.45,
      fill: { color: i % 2 === 0 ? colors.neutral[50] : colors.white }
    });

    slide.addText(item.area, {
      x: 0.9, y: y + 0.08, w: 2.5, h: 0.3,
      fontSize: 12, color: colors.neutral[700],
      fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 4.0, y: y + 0.08, w: 1.2, h: 0.3,
      fill: { color: item.color }
    });

    slide.addText(item.status, {
      x: 4.0, y: y + 0.08, w: 1.2, h: 0.3,
      fontSize: 11, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });
  });

  master.contentSlide.card(slide, pres, 6.2, 1.1, 3.3, 4.0, {
    fill: colors.white,
    borderColor: colors.neutral[200]
  });

  slide.addText("调查进度", {
    x: 6.2, y: 1.2, w: 3.3, h: 0.35,
    fontSize: 12, bold: true, color: colors.primary.dark,
    fontFace: "Microsoft YaHei", align: "center", margin: 0
  });

  slide.addChart(pres.charts.doughnut, [{
    name: "占比",
    labels: summaryData.chartData.map(d => d.label),
    values: summaryData.chartData.map(d => d.value)
  }], {
    x: 6.4, y: 1.7, w: 2.9, h: 2.9,
    chartColors: [colors.semantic.success, colors.accent.gold, colors.semantic.danger, colors.primary.mid],
    showPercent: true,
    showLegend: false,
    holeSize: 60
  });

  slide.addText("75%", {
    x: 6.2, y: 4.5, w: 3.3, h: 0.4,
    fontSize: 28, bold: true, color: colors.primary.dark,
    fontFace: "Microsoft YaHei", align: "center", margin: 0
  });
}

function createContractReviewTitleSlide(pres, config, data) {
  const slide = pres.addSlide();
  const crData = data || {
    title: "合同审查要点清单",
    subtitle: "常用合同条款风险识别与修订",
    type: "实务指南",
    date: "2024年"
  };

  master.coverSlide.apply(slide, pres, {
    style: 'default',
    decoration: 'cr',
    accentBar: { x: 0, y: 2.0, h: 1.8 },
    type: crData.type,
    title: crData.title,
    subtitle: crData.subtitle,
    titleY: 2.0,
    titleSize: 40,
    meta: { y: 4.3, text: crData.date, w: 3 }
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: brandVI.colors.accent.gold }
  });
}

function createContractReviewListSlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();
  const colors = brandVI.colors;

  master.contentSlide.apply(slide, pres, {
    title: "合同审查清单",
    subtitle: "Review Checklist",
    pageNum: pageNum,
    totalPages: 6
  });

  const reviewItems = data || [
    { category: "主体审查", items: [
      { name: "营业执照", checked: false },
      { name: "法定代表人", checked: false },
      { name: "授权委托书", checked: false },
      { name: "资信情况", checked: false }
    ]},
    { category: "标的条款", items: [
      { name: "标的描述", checked: false },
      { name: "数量规格", checked: false },
      { name: "质量标准", checked: false },
      { name: "交付方式", checked: false }
    ]},
    { category: "价款条款", items: [
      { name: "金额约定", checked: false },
      { name: "支付方式", checked: false },
      { name: "支付时间", checked: false },
      { name: "发票开具", checked: false }
    ]},
    { category: "违约责任", items: [
      { name: "违约金比例", checked: false },
      { name: "损失赔偿", checked: false },
      { name: "免责条款", checked: false },
      { name: "不可抗力", checked: false }
    ]}
  ];

  reviewItems.forEach((cat, i) => {
    const x = 0.5 + (i % 2) * 4.6;
    const y = 1.1 + Math.floor(i / 2) * 2.2;
    const headerColor = i % 2 === 0 ? colors.primary.main : colors.primary.mid;

    master.contentSlide.card(slide, pres, x, y, 4.3, 2.0, {
      fill: colors.white,
      borderColor: colors.neutral[200]
    });

    master.contentSlide.card(slide, pres, x, y, 4.3, 0.45, {
      fill: headerColor
    });

    slide.addText(cat.category, {
      x: x, y: y + 0.08, w: 4.3, h: 0.3,
      fontSize: 13, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    cat.items.forEach((item, j) => {
      const itemX = x + 0.2 + (j % 2) * 2.0;
      const itemY = y + 0.6 + Math.floor(j / 2) * 0.65;

      slide.addShape(pres.ShapeType.rect, {
        x: itemX, y: itemY, w: 0.35, h: 0.35,
        fill: { color: colors.white },
        line: { color: colors.primary.mid, width: 1.5 }
      });

      if (item.checked) {
        slide.addText("✓", {
          x: itemX, y: itemY - 0.05, w: 0.35, h: 0.35,
          fontSize: 14, bold: true, color: colors.semantic.success,
          fontFace: "Microsoft YaHei", align: "center", margin: 0
        });
      }

      slide.addText(item.name, {
        x: itemX + 0.45, y: itemY, w: 1.5, h: 0.35,
        fontSize: 11, color: colors.neutral[700],
        fontFace: "Microsoft YaHei", margin: 0
      });
    });
  });
}

function createCaseReportTitleSlide(pres, config, data) {
  const slide = pres.addSlide();
  const crData = data || {
    title: "建设工程施工合同纠纷案",
    caseNo: "(2024)湘01民初789号",
    type: "案件汇报",
    court: "长沙市中级人民法院",
    date: "2024年6月"
  };

  master.coverSlide.apply(slide, pres, {
    style: 'bottomAccent',
    accentBar: { x: 0, y: 1.6, h: 2.2 },
    type: crData.type,
    title: crData.title,
    subtitle: `案号：${crData.caseNo}`,
    titleY: 1.6,
    titleSize: 36,
    subY: 2.6,
    meta: { y: 3.2, text: `法院：${crData.court}  |  日期：${crData.date}`, w: 5 }
  });
}

function createCaseReportOverviewSlide(pres, config, data, pageNum) {
  const slide = pres.addSlide();
  const colors = brandVI.colors;

  master.contentSlide.apply(slide, pres, {
    title: "案件概要",
    subtitle: "Case Overview",
    pageNum: pageNum,
    totalPages: 10
  });

  const overviewData = data || {
    project: "长沙市某商业综合体项目",
    contractAmount: "¥5,800万元",
    disputedAmount: "¥2,100万元",
    projectProgress: "85%",
    issues: [
      { title: "工程款结算争议", desc: "甲方拖延结算审核", level: "高" },
      { title: "工期延误责任", desc: "双方互相指责", level: "中" },
      { title: "质量问题整改", desc: "整改费用分担", level: "低" }
    ]
  };

  master.contentSlide.card(slide, pres, 0.5, 1.1, 9, 0.9, {
    fill: colors.primary.dark
  });

  slide.addText("项目：" + overviewData.project, {
    x: 0.7, y: 1.2, w: 8.6, h: 0.35,
    fontSize: 16, bold: true, color: colors.white,
    fontFace: "Microsoft YaHei", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 0.7, y: 1.7, w: 2.0, h: 0.25,
    fill: { color: colors.semantic.success }
  });
  slide.addText(overviewData.contractAmount, {
    x: 0.7, y: 1.7, w: 2.0, h: 0.25,
    fontSize: 10, bold: true, color: colors.white,
    fontFace: "Microsoft YaHei", align: "center", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 2.9, y: 1.7, w: 2.0, h: 0.25,
    fill: { color: colors.accent.gold }
  });
  slide.addText(overviewData.disputedAmount, {
    x: 2.9, y: 1.7, w: 2.0, h: 0.25,
    fontSize: 10, bold: true, color: colors.white,
    fontFace: "Microsoft YaHei", align: "center", margin: 0
  });

  slide.addShape(pres.ShapeType.rect, {
    x: 5.1, y: 1.7, w: 1.5, h: 0.25,
    fill: { color: colors.primary.mid }
  });
  slide.addText(overviewData.projectProgress, {
    x: 5.1, y: 1.7, w: 1.5, h: 0.25,
    fontSize: 10, bold: true, color: colors.white,
    fontFace: "Microsoft YaHei", align: "center", margin: 0
  });

  master.contentSlide.sectionTitle(slide, pres, 0.5, 2.2, "主要争议事项", colors.primary.dark);

  overviewData.issues.forEach((issue, i) => {
    const y = 2.7 + i * 0.9;
    const levelColor = issue.level === "高" ? colors.semantic.danger : issue.level === "中" ? colors.accent.gold : colors.semantic.success;

    master.contentSlide.card(slide, pres, 0.5, y, 9, 0.8, {
      fill: colors.white,
      borderColor: colors.neutral[200]
    });

    master.contentSlide.accentBar(slide, pres, 0.5, y, 0.8, levelColor);

    slide.addShape(pres.ShapeType.rect, {
      x: 0.75, y: y + 0.2, w: 0.5, h: 0.4,
      fill: { color: levelColor }
    });
    slide.addText(issue.level, {
      x: 0.75, y: y + 0.2, w: 0.5, h: 0.4,
      fontSize: 10, bold: true, color: colors.white,
      fontFace: "Microsoft YaHei", align: "center", margin: 0
    });

    slide.addText(issue.title, {
      x: 1.4, y: y + 0.15, w: 3, h: 0.3,
      fontSize: 14, bold: true, color: colors.primary.dark,
      fontFace: "Microsoft YaHei", margin: 0
    });

    slide.addText(issue.desc, {
      x: 1.4, y: y + 0.45, w: 7.8, h: 0.3,
      fontSize: 12, color: colors.neutral[600],
      fontFace: "Microsoft YaHei", margin: 0
    });
  });
}

function createEndingSlide(pres, config) {
  const slide = pres.addSlide();
  master.endingSlide.apply(slide, pres);
}

function createLectureTemplate(config = {}, data = {}) {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = config.title || "法律专题讲座";
  pres.company = brandVI.lawFirm.name;
  pres.category = "Legal Lecture";

  createLectureTitleSlide(pres, config, data.title);
  createLectureAgendaSlide(pres, config, data.agenda, 2);
  createLectureContentSlide(pres, config, data.content1, 3);
  createLectureContentSlide(pres, config, data.content2, 4);
  createLectureCaseSlide(pres, config, data.case1, 5);
  createLectureContentSlide(pres, config, data.content3, 6);
  createLectureCaseSlide(pres, config, data.case2, 7);
  createLectureContentSlide(pres, config, data.content4, 8);
  createLectureCaseSlide(pres, config, data.case3, 9);
  createLectureContentSlide(pres, config, data.content5, 10);
  createLectureContentSlide(pres, config, data.content6, 11);
  createEndingSlide(pres, config);

  return pres;
}

function createDueDiligenceTemplate(config = {}, data = {}) {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = config.title || "尽职调查报告";
  pres.company = brandVI.lawFirm.name;
  pres.category = "Due Diligence";

  createDueDiligenceTitleSlide(pres, config, data.title);
  createDueDiligenceSummarySlide(pres, config, data.summary, 2);
  createDueDiligenceSummarySlide(pres, config, data.companyInfo, 3);
  createDueDiligenceSummarySlide(pres, config, data.financial, 4);
  createDueDiligenceSummarySlide(pres, config, data.legal, 5);
  createDueDiligenceSummarySlide(pres, config, data.risks, 6);
  createDueDiligenceSummarySlide(pres, config, data.recommendation, 7);
  createEndingSlide(pres, config);

  return pres;
}

function createContractReviewTemplate(config = {}, data = {}) {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = config.title || "合同审查清单";
  pres.company = brandVI.lawFirm.name;
  pres.category = "Contract Review";

  createContractReviewTitleSlide(pres, config, data.title);

  const defaultReviewItems = [
    { category: "主体审查", items: [
      { name: "营业执照", checked: false },
      { name: "法定代表人", checked: false },
      { name: "授权委托书", checked: false },
      { name: "资信情况", checked: false }
    ]},
    { category: "标的条款", items: [
      { name: "标的描述", checked: false },
      { name: "数量规格", checked: false },
      { name: "质量标准", checked: false },
      { name: "交付方式", checked: false }
    ]},
    { category: "价款条款", items: [
      { name: "金额约定", checked: false },
      { name: "支付方式", checked: false },
      { name: "支付时间", checked: false },
      { name: "发票开具", checked: false }
    ]},
    { category: "违约责任", items: [
      { name: "违约金比例", checked: false },
      { name: "损失赔偿", checked: false },
      { name: "免责条款", checked: false },
      { name: "不可抗力", checked: false }
    ]}
  ];

  createContractReviewListSlide(pres, config, defaultReviewItems, 2);
  createContractReviewListSlide(pres, config, defaultReviewItems, 3);
  createContractReviewListSlide(pres, config, defaultReviewItems, 4);
  createContractReviewListSlide(pres, config, defaultReviewItems, 5);
  createEndingSlide(pres, config);

  return pres;
}

function createCaseReportTemplate(config = {}, data = {}) {
  let pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.author = "法律智囊";
  pres.title = config.title || "案件汇报";
  pres.company = brandVI.lawFirm.name;
  pres.category = "Case Report";

  createCaseReportTitleSlide(pres, config, data.title);
  createCaseReportOverviewSlide(pres, config, data.overview, 2);
  createCaseReportOverviewSlide(pres, config, data.claims, 3);
  createCaseReportOverviewSlide(pres, config, data.evidence, 4);
  createCaseReportOverviewSlide(pres, config, data.defense, 5);
  createCaseReportOverviewSlide(pres, config, data.strategy, 6);
  createCaseReportOverviewSlide(pres, config, data.analysis, 7);
  createCaseReportOverviewSlide(pres, config, data.opinion, 8);
  createCaseReportOverviewSlide(pres, config, data.risks, 9);
  createEndingSlide(pres, config);

  return pres;
}

if (require.main === module) {
  const lecturePres = createLectureTemplate(
    { title: "企业合同风险管理专题讲座" },
    {
      title: {
        title: "企业合同风险管理",
        subtitle: "从签订到履行的全流程法律实务",
        topic: "专题讲座",
        speaker: "湖南金厚（宁乡）律师事务所",
        date: "2024年6月"
      },
      agenda: [
        { num: "01", title: "合同风险概述", desc: "企业合同管理的重要性与风险识别" },
        { num: "02", title: "签订阶段风险", desc: "主体审查、条款设计与风险防控" },
        { num: "03", title: "履行阶段风险", desc: "变更、转让、解除中的法律风险" },
        { num: "04", title: "纠纷解决路径", desc: "协商、调解、仲裁、诉讼的选择" }
      ],
      content1: {
        title: "合同风险概述",
        subtitle: "Risk Overview",
        contents: [
          { title: "合同风险的危害", points: ["经济损失", "声誉损害", "经营中断"] },
          { title: "风险的成因", points: ["法律意识薄弱", "管理制度缺失", "人员能力不足"] },
          { title: "风险管理原则", points: ["预防为主", "全程管控", "及时应对"] }
        ]
      },
      case1: {
        title: "某采购合同纠纷案",
        parties: ["原告：某设备公司", "被告：某建筑公司"],
        dispute: "设备验收合格后，建筑公司拖欠货款1200万元",
        points: [
          { title: "争议焦点", content: "设备是否达到验收标准" },
          { title: "原告主张", content: "设备已正常运行超过合同约定期限" },
          { title: "被告抗辩", content: "设备存在隐蔽瑕疵" },
          { title: "法院认定", content: "支持原告，请求被告支付全款及违约金" }
        ]
      }
    }
  );

  lecturePres.writeFile({ fileName: "法律专题讲座模板.pptx" })
    .then(() => console.log("PPT created: 法律专题讲座模板.pptx"))
    .catch(err => console.error(err));

  const ddPres = createDueDiligenceTemplate(
    { title: "股权收购尽职调查报告" },
    {
      title: {
        title: "股权收购尽职调查报告",
        subtitle: "湖南星辉科技有限公司",
        type: "尽职调查",
        date: "2024年5月"
      },
      summary: {
        company: "湖南星辉科技有限公司",
        target: "收购30%股权",
        amount: "人民币1000万元",
        conclusion: "建议开展正式尽调",
        items: [
          { area: "公司基本情况", status: "通过", color: brandVI.colors.semantic.success },
          { area: "股权结构", status: "通过", color: brandVI.colors.semantic.success },
          { area: "法律合规", status: "通过", color: brandVI.colors.semantic.success },
          { area: "经营状况", status: "待核实", color: brandVI.colors.accent.gold }
        ],
        chartData: [
          { label: "通过", value: 75 },
          { label: "待核实", value: 15 },
          { label: "问题", value: 10 }
        ]
      }
    }
  );

  ddPres.writeFile({ fileName: "尽职调查报告模板.pptx" })
    .then(() => console.log("PPT created: 尽职调查报告模板.pptx"))
    .catch(err => console.error(err));

  const crPres = createContractReviewTemplate(
    { title: "合同审查要点清单" },
    {
      title: {
        title: "合同审查要点清单",
        subtitle: "常用合同条款风险识别与修订",
        type: "实务指南",
        date: "2024年"
      }
    }
  );

  crPres.writeFile({ fileName: "合同审查清单模板.pptx" })
    .then(() => console.log("PPT created: 合同审查清单模板.pptx"))
    .catch(err => console.error(err));

  const casePres = createCaseReportTemplate(
    { title: "建设工程施工合同纠纷案" },
    {
      title: {
        title: "建设工程施工合同纠纷案",
        caseNo: "(2024)湘01民初789号",
        type: "案件汇报",
        court: "长沙市中级人民法院",
        date: "2024年6月"
      },
      overview: {
        project: "长沙市某商业综合体项目",
        contractAmount: "¥5,800万元",
        disputedAmount: "¥2,100万元",
        projectProgress: "85%",
        issues: [
          { title: "工程款结算争议", desc: "甲方拖延结算审核", level: "高" },
          { title: "工期延误责任", desc: "双方互相指责对方责任", level: "中" },
          { title: "质量问题整改", desc: "整改费用分担存在分歧", level: "低" }
        ]
      }
    }
  );

  casePres.writeFile({ fileName: "案件汇报模板.pptx" })
    .then(() => console.log("PPT created: 案件汇报模板.pptx"))
    .catch(err => console.error(err));
}

module.exports = {
  createLectureTemplate,
  createDueDiligenceTemplate,
  createContractReviewTemplate,
  createCaseReportTemplate,
  createLectureTitleSlide,
  createLectureAgendaSlide,
  createLectureContentSlide,
  createLectureCaseSlide,
  createDueDiligenceTitleSlide,
  createDueDiligenceSummarySlide,
  createContractReviewTitleSlide,
  createContractReviewListSlide,
  createCaseReportTitleSlide,
  createCaseReportOverviewSlide,
  createEndingSlide,
  brandVI,
  master
};