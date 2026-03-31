const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak, HeadingLevel } = require("docx");
const fs = require("fs");
const path = require("path");

const CONTENT_WIDTH = 9360;

function parseElementFile(filePath) {
  const content = fs.readFileSync(filePath, "utf-8");
  const elements = {};
  const sections = content.split(/--/).filter(s => s.trim());

  sections.forEach(section => {
    const lines = section.trim().split("\n");
    let currentSection = "";
    lines.forEach(line => {
      const [key, ...valueParts] = line.split("::");
      const value = valueParts.join("::").trim();
      if (!key.includes(":") && key.trim()) {
        currentSection = key.trim();
        if (!elements[currentSection]) elements[currentSection] = {};
      } else if (key && value) {
        if (!elements[currentSection]) elements[currentSection] = {};
        elements[currentSection][key.trim()] = value;
      } else if (value) {
        elements[key.trim()] = value;
      }
    });
  });

  return elements;
}

function getVal(elements, key) {
  for (const section of Object.values(elements)) {
    if (section && section[key]) return section[key];
  }
  return null;
}

function fillTemplate(elements) {
  const border = {
    top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" }
  };

  const caseData = {
    bianh: getVal(elements, "bianh") || "",
    leib: getVal(elements, "leib") || "民事",
    anyou: getVal(elements, "anyou") || "",
    weitr: getVal(elements, "weitr") || "",
    dfdsr: getVal(elements, "dfdsr") || "",
    badw: getVal(elements, "badw") || "",
    jied: getVal(elements, "jied") || "一审",
    zprq: getVal(elements, "zprq") || "",
    wtrsfz: getVal(elements, "wtrsfz") || "",
    wtrdh: getVal(elements, "wtrdh") || "",
    wtrxb: getVal(elements, "wtrxb") || "",
    wtrcs: getVal(elements, "wtrcs") || "",
    wtrzz: getVal(elements, "wtrzz") || "",
    dcdw: getVal(elements, "dcdw") || "",
    bdqxx: getVal(elements, "bdqxx") || "",
    thsj1: getVal(elements, "thsj1") || "",
    thdd1: getVal(elements, "thdd1") || "",
    ajqk: getVal(elements, "ajqk") || "",
    ssqq: getVal(elements, "ssqq") || "",
    basl: getVal(elements, "basl") || "",
    ktsj: getVal(elements, "ktsj") || "",
    ktdd: getVal(elements, "ktdd") || "",
    cbfg: getVal(elements, "cbfg") || "",
    sjy: getVal(elements, "sjy") || "",
    yjrq: getVal(elements, "yjrq") || "",
    cbxj: getVal(elements, "cbxj") || "",
    jarq: getVal(elements, "jarq") || "",
    ljsm: getVal(elements, "ljsm") || "",
    gdrq: getVal(elements, "gdrq") || ""
  };

  const gclist = [];
  for (let i = 1; i <= 9; i++) {
    const sj = getVal(elements, `gcsj${i}`);
    const fs = getVal(elements, `gcfs${i}`);
    const nr = getVal(elements, `gcnr${i}`);
    if (sj || nr) {
      gclist.push({ sj: sj || "", fs: fs || "", nr: nr || "" });
    }
  }

  while (gclist.length < 9) {
    gclist.push({ sj: "", fs: "", nr: "" });
  }

  const processTable = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [1000, 2000, 2000, 5360],
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1000, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "序号", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "时间", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "方式", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 5360, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "主要内容", bold: true, font: "宋体", size: 20 })] })] })
        ]
      }),
      ...gclist.map((gc, i) => new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1000, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: String(i + 1), font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2000, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: gc.sj, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2000, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: gc.fs, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 5360, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: gc.nr, font: "宋体", size: 20 })] })] })
        ]
      }))
    ]
  });

  const reportTable = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [1500, 2500, 1500, 3860],
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "案件编号", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.bianh, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 3860, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "承办机构", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "湖南金厚（宁乡）律师事务所", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "承 办 人", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 3860, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "陈伟", font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "受 援 人", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.weitr, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "案 由", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 3860, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.anyou, font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "办案机关", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.badw, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "所处阶段", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 3860, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.jied, font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "指派日期", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.zprq, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 1500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "结案日期", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 3860, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.jarq || "", font: "宋体", size: 20 })] })] })
        ]
      })
    ]
  });

  const subsidyTable = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2500, 2500, 2180, 2180],
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "承办单位", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "湖南金厚（宁乡）律师事务所", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "案 由", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.anyou, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "案件编号", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.bianh, font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "承办人", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "陈伟", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "电话", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "13975892485", font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "结案时间", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.jarq || "", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "补贴金额", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "领款人", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2500, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "", font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, shading: { fill: "E8E8E8", type: ShadingType.CLEAR }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "受援人姓名", bold: true, font: "宋体", size: 20 })] })] }),
          new TableCell({ borders: border, width: { size: 2180, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: caseData.weitr, font: "宋体", size: 20 })] })] })
        ]
      })
    ]
  });

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "宋体", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "微软雅黑", color: "0F2A47" },
          paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "微软雅黑", color: "1E3A5F" },
          paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 1 } }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: "center",
            children: [new TextRun({ text: "法律援助案件承办情况通报/报告记录", bold: true, font: "微软雅黑", size: 22, color: "0F2A47" })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: "right",
            children: [new TextRun({ text: `受援人：${caseData.weitr}  承办人：陈伟  案件编号：${caseData.bianh}`, font: "宋体", size: 16, color: "666666" })]
          })]
        })
      },
      children: [
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "受援人：", bold: true, font: "宋体", size: 24 }),
            new TextRun({ text: caseData.weitr, font: "宋体", size: 24 }),
            new TextRun({ text: "    承办人：陈伟    案件编号：", bold: true, font: "宋体", size: 24 }),
            new TextRun({ text: caseData.bianh, font: "宋体", size: 24 })
          ]
        }),
        processTable,
        new Paragraph({ spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "填报说明：", bold: true, font: "宋体", size: 20 })] }),
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "1.方式为直接、邮寄、电话等；", font: "宋体", size: 18 })] }),
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "2.承办人需记录以下事项：接受指派及领取案卷，联系办案单位及办案人，阅卷，会见/约见受援人，调查取证，向办案单位联系沟通或者提交材料，向受援人通报案件进展情况、沟通案情、履行告知事项，律所集体讨论，向法援机构报告开庭日期，参加庭审，提交法律意见，领取结案文书，以及其他办理法援案件事项等。", font: "宋体", size: 18 })] }),
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          alignment: "center",
          spacing: { before: 200, after: 300 },
          children: [new TextRun({ text: "法律援助案件结案报告表", bold: true, font: "微软雅黑", size: 28, color: "0F2A47" })]
        }),
        reportTable,
        new Paragraph({ spacing: { before: 200, after: 50 }, children: [] }),
        new Paragraph({
          spacing: { before: 100, after: 50 },
          children: [
            new TextRun({ text: "援助形式：", bold: true, font: "宋体", size: 20 }),
            new TextRun({ text: "□刑事辩护  □刑事代理  ", font: "宋体", size: 20 }),
            new TextRun({ text: caseData.leib === "民事" ? "☑" : "□", font: "宋体", size: 20 }),
            new TextRun({ text: "民事诉讼代理  ", font: "宋体", size: 20 }),
            new TextRun({ text: caseData.leib === "行政" ? "☑" : "□", font: "宋体", size: 20 }),
            new TextRun({ text: "行政诉讼代理  □国家赔偿案件代理  □劳动争议调解与仲裁代理  □值班律师法律帮助  □其他", font: "宋体", size: 20 })
          ]
        }),
        new Paragraph({
          spacing: { before: 200, after: 50 },
          children: [new TextRun({ text: "承办情况小结（可附页）：", bold: true, font: "宋体", size: 20 })]
        }),
        new Paragraph({
          spacing: { after: 100 },
          shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
          children: [new TextRun({ text: caseData.cbxj || "（一、所做工作  \n二、基本案情  \n三、主要代理或者答辩意见  \n四、案件结果  \n五、办案心得）", font: "宋体", size: 20 })]
        }),
        new Paragraph({
          spacing: { before: 100, after: 50 },
          children: [new TextRun({ text: "承办结果：", bold: true, font: "宋体", size: 20 })]
        }),
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "民事/行政：□判决/裁决结案（□胜诉 □败诉 □部分胜诉）  □调解  □和解  □撤诉", font: "宋体", size: 18 })] }),
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "刑  事：□全部采纳 □部分采纳 □未采纳", font: "宋体", size: 18 })] }),
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "其  他：", font: "宋体", size: 18 })] }),
        new Paragraph({
          spacing: { before: 100, after: 50 },
          children: [
            new TextRun({ text: "备  注：", bold: true, font: "宋体", size: 20 }),
            new TextRun({ text: "为受援人挽回损失：              其他：", font: "宋体", size: 20 })
          ]
        }),
        new Paragraph({
          spacing: { before: 200, after: 100 },
          children: [
            new TextRun({ text: "填表时间：", font: "宋体", size: 20 }),
            new TextRun({ text: caseData.gdrq || "", font: "宋体", size: 20 }),
            new TextRun({ text: "                承办人签名：", font: "宋体", size: 20 })
          ]
        }),
        new Paragraph({
          alignment: "center",
          spacing: { before: 300, after: 200 },
          children: [new TextRun({ text: "宁乡市法律援助中心", bold: true, font: "微软雅黑", size: 24, color: "0F2A47" })]
        }),
        new Paragraph({
          alignment: "center",
          spacing: { before: 200, after: 300 },
          children: [new TextRun({ text: "支付援助案件补贴审批表", bold: true, font: "微软雅黑", size: 24, color: "0F2A47" })]
        }),
        subsidyTable,
        new Paragraph({ spacing: { before: 400 }, children: [] }),
        new Paragraph({
          children: [
            new TextRun({ text: "核卷人：                    审批人：", font: "宋体", size: 20 })
          ]
        }),
        new Paragraph({
          spacing: { before: 200 },
          children: [
            new TextRun({ text: "        年   月   日                                         年   月   日", font: "宋体", size: 20 })
          ]
        }),
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          alignment: "center",
          spacing: { before: 200, after: 300 },
          children: [new TextRun({ text: "卷 内 备 考 表", bold: true, font: "微软雅黑", size: 28, color: "0F2A47" })]
        }),
        new Paragraph({
          spacing: { before: 100, after: 200 },
          children: [new TextRun({ text: "立 卷 情 况 说 明：", bold: true, font: "宋体", size: 22 })]
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: caseData.ljsm || "", font: "宋体", size: 22 })]
        }),
        new Paragraph({
          spacing: { before: 400 },
          children: [
            new TextRun({ text: "本卷共          页，已按相关规定立卷归档。", font: "宋体", size: 22 })
          ]
        }),
        new Paragraph({
          spacing: { before: 300 },
          children: [
            new TextRun({ text: "立 卷 人：              检 查 人：              立卷时间：", font: "宋体", size: 22 })
          ]
        })
      ]
    }]
  });

  return doc;
}

async function main() {
  const inputDir = "d:\\trae\\法律援助文书";
  const elementFile = path.join(inputDir, "民事_元素.txt");
  const outputFile = path.join(inputDir, "法援结案_填充后.docx");

  console.log("Reading element file:", elementFile);
  const elements = parseElementFile(elementFile);
  console.log("Elements parsed successfully");

  const doc = fillTemplate(elements);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputFile, buffer);
  console.log("Generated:", outputFile);
}

main().catch(console.error);