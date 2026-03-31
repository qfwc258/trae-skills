const { TextRun, Header, Footer, Paragraph, PageNumber } = require("docx");
const brandVI = require("./brand_vi");

function createHeader(firmName = brandVI.lawFirm.name, branch = brandVI.lawFirm.branch) {
  return {
    default: new Header({
      children: [
        new Paragraph({
          children: [
            new TextRun({ text: firmName, bold: true, font: "微软雅黑", size: 20, color: brandVI.colors.primary.dark }),
            new TextRun({ text: "    " }),
            new TextRun({ text: branch, font: "宋体", size: 18, color: brandVI.colors.neutral[500] })
          ],
          border: { bottom: { style: "single", size: 6, color: brandVI.colors.primary.main } }
        })
      ]
    })
  };
}

function createFooter(tel = brandVI.lawFirm.tel) {
  return {
    default: new Footer({
      children: [
        new Paragraph({
          alignment: "center",
          children: [
            new TextRun({ text: `联系电话: ${tel}    `, font: "宋体", size: 18, color: brandVI.colors.neutral[500] }),
            new TextRun({ children: ["第 ", PageNumber.CURRENT, " 页 / 共 ", PageNumber.TOTAL_PAGES, " 页"], font: "宋体", size: 18, color: brandVI.colors.neutral[500] })
          ]
        })
      ]
    })
  };
}

module.exports = {
  brandVI,
  createHeader,
  createFooter
};