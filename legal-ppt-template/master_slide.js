const brandVI = require("./brand_vi");

const masterSlides = {
  titleSlide: {
    addHeader: function(slide, pres, title, subtitle, pageNum) {
      const colors = brandVI.colors;
      const spacing = brandVI.spacing;

      slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 0.06,
        fill: { color: colors.primary.deep }
      });

      slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 0.06, w: 10, h: 0.02,
        fill: { color: colors.accent.gold }
      });

      if (title) {
        slide.addText(title, {
          x: spacing.pageMargin, y: 0.25, w: 5, h: 0.5,
          fontSize: 28, bold: true,
          color: colors.primary.dark,
          fontFace: "Microsoft YaHei", margin: 0
        });
      }

      if (subtitle) {
        slide.addText(subtitle, {
          x: 5.5, y: 0.32, w: 2.5, h: 0.35,
          fontSize: 12, color: colors.neutral[400],
          fontFace: "Microsoft YaHei", align: "right", margin: 0
        });
      }

      if (pageNum) {
        slide.addText(`${String(pageNum).padStart(2, "0")} / ${String(this.totalPages || 10).padStart(2, "0")}`, {
          x: 8.3, y: 0.32, w: 1.1, h: 0.35,
          fontSize: 11, color: colors.neutral[500],
          fontFace: "Microsoft YaHei", align: "right", margin: 0
        });
      }

      slide.addShape(pres.ShapeType.rect, {
        x: spacing.pageMargin, y: 0.82, w: 1.5, h: 0.04,
        fill: { color: colors.accent.gold }
      });
    },

    addFooter: function(slide, pres) {
      const colors = brandVI.colors;
      const spacing = brandVI.spacing;

      slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 5.45, w: 10, h: 0.175,
        fill: { color: colors.primary.deep }
      });

      slide.addText(brandVI.lawFirm.name, {
        x: spacing.pageMargin, y: 5.48, w: 4, h: 0.12,
        fontSize: 9, color: colors.neutral[400], fontFace: "Microsoft YaHei", margin: 0
      });
    },

    apply: function(slide, pres, options = {}) {
      const { title, subtitle, pageNum, totalPages, showHeader = true, showFooter = true } = options;
      this.totalPages = totalPages;
      if (showHeader) this.addHeader(slide, pres, title, subtitle, pageNum);
      if (showFooter) this.addFooter(slide, pres);
    }
  },

  contentSlide: {
    background: function(slide, pres) {
      slide.background = { color: brandVI.colors.neutral[50] };
    },

    headerBand: function(slide, pres, title, subtitle, pageNum, totalPages) {
      masterSlides.titleSlide.addHeader(slide, pres, title, subtitle, pageNum);
      masterSlides.titleSlide.totalPages = totalPages;
    },

    footerBand: function(slide, pres) {
      masterSlides.titleSlide.addFooter(slide, pres);
    },

    card: function(slide, pres, x, y, w, h, options = {}) {
      const colors = brandVI.colors;
      const border = brandVI.border;

      slide.addShape(pres.ShapeType.rect, {
        x: x, y: y, w: w, h: h,
        fill: { color: options.fill || colors.white },
        line: { color: options.borderColor || colors.neutral[200], width: border.width }
      });

      if (options.headerColor) {
        slide.addShape(pres.ShapeType.rect, {
          x: x, y: y, w: w, h: options.headerHeight || 0.45,
          fill: { color: options.headerColor }
        });
      }

      return { x, y, w, h };
    },

    accentBar: function(slide, pres, x, y, h, color) {
      const colors = brandVI.colors;
      slide.addShape(pres.ShapeType.rect, {
        x: x, y: y, w: 0.08, h: h,
        fill: { color: color || colors.accent.gold }
      });
    },

    sectionTitle: function(slide, pres, x, y, text, color) {
      const colors = brandVI.colors;
      slide.addText(text, {
        x: x, y: y, w: 3, h: 0.35,
        fontSize: 16, bold: true,
        color: color || colors.primary.dark,
        fontFace: "Microsoft YaHei", margin: 0
      });
    },

    bodyText: function(slide, pres, x, y, text, w) {
      const colors = brandVI.colors;
      slide.addText(text, {
        x: x, y: y, w: w || 8, h: 0.3,
        fontSize: 12, color: colors.neutral[700],
        fontFace: "Microsoft YaHei", margin: 0
      });
    },

    apply: function(slide, pres, options = {}) {
      const { title, subtitle, pageNum, totalPages, showHeader = true, showFooter = true } = options;
      this.background(slide, pres);
      if (showHeader) this.headerBand(slide, pres, title, subtitle, pageNum, totalPages);
      if (showFooter) this.footerBand(slide, pres);
    }
  },

  coverSlide: {
    background: function(slide, pres, style = 'default') {
      const colors = brandVI.colors;

      if (style === 'split') {
        slide.background = { color: colors.primary.deep };
        slide.addShape(pres.ShapeType.rect, {
          x: 0, y: 0, w: 6, h: 5.625,
          fill: { color: colors.primary.deep }
        });
        slide.addShape(pres.ShapeType.rect, {
          x: 5.8, y: 0, w: 4.2, h: 5.625,
          fill: { color: colors.primary.dark, transparency: 50 }
        });
      } else if (style === 'bottomAccent') {
        slide.background = { color: colors.primary.deep };
        slide.addShape(pres.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: colors.primary.dark, transparency: 30 }
        });
        slide.addShape(pres.ShapeType.rect, {
          x: 0, y: 4.8, w: 10, h: 0.825,
          fill: { color: colors.accent.gold }
        });
      } else {
        slide.background = { color: colors.primary.deep };
        slide.addShape(pres.ShapeType.rect, {
          x: 0, y: 0, w: 10, h: 5.625,
          fill: { color: colors.primary.dark, transparency: 30 }
        });
      }
    },

    verticalAccentBar: function(slide, pres, x, y, h) {
      const colors = brandVI.colors;
      slide.addShape(pres.ShapeType.rect, {
        x: x, y: y, w: 0.2, h: h,
        fill: { color: colors.accent.gold }
      });
    },

    geometricDecoration: function(slide, pres, style = 'default') {
      const colors = brandVI.colors;

      if (style === 'lecture') {
        for (let i = 0; i < 5; i++) {
          slide.addShape(pres.ShapeType.rect, {
            x: 7.5 + (i % 3) * 0.6, y: 0.5 + Math.floor(i / 3) * 0.6,
            w: 0.5, h: 0.5,
            fill: { color: colors.accent.gold, transparency: 70 - i * 10 }
          });
        }
      } else if (style === 'dd') {
        for (let i = 0; i < 8; i++) {
          slide.addShape(pres.ShapeType.rect, {
            x: 6.5 + (i % 4) * 0.8, y: 0.5 + Math.floor(i / 4) * 2.5,
            w: 0.6, h: 0.6,
            fill: { color: colors.accent.gold, transparency: 60 + i * 5 }
          });
        }
      } else if (style === 'cr') {
        for (let i = 0; i < 20; i++) {
          slide.addShape(pres.ShapeType.rect, {
            x: 0.3 + i * 0.5, y: 1.0 + (i % 3) * 0.3,
            w: 0.35, h: 0.35,
            fill: { color: colors.accent.gold, transparency: 80 - i * 2 }
          });
        }
      }
    },

    titleBlock: function(slide, pres, type, title, subtitle, options = {}) {
      const colors = brandVI.colors;

      if (type) {
        slide.addText(type, {
          x: options.typeX || 0.5, y: options.typeY || 1.2,
          w: options.typeW || 9, h: 0.4,
          fontSize: 14, color: colors.accent.gold,
          fontFace: "Microsoft YaHei", margin: 0
        });
      }

      slide.addText(title, {
        x: options.titleX || 0.5, y: options.titleY || (type ? 1.8 : 1.5),
        w: options.titleW || 8, h: options.titleH || 0.8,
        fontSize: options.titleSize || 40, bold: true,
        color: colors.neutral[50],
        fontFace: "Microsoft YaHei", margin: 0
      });

      if (subtitle) {
        slide.addText(subtitle, {
          x: options.subX || 0.5, y: options.subY || (type ? 2.7 : 2.3),
          w: options.subW || 8, h: 0.5,
          fontSize: 18, color: colors.neutral[300],
          fontFace: "Microsoft YaHei", margin: 0
        });
      }
    },

    metaLine: function(slide, pres, y, text, w) {
      const colors = brandVI.colors;
      slide.addShape(pres.ShapeType.line, {
        x: 0.5, y: y, w: w || 3, h: 0,
        line: { color: colors.neutral[500], width: 1 }
      });
      slide.addText(text, {
        x: 0.5, y: y + 0.15, w: w || 5, h: 0.3,
        fontSize: 12, color: colors.neutral[400],
        fontFace: "Microsoft YaHei", margin: 0
      });
    },

    apply: function(slide, pres, options = {}) {
      const { style = 'default', type, title, subtitle, meta } = options;
      this.background(slide, pres, style);
      if (options.decoration) this.geometricDecoration(slide, pres, options.decoration);
      if (options.accentBar) this.verticalAccentBar(slide, pres, options.accentBar.x, options.accentBar.y, options.accentBar.h);
      if (title) this.titleBlock(slide, pres, type, title, subtitle, options);
      if (meta) this.metaLine(slide, pres, meta.y, meta.text, meta.w);
    }
  },

  endingSlide: {
    background: function(slide, pres) {
      const colors = brandVI.colors;
      slide.background = { color: colors.primary.deep };
      slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 5.625,
        fill: { color: colors.primary.dark, transparency: 40 }
      });
    },

    accentBar: function(slide, pres) {
      const colors = brandVI.colors;
      slide.addShape(pres.ShapeType.rect, {
        x: 4.5, y: 2.0, w: 0.1, h: 1.5,
        fill: { color: colors.accent.gold }
      });
    },

    thankYouText: function(slide, pres, mainText, subText) {
      const colors = brandVI.colors;
      slide.addText(mainText, {
        x: 4.8, y: 2.0, w: 4.5, h: 0.7,
        fontSize: 36, bold: true, color: colors.neutral[50],
        fontFace: "Microsoft YaHei", margin: 0
      });
      slide.addText(subText || "THANK YOU", {
        x: 4.8, y: 2.7, w: 4.5, h: 0.5,
        fontSize: 16, color: colors.neutral[400],
        fontFace: "Microsoft YaHei", margin: 0
      });
    },

    contactInfo: function(slide, pres) {
      const colors = brandVI.colors;
      slide.addShape(pres.ShapeType.line, {
        x: 4.8, y: 3.4, w: 4, h: 0,
        line: { color: colors.neutral[600], width: 0.5 }
      });
      slide.addText(brandVI.lawFirm.name, {
        x: 4.8, y: 3.6, w: 4.5, h: 0.35,
        fontSize: 14, color: colors.accent.goldLight,
        fontFace: "Microsoft YaHei", margin: 0
      });
      slide.addText(brandVI.lawFirm.branch, {
        x: 4.8, y: 3.95, w: 4.5, h: 0.3,
        fontSize: 11, color: colors.neutral[400],
        fontFace: "Microsoft YaHei", margin: 0
      });
      slide.addText(`TEL: ${brandVI.lawFirm.tel}`, {
        x: 4.8, y: 4.3, w: 4.5, h: 0.3,
        fontSize: 11, color: colors.neutral[400],
        fontFace: "Microsoft YaHei", margin: 0
      });
    },

    apply: function(slide, pres, options = {}) {
      const { mainText = "感谢聆听", subText } = options;
      this.background(slide, pres);
      this.accentBar(slide, pres);
      this.thankYouText(slide, pres, mainText, subText);
      this.contactInfo(slide, pres);
    }
  }
};

module.exports = masterSlides;