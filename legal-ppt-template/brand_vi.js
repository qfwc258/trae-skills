const brandVI = {
  lawFirm: {
    name: "  ",
    tel: "139 7589 2485",
    branch: "湖南金厚（宁乡）律师事务所",
    address: "  "
  },
  colors: {
    // 现代商务蓝系 - Modern Business Blue
    primary: {
      deep: "1a365d",      // 深蓝 - 主色
      dark: "2c5282",      // 中蓝 - 辅助
      main: "3182ce",      // 亮蓝 - 强调
      mid: "4299e1",       // 中亮蓝
      light: "90cdf4"      // 浅蓝 - 装饰
    },
    accent: {
      gold: "C9A227",      // 金色（保留备用）
      goldLight: "E8D48A",
      warm: "D97706"
    },
    // 扩展蓝色系
    blue: {
      900: "1a365d",
      800: "2c5282",
      700: "3182ce",
      600: "4299e1",
      500: "63b3ed",
      400: "90cdf4",
      300: "bee3f8",
      200: "ebf8ff",       // 极浅蓝 - 背景
      100: "f0f9ff"
    },
    neutral: {
      900: "1A202C",
      800: "2D3748",
      700: "374151",
      600: "4B5563",
      500: "6B7280",
      400: "9CA3AF",
      300: "D1D5DB",
      200: "E5E7EB",
      100: "F3F4F6",
      50: "F9FAFB"
    },
    white: "FFFFFF",
    semantic: {
      success: "059669",
      warning: "D97706",
      danger: "DC2626",
      info: "2563EB"
    }
  },
  typography: {
    h1: { fontSize: 44, bold: true, color: "primary.deep", align: "left", fontFace: "Microsoft YaHei" },
    h2: { fontSize: 28, bold: true, color: "primary.dark", align: "left", fontFace: "Microsoft YaHei" },
    h3: { fontSize: 20, bold: true, color: "primary.main", align: "left", fontFace: "Microsoft YaHei" },
    h4: { fontSize: 16, bold: true, color: "neutral.700", align: "left", fontFace: "Microsoft YaHei" },
    body: { fontSize: 14, color: "neutral.700", align: "left", fontFace: "Microsoft YaHei" },
    bodySmall: { fontSize: 12, color: "neutral.600", align: "left", fontFace: "Microsoft YaHei" },
    caption: { fontSize: 10, color: "neutral.500", align: "left", fontFace: "Microsoft YaHei" }
  },
  spacing: {
    pageMargin: 0.6,
    contentPadding: 0.4,
    gridUnit: 0.125,
    sectionGap: 0.5,
    elementGap: 0.25
  },
  border: {
    radius: 0.08,
    width: 1
  }
};

module.exports = brandVI;