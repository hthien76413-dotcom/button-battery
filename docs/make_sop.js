const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak, PageNumber, Header, Footer, ImageRun, convertInchesToTwip,
} = require("docx");
const fs = require("fs");

const FONT = "Microsoft YaHei";
const W = 9026;                       // A4 正文宽 (DXA)
const INK = "1A1A1A", MUT = "4E5961", ACC = "0E5B6E", WARN = "A0480A";

const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 300 },
  alignment: o.align,
  indent: o.indent,
  border: o.border,
  shading: o.shading,
  children: [new TextRun({
    text, font: FONT, size: o.size ?? 21,
    bold: o.bold, italics: o.italics,
    color: o.color ?? INK,
  })],
});

// 多段文字混排（用于局部加粗）
const PR = (runs, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 300 },
  alignment: o.align, indent: o.indent, shading: o.shading,
  children: runs.map(r => new TextRun({
    text: r.t, font: FONT, size: r.size ?? o.size ?? 21,
    bold: r.b, italics: r.i, color: r.c ?? o.color ?? INK,
  })),
});

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 320, after: 160 },
  children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: INK })],
});
const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 120 },
  children: [new TextRun({ text, font: FONT, size: 23, bold: true, color: ACC })],
});

const BUL = (text, o = {}) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text, font: FONT, size: 21, color: o.color ?? INK, bold: o.bold })],
});
const NUM = (text, o = {}) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text, font: FONT, size: 21, color: o.color ?? INK, bold: o.bold })],
});

// 提示框（整段底色 + 左侧粗边）
const BOX = (label, lines, color) => {
  const out = [new Paragraph({
    spacing: { before: 160, after: 0, line: 280 },
    shading: { type: ShadingType.CLEAR, fill: color === WARN ? "FAEBDE" : "E0EDF1" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color, space: 8 } },
    children: [new TextRun({ text: label, font: FONT, size: 19, bold: true, color })],
  })];
  lines.forEach((ln, i) => out.push(new Paragraph({
    spacing: { after: i === lines.length - 1 ? 200 : 60, line: 280 },
    shading: { type: ShadingType.CLEAR, fill: color === WARN ? "FAEBDE" : "E0EDF1" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color, space: 8 } },
    children: [new TextRun({ text: ln, font: FONT, size: 20, color: INK })],
  })));
  return out;
};

const cell = (text, { w, b, head, align, color } = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: head ? { type: ShadingType.CLEAR, fill: "ECEFF1" } : undefined,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  children: [new Paragraph({
    alignment: align,
    spacing: { after: 0, line: 260 },
    children: [new TextRun({
      text, font: FONT, size: 19, bold: b || head, color: color ?? INK,
    })],
  })],
});

const table = (widths, rows, { headFirst = true } = {}) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: "BFC8CE" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFC8CE" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "BFC8CE" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "BFC8CE" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0" },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0" },
  },
  rows: rows.map((r, ri) => new TableRow({
    tableHeader: headFirst && ri === 0,
    children: r.map((c, ci) => cell(
      typeof c === "string" ? c : c.t,
      { w: widths[ci], head: headFirst && ri === 0, b: typeof c === "object" && c.b,
        align: typeof c === "object" ? c.align : undefined,
        color: typeof c === "object" ? c.color : undefined },
    )),
  })),
});

const doc = new Document({
  creator: "button-battery study",
  title: "纽扣电池直径测量标准操作规程",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21, color: INK } },
    },
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 420, hanging: 220 } } },
      }] },
      { reference: "steps", levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 420, hanging: 220 } } },
      }] },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0", space: 6 } },
        children: [new TextRun({
          text: "纽扣电池直径测量 SOP　·　已过食管纽扣电池研究",
          font: FONT, size: 16, color: MUT })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ["第 ", PageNumber.CURRENT, " 页 / 共 ", PageNumber.TOTAL_PAGES, " 页"],
          font: FONT, size: 16, color: MUT })],
      })] }),
    },
    children: [
      // ============ 封面区 ============
      new Paragraph({
        spacing: { before: 240, after: 60 },
        children: [new TextRun({ text: "标准操作规程 (SOP)", font: FONT, size: 18, bold: true, color: ACC, characterSpacing: 40 })],
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: "纽扣电池直径测量", font: FONT, size: 44, bold: true, color: INK })],
      }),
      P("在 PACS 的 DR 影像上恢复病历未记录的电池直径", { size: 24, color: MUT, after: 240 }),

      table([2000, 2513, 2000, 2513], [
        [{ t: "文件编号", b: true }, "BB-SOP-01", { t: "版本", b: true }, "v1.0"],
        [{ t: "适用研究", b: true }, "已过食管纽扣电池是否必须急诊取出", { t: "生效日期", b: true }, "　"],
        [{ t: "编写", b: true }, "　", { t: "审核", b: true }, "　"],
      ], { headFirst: false }),
      P("", { after: 200 }),

      // ============ 1 背景 ============
      H1("1　背景与目的"),
      PR([
        { t: "电池直径 " }, { t: "≥20 mm", b: true },
        { t: " 是文献中最强的风险因素（任何并发症 " }, { t: "OR 4.34", b: true },
        { t: "），也是本研究风险分层算法（Fig 3）的一条分层线。但本队列 " },
        { t: "125 例中仅 2 例", b: true }, { t: "病历记录了电池规格。" },
      ]),
      P("这个信息并没有丢失——它一直在已经拍过的腹平片里，只是从未有人从图像上读取。本规程的目的，就是把这个变量从影像中恢复出来。"),
      ...BOX("为什么必须校正放大", [
        "投照为锥形束，电池在体内离探测器有距离，成像会放大约 1.1–1.2 倍。",
        "而分层阈值恰好是 20 mm：一枚真实 20 mm 的 CR2032 可能量出 22–24 mm，",
        "一枚 16 mm 的 CR1620 可能量出 18–19 mm。直接拿卡尺读数去卡 20 mm，",
        "会系统性高估，把小电池误判为大电池——错在最关键的那条线上。",
      ], WARN),

      // ============ 2 人员 ============
      H1("2　人员与盲法"),
      table([2300, 6726], [
        ["角色", "要求"],
        [{ t: "测量者 A / B", b: true }, "2 名，各自独立完成测量，过程中不得知晓对方读数"],
        [{ t: "盲法", b: true }, "测量时不得知晓该患儿的结局——有无黏膜损伤、是否接受内镜取出、住院天数。建议由第三方先行剥离结局字段后再交付影像清单"],
        [{ t: "裁定者", b: true }, "1 名，仅在两名测量者分类不一致时介入"],
        [{ t: "记录者", b: true }, "按第 7 节字段录入，不参与读数"],
      ]),
      ...BOX("盲法为什么重要", [
        "知道某个孩子后来有食管溃疡，会不自觉地把边缘读得大一些。",
        "审稿人一定会问测量是否设盲，以及测量者间一致性（ICC）。",
      ], ACC),

      // ============ 3 第一步 ============
      H1("3　第一步：确认影像可测量"),
      P("先做这一步，不要直接开始量。若整批影像缺少标定信息，测量无从谈起，应改走备选方案（第 8 节）。"),
      P("导出若干份 DICOM 至同一目录，运行："),
      new Paragraph({
        spacing: { before: 60, after: 160, line: 280 },
        shading: { type: ShadingType.CLEAR, fill: "ECEFF1" },
        children: [new TextRun({
          text: "  python3 scripts/dicom_calibration.py audit /path/to/dicom_dir",
          font: "Consolas", size: 19, color: INK })],
      }),
      P("脚本逐文件报告三项，据此判断该片能否用于测量：", { after: 100 }),
      table([2600, 3400, 3026], [
        ["检查项", "含义", "不合格时"],
        ["像素间距（ImagerPixelSpacing）", "每像素对应多少毫米，是卡尺换算的基础", "该片不可测量，记为缺失"],
        ["SID / SOD", "球管到探测器 / 到物体的距离，用于算放大倍数", "改用经验放大倍数，并按第 6 节判定是否为不确定病例"],
        ["ImageType", "若为 DERIVED / SECONDARY，说明图像可能已重采样", "标定存疑，不得用于测量"],
      ]),
      ...BOX("不要忽略这一步的另一个收获", [
        "导出数据集中有 14 例标注“无 X 线片”，但那只说明本次导出里没有，",
        "不代表 PACS 里没有。核查标定的同时请一并查这 14 例，很可能片子是在的。",
      ], ACC),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ 4 第二步 ============
      H1("4　第二步：在 PACS 中测量"),
      NUM("调取该患儿的首次腹部平片（正位）。使用首次片，不要用复查片——研究定义的是就诊时的电池规格。"),
      NUM("先确认这是电池而非硬币：正位可见双环征（double-halo）或台阶征，侧位可见斜面征（beveled edge）。若无法与硬币区分，记入待裁定。"),
      NUM("使用 PACS 的测量/卡尺工具，沿圆盘直径拉一条线，软件会读取 DICOM 标定直接给出毫米数。"),
      NUM("量外环（见下图）。双环征的外环是电池卷边外缘，那才是标称直径；内环偏小，量错会系统性低估。"),
      NUM("同一枚电池量 3 次，取中位数，保留 2 位小数，记为“投影直径”。"),
      NUM("若为多枚电池，逐枚测量并编号（#1、#2…），分别记录。"),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new ImageRun({
          type: "png",
          data: fs.readFileSync(__dirname + "/halo.png"),
          transformation: { width: 580, height: 251 },
        })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 220 },
        children: [new TextRun({
          text: "图 1　双环征的测量位置。卡尺应跨越外环全径；量内环会系统性低估约 3–5 mm，"
              + "足以把一枚 20 mm 电池误判为 16 mm——恰好跨过分层线。",
          font: FONT, size: 18, color: MUT })],
      }),
      ...BOX("这一步量到的是投影直径，不是真实直径", [
        "PACS 给出的毫米数是探测器平面上的尺寸，仍包含放大。",
        "不要直接把这个数填进“电池直径”字段，必须先做第 5 节的校正。",
      ], WARN),

      // ============ 5 第三步 ============
      H1("5　第三步：放大校正与规格吸附"),
      H2("5.1　运行校正"),
      P("以第 3 节 audit 输出中的 SID、SOD 代入："),
      new Paragraph({
        spacing: { before: 60, after: 60, line: 280 },
        shading: { type: ShadingType.CLEAR, fill: "ECEFF1" },
        children: [new TextRun({
          text: "  python3 scripts/dicom_calibration.py size 22.70 --sid 1000 --sod 880",
          font: "Consolas", size: 19, color: INK })],
      }),
      P("若 SOD 缺失，改用体表到探测器的估计距离（cm）："),
      new Paragraph({
        spacing: { before: 60, after: 160, line: 280 },
        shading: { type: ShadingType.CLEAR, fill: "ECEFF1" },
        children: [new TextRun({
          text: "  python3 scripts/dicom_calibration.py size 22.70 --sid 1000 --odd 10",
          font: "Consolas", size: 19, color: INK })],
      }),
      P("放大倍数 M = SID / SOD；真实直径 = 投影直径 / M。"),

      H2("5.2　吸附到标准规格"),
      P("纽扣电池的直径不是连续的，只有有限几档标准值。因此不需要测得非常精确，只需判断落在哪一档。脚本会自动完成吸附。"),
      table([1800, 3200, 4026], [
        ["直径", "常见型号", "说明"],
        ["12.5 mm", "CR1216 / 1220 / 1225", "小型锂电池"],
        ["16.0 mm", "CR1616 / 1620 / 1632", "—"],
        [{ t: "20.0 mm", b: true }, { t: "CR2016 / 2025 / 2032", b: true }, { t: "分层阈值所在档，最常见", b: true, color: WARN }],
        ["23.0 mm", "CR2320 / 2330 / 2354", "—"],
        ["24.5 mm", "CR2430 / 2450", "—"],
        ["< 12 mm", "SR/LR 系列助听器、手表电池", "4.8 / 5.8 / 6.8 / 7.9 / 9.5 / 11.6 mm"],
      ]),
      ...BOX("CR 型号的命名规则", [
        "前两位数字即直径（毫米），后两位为厚度（十分之一毫米）。",
        "CR2032 = 直径 20 mm、厚 3.2 mm。若内镜取出后见到实物或型号，可直接据此登记，无需测量。",
      ], ACC),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ 6 不确定 ============
      H1("6　第四步：不确定病例的判定"),
      P("脚本会在物体距探测器 5–20 cm 的合理范围内扫一遍放大倍数，看 ≥20 mm 这个分类是否稳定。"),
      table([4513, 4513], [
        ["脚本输出", "处置"],
        ["分类在整个范围内稳定", "按该分类登记（≥20 mm 或 <20 mm）"],
        [{ t: "分类随放大倍数改变", b: true, color: WARN },
         { t: "登记为 indeterminate（不确定），不要强行给一个值", b: true, color: WARN }],
      ]),
      P("举例：投影读数 21.0 mm、SOD 未知时，物体距探测器 5 cm 则校正为 19.95 mm（吸附 20 mm），20 cm 则校正为 16.80 mm（吸附 16 mm）——同一个读数跨过了分层线。这类病例必须记为不确定。"),
      ...BOX("不确定不是失败", [
        "一批诚实标注的 indeterminate，写进敏感性分析，比假装每一例都测得准要经得起审稿。",
        "分析时对这些病例分别按 ≥20 mm 和 <20 mm 两种极端假设各跑一次，看结论是否改变。",
      ], ACC),

      // ============ 7 记录字段 ============
      H1("7　数据记录字段"),
      P("每枚电池一行。两名测量者各填一份，最后合并比对。"),
      table([2600, 1500, 4926], [
        ["字段", "类型", "说明"],
        ["科研就诊编号", "整数", "与主数据集对应，不要填住院号"],
        ["battery_index", "整数", "同一患儿多枚时编号，单枚填 1"],
        ["reader", "A / B", "测量者标识"],
        ["projected_mm", "小数", "PACS 卡尺读数（3 次中位），2 位小数"],
        ["sid_mm / sod_mm", "小数", "取自 audit 输出，缺失留空"],
        ["magnification", "小数", "脚本给出，或注明所用经验倍数"],
        ["true_mm", "小数", "校正后真实直径"],
        ["snapped_mm", "小数", "吸附后的标准规格"],
        ["size_class", "文本", "≥20mm / <20mm / indeterminate"],
        ["image_usable", "是 / 否", "标定缺失或 DERIVED 则填否"],
        ["note", "文本", "边缘不清、与硬币难辨、多枚重叠等情况在此说明"],
      ]),

      // ============ 8 质控 ============
      H1("8　质量控制"),
      BUL("两名测量者独立完成全部病例；若工作量过大，至少独立重复 30 例子集用于计算一致性。"),
      BUL("计算测量者间组内相关系数（ICC）与分类一致率（Cohen κ）。ICC < 0.75 应重新培训后重测。"),
      BUL("分类不一致的病例交裁定者判定，裁定过程同样设盲。"),
      BUL("定期抽查盲法是否被破坏（例如影像清单中混入了结局信息）。"),
      BUL("测量结果不得回填进原始病历系统，仅用于研究数据集。"),

      H2("8.1　备选方案：标定缺失时"),
      P("若 audit 显示整批影像无像素间距，改用同层椎体宽度作内参：在同一张片、与电池大致同深度处测量椎体横径，以该儿童年龄段椎体宽度的参考值换算比例。此法误差更大，应在论文局限中写明，并把全部病例记为近似值。"),

      // ============ 9 FAQ ============
      H1("9　常见问题"),
      H2("电池边缘模糊，量不准"),
      P("适当调整窗宽窗位提高边缘对比；仍不清晰则在 note 中注明，并由两名测量者分别读数，差值超过 1 mm 时交裁定。"),
      H2("分不清是电池还是硬币"),
      P("硬币为均质圆盘，无双环征与台阶征；侧位电池可见斜面边缘。仍无法区分者不得纳入，交病例核对环节处理。"),
      H2("要不要用侧位片"),
      P("正位用于测量直径。侧位的价值在于确认斜面征，以及在 SOD 缺失时帮助估计电池距探测器的深度。"),
      H2("多枚电池重叠"),
      P("选择边缘完整、未被遮挡的一枚测量；若全部重叠，改用后续复查片中分离开的时相，并在 note 中注明所用片次。"),

      // ============ 附录 ============
      new Paragraph({ children: [new PageBreak()] }),
      H1("附录　测量记录表"),
      P("可打印后手工填写，或按第 7 节字段建立电子表格。", { color: MUT }),
      table([1500, 900, 800, 1400, 1200, 1200, 2026], [
        ["就诊编号", "枚号", "读者", "投影 mm", "倍数", "吸附 mm", "分类 / 备注"],
        ...Array.from({ length: 12 }, () => ["　", "　", "　", "　", "　", "　", "　"]),
      ]),
      P("", { after: 200 }),
      PR([
        { t: "测量者签名：", b: true }, { t: "　　　　　　　　　　　　" },
        { t: "日期：", b: true }, { t: "　　　　　　　　" },
      ]),
      P("", { after: 240 }),
      new Paragraph({
        spacing: { before: 200 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0", space: 8 } },
        children: [new TextRun({
          text: "本规程随研究方案 PROTOCOL.md 一并归档；配套脚本 scripts/dicom_calibration.py。所有测量数据含患者层面信息，不得提交至代码仓库。",
          font: FONT, size: 17, color: MUT })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync("纽扣电池直径测量SOP.docx", b);
  console.log("已生成 纽扣电池直径测量SOP.docx");
});
