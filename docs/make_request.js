const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak, PageNumber, Header, Footer,
} = require("docx");
const fs = require("fs");

const FONT = "Microsoft YaHei";
const INK = "1A1A1A", MUT = "4E5961", ACC = "0E5B6E", WARN = "A0480A";

const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 300 },
  alignment: o.align, indent: o.indent,
  children: [new TextRun({
    text, font: FONT, size: o.size ?? 21,
    bold: o.bold, italics: o.italics, color: o.color ?? INK,
  })],
});

const PR = (runs, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 300 },
  alignment: o.align, indent: o.indent,
  children: runs.map(r => new TextRun({
    text: r.t, font: FONT, size: r.size ?? o.size ?? 21,
    bold: r.b, color: r.c ?? o.color ?? INK,
  })),
});

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 300, after: 140 },
  children: [new TextRun({ text, font: FONT, size: 25, bold: true, color: INK })],
});
const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, font: FONT, size: 21, bold: true, color: ACC })],
});

const BUL = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 70, line: 290 },
  children: [new TextRun({ text, font: FONT, size: 21, color: INK })],
});

const BOX = (label, lines, color) => {
  const fill = color === WARN ? "FAEBDE" : "E0EDF1";
  const out = [new Paragraph({
    spacing: { before: 150, after: 0, line: 280 },
    shading: { type: ShadingType.CLEAR, fill },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color, space: 8 } },
    children: [new TextRun({ text: label, font: FONT, size: 19, bold: true, color })],
  })];
  lines.forEach((ln, i) => out.push(new Paragraph({
    spacing: { after: i === lines.length - 1 ? 190 : 60, line: 280 },
    shading: { type: ShadingType.CLEAR, fill },
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
    alignment: align, spacing: { after: 0, line: 260 },
    children: [new TextRun({
      text, font: FONT, size: 19, bold: b || head, color: color ?? INK })],
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
      { w: widths[ci], head: headFirst && ri === 0,
        b: typeof c === "object" && c.b,
        align: typeof c === "object" ? c.align : undefined,
        color: typeof c === "object" ? c.color : undefined },
    )),
  })),
});

const doc = new Document({
  creator: "button-battery study",
  title: "科研数据补充导出申请",
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } } },
  numbering: {
    config: [{ reference: "bullets", levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 420, hanging: 220 } } },
    }] }],
  },
  sections: [{
    properties: { page: { margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT, spacing: { after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0", space: 6 } },
        children: [new TextRun({
          text: "科研数据补充导出申请　·　已过食管纽扣电池研究",
          font: FONT, size: 16, color: MUT })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({
          children: ["第 ", PageNumber.CURRENT, " 页 / 共 ", PageNumber.TOTAL_PAGES, " 页"],
          font: FONT, size: 16, color: MUT })],
      })] }),
    },
    children: [
      // ================= 抬头 =================
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({
          text: "科研数据补充导出申请", font: FONT, size: 36, bold: true, color: INK })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [new TextRun({
          text: "已过食管纽扣电池的处置与转归研究（回顾性队列）",
          font: FONT, size: 22, color: MUT })],
      }),

      P("致：病案统计科　／　信息科", { bold: true, after: 160 }),

      table([1800, 2713, 1800, 2713], [
        [{ t: "申请人", b: true }, "　", { t: "科室", b: true }, "小儿外科"],
        [{ t: "联系电话", b: true }, "　", { t: "申请日期", b: true }, "　"],
        [{ t: "项目负责人", b: true }, "　", { t: "伦理批件号", b: true }, "　"],
      ], { headFirst: false }),
      P("", { after: 200 }),

      // ================= 一、理由 =================
      H1("一、申请理由"),
      P("本课题已于前期获得一份「近 10 年消化道异物」住院数据导出（2016-07 至 2026-06，共 1244 例次），并据此完成了队列构建与初步分析。分析过程中发现该份数据存在一处结构性缺口，须补充导出方可完成研究。"),
      ...BOX("缺口所在", [
        "前次导出以「消化道异物」为纳入条件。患儿若在本次出院后因电池腐蚀的",
        "迟发并发症（食管狭窄、食管气管瘘等）再次就诊，其诊断名称中不含「异物」",
        "字样，因而不会出现在该份导出结果中。",
        "",
        "实测：141 例纽扣电池患儿在该数据集中均无第二次就诊记录，但这一「零」",
        "由抽取条件本身决定，不能说明确实未发生并发症。",
      ], WARN),
      P("纽扣电池的迟发损伤典型出现在摄入后 1—3 周，而本队列住院天数中位数仅 1 天。若不补充出院后的就诊记录，本研究无法对迟发并发症作出任何陈述。"),

      // ================= 二、事项一 =================
      H1("二、申请事项一：索引住院后 1 年内的全部就诊记录"),
      P("这是本次申请的核心内容。", { bold: true }),

      H2("2.1　患者范围"),
      BUL("由申请人提供患者名单，沿用前次导出的「科研患者编号」，共 141 名。"),
      BUL("请在导出结果中保留同一「科研患者编号」，以便与前期数据关联。"),

      H2("2.2　时间窗"),
      BUL("起点：该患儿本次（索引）住院的出院日期。"),
      BUL("终点：出院后 365 天。"),

      H2("2.3　纳入条件（关键）"),
      ...BOX("请勿设置任何诊断限制", [
        "本次查询须涵盖上述时间窗内该患儿在本院的全部就诊记录——门诊、急诊、",
        "住院均需纳入，不限科室，不限诊断。",
        "",
        "若按诊断筛选，恰好会漏掉本研究要找的那部分病例。",
      ], ACC),
      P("为便于核对是否遗漏，以下为本研究重点关注的诊断（仅供参考，不作为筛选条件）：", { after: 100 }),
      table([4513, 4513], [
        ["重点关注诊断", "关注原因"],
        ["食管狭窄、食管良性狭窄", "电池腐蚀后最常见的迟发并发症"],
        ["食管气管瘘、气管食管瘘", "迟发重症"],
        ["主动脉食管瘘", "迟发致死性并发症"],
        ["纵隔炎、纵隔脓肿", "穿孔后继发"],
        ["消化道穿孔（各部位）", "迟发穿孔"],
        ["吞咽困难、进食梗阻、反复呕吐", "狭窄的早期表现"],
        ["食管扩张术、食管支架置入", "提示已发生狭窄"],
        ["声带麻痹、声音嘶哑", "喉返神经损伤"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      H2("2.4　需要的字段"),
      table([2600, 6426], [
        ["字段", "说明"],
        ["科研患者编号", "须与前次导出一致，否则无法关联"],
        ["科研就诊编号", "本次就诊的唯一标识"],
        ["就诊类型", "门诊 / 急诊 / 住院"],
        ["就诊日期、出院日期", "住院者两者均需"],
        ["就诊科别", "—"],
        ["主诉", "自由文本"],
        ["门（急）诊诊断名称", "含诊断编码"],
        ["出院诊断名称、诊断编码", "住院者；请提供全部诊断，不限主诊断"],
        ["是否手术或侵入性操作", "是 / 否"],
        ["手术（操作）名称、日期", "含内镜、食管扩张等操作"],
        ["出院情况 / 转归", "—"],
      ]),
      P("若同一患儿在时间窗内有多次就诊，请逐次列出，不要合并。", { color: WARN }),

      // ================= 三、事项二 =================
      H1("三、申请事项二：影像资料补充"),
      P("前次导出中，有 14 例纽扣电池患儿未见任何 X 线报告，但其病历记载已行腹部平片检查。推测为该次导出未覆盖，而非确无检查。"),
      H2("3.1　需要核查与补充的内容"),
      BUL("上述 14 例在索引住院期间的全部 X 线（DR）检查记录。"),
      BUL("检查日期与时间、检查名称、检查所见与结论全文。"),
      BUL("能够在 PACS 中定位到原始图像的标识（检查号 / Accession Number / Study UID 任一即可）。"),
      ...BOX("为何需要 PACS 定位标识", [
        "本研究需在原始 DR 影像上测量电池直径——直径 ≥20 mm 是文献中最强的",
        "风险因素，而病历中仅 3/141 例记录了规格。该信息可从影像中恢复，",
        "但须能定位到原始 DICOM 文件。",
        "",
        "如条件允许，恳请一并为全部 141 例提供该定位标识。",
      ], ACC),

      // ================= 四、事项三 =================
      H1("四、申请事项三：内镜图像（如可提供）"),
      P("前次导出的内镜相关记录仅有文字，无图像。本研究需由两名内镜医师对黏膜损伤进行盲法分级（Zargar 分级），文字记录不足以支撑。"),
      BUL("范围：49 例接受内镜检查或内镜下取出者。"),
      BUL("内容：内镜图像或录像的存档文件，及其与就诊编号的对应关系。"),
      P("此项为非必需。如系统不便批量导出，可改为提供检索路径，由研究者在院内系统内逐例查阅。", { color: MUT }),

      // ================= 五、格式与去标识 =================
      H1("五、数据形式与去标识要求"),
      BUL("格式：Excel（.xlsx）或 CSV，编码 UTF-8，与前次导出保持一致。"),
      BUL("每类数据一张工作表，表内每行一条记录。"),
      BUL("请以「科研患者编号」「科研就诊编号」标识，不要包含姓名、身份证号、家庭住址、联系电话。"),
      BUL("住院号如非关联所必需，请勿提供；确需提供时请以研究编号替代。"),

      // ================= 六、承诺 =================
      H1("六、数据使用承诺"),
      BUL("本次申请数据仅用于上述科研项目，不作任何其他用途。"),
      BUL("数据存放于院内指定设备，不上传至公共网络或第三方平台。"),
      BUL("发表成果中仅呈现汇总统计结果，不出现任何可识别到个人的信息。"),
      BUL("项目结束后按医院规定处理研究数据。"),
      BUL("本项目已通过（拟申请）本院伦理委员会审查，批件号见首页。"),

      // ================= 签字 =================
      new Paragraph({ spacing: { before: 320, after: 200 },
        children: [new TextRun({ text: "　", font: FONT, size: 21 })] }),
      table([2255, 2255, 2255, 2261], [
        [{ t: "申请人签字", b: true }, "　", { t: "日期", b: true }, "　"],
        [{ t: "科室负责人", b: true }, "　", { t: "日期", b: true }, "　"],
        [{ t: "病案科意见", b: true }, "　", { t: "日期", b: true }, "　"],
      ], { headFirst: false }),

      new Paragraph({
        spacing: { before: 260 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "D6DCE0", space: 8 } },
        children: [new TextRun({
          text: "附：141 例患者的「科研患者编号」名单随本申请一并提交（电子文件）。",
          font: FONT, size: 17, color: MUT })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(__dirname + "/数据补充导出申请.docx", b);
  console.log("已生成 数据补充导出申请.docx");
});
