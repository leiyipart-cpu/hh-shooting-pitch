import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const W = 1280;
const H = 720;
const M = 64;

const C = {
  navy: "#071925",
  navy2: "#0D2633",
  ink: "#13212B",
  sea: "#31566A",
  mist: "#DDE4E5",
  paper: "#F2EFE8",
  white: "#FFFFFF",
  grey: "#829099",
  grey2: "#AAB3B7",
  line: "#CBD2D1",
  red: "#E3272E",
  acid: "#D7F03A",
  rose: "#D6A09A",
  brass: "#B98B45",
};

const FONT_SANS = "Helvetica Neue";
const FONT_CN = "PingFang SC";
const FONT_SERIF = "Songti SC";

const ROOT = "/Users/leiyi/Documents/ChatGPT/HH拍摄pitch";
const BUILD = `${ROOT}/tmp_hh_pitch/deck_build_v2`;
const OUT = `${ROOT}/output`;

const ASSET_PATHS = {
  sailing: `${ROOT}/assets/generated/sailing-crew-hero.png`,
  logbook: `${ROOT}/assets/generated/logbook-hero.png`,
  trail: `${ROOT}/assets/generated/trail-hero.png`,
  shoe: `${ROOT}/assets/generated/trail-shoe-grip.png`,
  night: `${ROOT}/assets/generated/sailing-nightglow.png`,
  hands: `${ROOT}/assets/generated/sailing-hands.png`,
  afterglow: `${ROOT}/assets/generated/trail-afterglow.png`,
  brief07: `${ROOT}/tmp_hh_pitch/renders/brief/page-07.jpg`,
  brief16: `${ROOT}/tmp_hh_pitch/renders/brief/page-16.jpg`,
  brief22: `${ROOT}/tmp_hh_pitch/renders/brief/page-22.jpg`,
  brief27: `${ROOT}/tmp_hh_pitch/renders/brief/page-27.jpg`,
  brief35: `${ROOT}/tmp_hh_pitch/renders/brief/page-35.jpg`,
  brief44: `${ROOT}/tmp_hh_pitch/renders/brief/page-44.jpg`,
};

function addShape(slide, x, y, w, h, fill = "none", lineFill = "none", lineWidth = 0, radius = 0, name) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
    ...(radius ? { borderRadius: radius } : {}),
  });
}

function addLine(slide, x, y, w, color = C.red, width = 2, name) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x, top: y, width: w, height: 0 },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function addVLine(slide, x, y, h, color = C.line, width = 2, name) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x, top: y, width: 0, height: h },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function addText(slide, value, x, y, w, h, opts = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: opts.fill ?? "none",
    line: { style: "solid", fill: opts.lineFill ?? "none", width: opts.lineWidth ?? 0 },
    ...(opts.radius ? { borderRadius: opts.radius } : {}),
  });
  shape.text = value;
  shape.text.style = {
    fontSize: opts.fontSize ?? 24,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    color: opts.color ?? C.ink,
    alignment: opts.align ?? "left",
    verticalAlignment: opts.valign ?? "top",
    typeface: opts.typeface ?? FONT_CN,
    lineSpacing: opts.lineSpacing ?? 1.0,
    autoFit: opts.autoFit ?? "none",
    insets: opts.insets ?? { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return shape;
}

function addImage(slide, bytes, x, y, w, h, alt, opts = {}) {
  return slide.images.add({
    blob: bytes,
    contentType: "image/png",
    alt,
    fit: opts.fit ?? "cover",
    position: { left: x, top: y, width: w, height: h },
    ...(opts.crop ? { crop: opts.crop } : {}),
    ...(opts.radius ? { geometry: "roundRect", borderRadius: opts.radius } : {}),
  });
}

function addTag(slide, label, x, y, opts = {}) {
  const width = opts.width ?? Math.max(120, label.length * 15 + 32);
  addShape(slide, x, y, width, 34, opts.fill ?? C.red, "none", 0, 17);
  addText(slide, label, x + 16, y + 6, width - 32, 22, {
    fontSize: 14,
    bold: true,
    color: opts.color ?? C.white,
    typeface: FONT_SANS,
    valign: "middle",
  });
}

function footer(slide, page, dark = false) {
  const actualPage = SLIDE_NUMBERS.get(slide) ?? page;
  const color = dark ? "#B8C4C8" : C.grey;
  addLine(slide, M, 688, 120, C.red, 3, `footer-rule-${actualPage}`);
  addText(slide, "HELLY HANSEN 150 · THE LOGBOOK", M + 140, 678, 420, 22, {
    fontSize: 12,
    bold: true,
    color,
    typeface: FONT_SANS,
    valign: "middle",
  });
  addText(slide, String(actualPage).padStart(2, "0"), 1160, 678, 56, 22, {
    fontSize: 12,
    bold: true,
    color,
    typeface: FONT_SANS,
    align: "right",
    valign: "middle",
  });
}

function note(slide, lines) {
  slide.speakerNotes.textFrame.setText(`[Sources]\n${lines.map((line) => `- ${line}`).join("\n")}`);
}

function titleBlock(slide, kicker, title, page, dark = false, subtitle) {
  const titleColor = dark ? C.white : C.navy;
  const bodyColor = dark ? "#C5D0D4" : C.sea;
  addText(slide, kicker.toUpperCase(), M, 48, 560, 22, {
    fontSize: 14,
    bold: true,
    color: dark ? C.red : C.red,
    typeface: FONT_SANS,
  });
  addText(slide, title, M, 82, 1120, 72, {
    fontSize: 48,
    bold: true,
    color: titleColor,
    lineSpacing: 0.95,
  });
  if (subtitle) {
    addText(slide, subtitle, M, 154, 1080, 54, {
      fontSize: 22,
      color: bodyColor,
      lineSpacing: 1.1,
    });
  }
  footer(slide, page, dark);
}

function metric(slide, value, label, x, y, w, dark = false, accent = C.red) {
  addLine(slide, x, y, Math.min(86, w), accent, 4);
  addText(slide, value, x, y + 20, w, 64, {
    fontSize: 42,
    bold: true,
    color: dark ? C.white : C.navy,
    typeface: FONT_SANS,
  });
  addText(slide, label, x, y + 82, w, 70, {
    fontSize: 20,
    color: dark ? "#C5D0D4" : C.sea,
    lineSpacing: 1.15,
  });
}

function addChapterNumber(slide, n, x, y, dark = false) {
  addText(slide, n, x, y, 100, 90, {
    fontSize: 82,
    bold: true,
    color: dark ? "#FFFFFF/18" : "#071925/12",
    typeface: FONT_SANS,
  });
}

const SLIDE_NUMBERS = new WeakMap();

function newSlide(presentation) {
  const slide = presentation.slides.add();
  SLIDE_NUMBERS.set(slide, presentation.slides.items.length);
  return slide;
}

async function main() {
  await fs.mkdir(`${BUILD}/rendered`, { recursive: true });
  await fs.mkdir(OUT, { recursive: true });
  const entries = await Promise.all(Object.entries(ASSET_PATHS).map(async ([key, path]) => [key, await fs.readFile(path)]));
  const A = Object.fromEntries(entries);
  const p = Presentation.create({ slideSize: { width: W, height: H } });

  // 01 — Cover
  {
    const s = newSlide(p);
    addImage(s, A.sailing, 0, 0, W, H, "Concept visual: professional crew on Lake Garda");
    addShape(s, 0, 0, 590, H, "#071925/88");
    addLine(s, M, 54, 150, C.red, 6);
    addText(s, "HELLY HANSEN · 150TH ANNIVERSARY", M, 78, 520, 26, { fontSize: 15, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "THE\nLOGBOOK", M, 168, 500, 170, { fontSize: 76, bold: true, color: C.white, typeface: FONT_SANS, lineSpacing: 0.82 });
    addText(s, "《航海志》", M, 356, 500, 54, { fontSize: 40, bold: true, color: C.paper, typeface: FONT_SERIF });
    addText(s, "向更远 · 赴新域\n27SS 代言人拍摄整合创意提案", M, 470, 460, 86, { fontSize: 23, color: "#DDE4E5", lineSpacing: 1.25 });
    addText(s, "CREATIVE PITCH · AUG 2026", M, 642, 360, 20, { fontSize: 12, bold: true, color: "#9CB0BA", typeface: FONT_SANS });
    note(s, ["User-provided HH 27SS Ambassador Shooting Brief, pp. 1-60.", "OpenAI ImageGen concept visual; provenance recorded in source-notes.txt."]);
  }

  // 02 — Opening thesis
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    addText(s, "这次不拍一场庆典。", M, 116, 780, 74, { fontSize: 54, bold: true, color: C.navy });
    addText(s, "我们写下下一段 150 年的第一页。", M, 212, 1010, 88, { fontSize: 58, bold: true, color: C.red, lineSpacing: 0.95 });
    addLine(s, M, 348, 420, C.navy, 2);
    addText(s, "让历史不只被回望，而是变成继续出发的证据；\n让王一博带来的关注，最终落到 HH 自身的专业与信任。", M, 382, 920, 100, { fontSize: 25, color: C.sea, lineSpacing: 1.28 });
    addText(s, "THE ASK", 1010, 565, 180, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS, align: "right" });
    addText(s, "赢得的不是一次曝光，\n而是一套可续写的品牌资产。", 760, 596, 430, 64, { fontSize: 22, bold: true, color: C.navy, align: "right", lineSpacing: 1.1 });
    footer(s, 2, false);
    note(s, ["User-provided HH Shooting Brief, pp. 2-5, 12-14.", "HH150周年《航海志》世界观深化创意文档_V2.docx."]);
  }

  // 03 — Business task
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "01 · BRAND TASK", "150周年要回答三个品牌问题", 3, false, "不是再证明代言人有多大声量，而是给市场三个长期相信 HH 的理由。");
    const xs = [64, 448, 832];
    const nums = ["01", "02", "03"];
    const qs = ["HH 是谁？", "为什么值得信任？", "接下来要去哪里？"];
    const as = ["1877 年源于真实海上风浪的专业品牌", "150 年技术、Crew 协作与真实环境验证", "从航海根基出发，进入越野跑与更多新场景"];
    for (let i = 0; i < 3; i++) {
      addText(s, nums[i], xs[i], 256, 290, 40, { fontSize: 18, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, qs[i], xs[i], 306, 320, 54, { fontSize: 31, bold: true, color: C.navy });
      addLine(s, xs[i], 380, 250, C.line, 1);
      addText(s, as[i], xs[i], 408, 304, 128, { fontSize: 21, color: C.sea, lineSpacing: 1.25 });
    }
    note(s, ["User-provided HH Shooting Brief, pp. 3-14.", "智能纪要：品牌150周年拍摄项目规划, pp. 2-7."]);
  }

  // 04 — Strategic shift
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "STRATEGIC SHIFT", "代言人完成破圈，150周年完成信任落点", 4, true);
    addLine(s, 120, 350, 1038, "#35505D", 2);
    addLine(s, 620, 350, 170, C.red, 6);
    addText(s, "看见", 120, 260, 300, 66, { fontSize: 50, bold: true, color: C.white });
    addText(s, "相信", 810, 260, 300, 66, { fontSize: 50, bold: true, color: C.white });
    addText(s, "王一博身上的 HH", 120, 378, 360, 40, { fontSize: 24, color: "#AFC0C8" });
    addText(s, "HH 本身", 810, 378, 360, 40, { fontSize: 24, color: "#AFC0C8" });
    addText(s, "明星流量", 120, 474, 260, 30, { fontSize: 18, bold: true, color: C.red });
    addText(s, "品牌历史 · 专业能力 · 精神认同", 810, 474, 360, 56, { fontSize: 21, bold: true, color: C.paper, lineSpacing: 1.15 });
    addText(s, "我们不让明星消失；我们让他成为观众进入品牌世界的第一位 Crew。", 288, 592, 704, 44, { fontSize: 24, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 4-5.", "HH150周年拍摄项目创意提案.docx, section 1."]);
  }

  // 05 — Strategic answer
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "OUR STRATEGIC ANSWER", "向远，不是交出掌控", 5, false);
    addText(s, "是相信同伴与装备，\n让未知成为可以抵达的前方。", M, 236, 760, 150, { fontSize: 50, bold: true, color: C.navy, lineSpacing: 1.02 });
    addShape(s, 910, 230, 245, 245, C.navy, "none", 0, 123);
    addText(s, "TRUST", 930, 294, 205, 50, { fontSize: 34, bold: true, color: C.white, typeface: FONT_SANS, align: "center", valign: "middle" });
    addText(s, "BY\nPROFESSIONAL", 930, 350, 205, 70, { fontSize: 18, bold: true, color: C.red, typeface: FONT_SANS, align: "center", lineSpacing: 0.95 });
    addText(s, "品牌主张", M, 500, 180, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "#向更远 赴新域#", M, 532, 480, 50, { fontSize: 34, bold: true, color: C.red });
    addText(s, "创意母题", 620, 500, 180, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "专业，让每一次向远都有所托付", 620, 532, 535, 56, { fontSize: 28, bold: true, color: C.navy });
    note(s, ["User-provided HH Shooting Brief, pp. 12-14, 55-57.", "Editorial synthesis based on all user-provided materials."]);
  }

  // 06 — Big idea / logbook
  // 06 — The missing worldview question
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "WHY A WORLDVIEW", "真正缺的不是一句口号，而是一种看世界的方法", 6, false,
      "世界观必须先定义人与自然的关系，再成为影片、平面、人物和场地共同遵守的上位规则。");
    addText(s, "行业惯性", M, 242, 180, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    const oldWords = [
      ["CONQUER", "征服自然"],
      ["ESCAPE", "逃离日常"],
      ["CONSUME", "消费风景"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 294 + i * 82;
      addText(s, oldWords[i][0], M, y, 240, 34, { fontSize: 24, bold: true, color: C.grey2, typeface: FONT_SANS });
      addText(s, oldWords[i][1], 286, y + 2, 230, 32, { fontSize: 21, color: C.grey });
      addLine(s, M, y + 22, 452, C.red, 2);
    }
    addShape(s, 662, 236, 506, 328, C.navy, "none", 0, 24);
    addText(s, "HH 的答案", 704, 270, 180, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "ENTER", 704, 326, 400, 74, { fontSize: 60, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "进入真实世界", 704, 414, 390, 50, { fontSize: 33, bold: true, color: C.paper });
    addText(s, "不把自然变成布景、敌人或奖杯；\n而是学习它的规则，并在其中真实地生活。", 704, 478, 404, 72, { fontSize: 20, color: "#C4D1D6", lineSpacing: 1.25 });
    note(s, ["User-provided HH Shooting Brief, pp. 8-14, 35-42.", "HH150周年《航海志》世界观深化创意文档_V2.docx, sections 1-2."]);
  }

  // 07 — Philosophical root
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "PHILOSOPHICAL ROOT", "Friluftsliv：自由地生活在自然中", 7, true,
      "它不是一句北欧生活方式标签，而是 HH 世界观最可信的文化起点。");
    const roots = [
      ["FRI", "自由", "离开被预设好的路径"],
      ["LUFT", "空气", "重新感受天气与呼吸"],
      ["LIV", "生活", "让户外成为持续发生的日常"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 64 + i * 384;
      addText(s, roots[i][0], x, 246, 330, 62, { fontSize: 48, bold: true, color: i === 1 ? C.red : C.white, typeface: FONT_SANS });
      addText(s, roots[i][1], x, 326, 300, 38, { fontSize: 27, bold: true, color: C.paper });
      addLine(s, x, 382, 300, i === 1 ? C.red : "#38515D", i === 1 ? 4 : 1);
      addText(s, roots[i][2], x, 412, 306, 64, { fontSize: 19, color: "#B8C8CE", lineSpacing: 1.2 });
    }
    addText(s, "对本次提案的意义", M, 520, 260, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "自然不是人类证明自己的舞台；人重新成为自然的一部分。", M, 558, 1120, 46, { fontSize: 29, bold: true, color: C.white });
    addText(s, "这决定了 HH 不讲征服，不拍摆拍，也不把专业简化成更高、更快、更险。", M, 612, 1050, 30, { fontSize: 19, color: "#AFC0C8" });
    note(s, ["HH150周年《航海志》世界观深化创意文档_V2.docx, section 1.", "Friluftsliv is used as a cultural organizing principle rather than a consumer-facing campaign slogan."]);
  }

  // 08 — Worldview anchor
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    addText(s, "HELLY HANSEN WORLDVIEW", M, 58, 360, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "世界不是用来征服的，", M, 140, 1080, 78, { fontSize: 58, bold: true, color: C.navy });
    addText(s, "是用来进入的。", M, 236, 1080, 86, { fontSize: 64, bold: true, color: C.red });
    addLine(s, M, 354, 460, C.navy, 2);
    addText(s,
      "风、浪、雨与岩壁始终按自己的规则发生。\n真正的专业，不是让自然变得容易，\n而是让人看懂环境、信任同伴、依靠装备，\n在真实世界里多停留一会，再向前一步。",
      M, 394, 760, 182, { fontSize: 25, color: C.sea, lineSpacing: 1.34 });
    addShape(s, 900, 388, 244, 244, C.navy, "none", 0, 122);
    addText(s, "ENTER\nTHE REAL\nWORLD", 924, 440, 196, 130, { fontSize: 28, bold: true, color: C.white, typeface: FONT_SANS, align: "center", valign: "middle", lineSpacing: 0.9 });
    footer(s, 8, false);
    note(s, ["Worldview proposition is an original synthesis of the user-provided HH brief and the Friluftsliv development document."]);
  }

  // 09 — Complete philosophy chain
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "THE COMPLETE IDEA CHAIN", "文化根源、品牌方法与人的结果，必须首尾相接", 9, false);
    addLine(s, 148, 394, 984, C.line, 2);
    addShape(s, 120, 366, 56, 56, C.red, "none", 0, 28);
    addShape(s, 612, 366, 56, 56, C.red, "none", 0, 28);
    addShape(s, 1104, 366, 56, 56, C.red, "none", 0, 28);
    const chain = [
      ["FRILUFTSLIV", "人与世界的关系", "进入自然，与它共同生活"],
      ["TRUST BY PROFESSIONAL", "HH 的品牌方法", "用专业让未知变得可进入"],
      ["STAY & FEEL ALIVE", "人的最终状态", "在真实环境中感到生命正在发生"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 64 + i * 410;
      addText(s, chain[i][0], x, 242, 360, 32, { fontSize: 18, bold: true, color: C.red, typeface: FONT_SANS, align: "center" });
      addText(s, chain[i][1], x, 288, 360, 42, { fontSize: 28, bold: true, color: C.navy, align: "center" });
      addText(s, chain[i][2], x, 458, 360, 76, { fontSize: 20, color: C.sea, align: "center", lineSpacing: 1.22 });
    }
    addText(s, "+", 448, 350, 64, 64, { fontSize: 44, bold: true, color: C.navy, typeface: FONT_SANS, align: "center", valign: "middle" });
    addText(s, "→", 942, 350, 64, 64, { fontSize: 44, bold: true, color: C.navy, typeface: FONT_SANS, align: "center", valign: "middle" });
    addShape(s, 228, 570, 824, 54, C.navy, "none", 0, 27);
    addText(s, "年度表达：向更远 · 赴新域", 248, 584, 784, 28, { fontSize: 23, bold: true, color: C.white, align: "center", valign: "middle" });
    note(s, ["User-provided HH Shooting Brief, pp. 54-57.", "HH150周年《航海志》世界观深化创意文档_V2.docx, sections 1-2."]);
  }

  // 10 — Explain the annual line
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "IDEA EXPLANATION", "“向更远 赴新域”因此不再只是扩张口号", 10, false,
      "它描述的是人与世界的关系如何被打开，也是 HH 从航海走向新场景的品牌路径。");
    addText(s, "向更远", M, 246, 420, 60, { fontSize: 44, bold: true, color: C.red });
    addText(s, "不是把距离推得更远，\n而是走出已知经验，进入更大的真实世界。", M, 328, 470, 108, { fontSize: 25, color: C.navy, lineSpacing: 1.25 });
    addVLine(s, 610, 236, 250, C.line, 1);
    addText(s, "赴新域", 680, 246, 420, 60, { fontSize: 44, bold: true, color: C.red });
    addText(s, "既是新的自然环境与运动品类，\n也是新的身体经验、信任关系与自我感受。", 680, 328, 470, 108, { fontSize: 25, color: C.navy, lineSpacing: 1.25 });
    addShape(s, M, 520, 1088, 72, C.navy, "none", 0, 0);
    addText(s, "向远的价值不在抵达了哪里，而在我们因此有能力进入怎样的世界。", 90, 538, 1036, 36, { fontSize: 25, bold: true, color: C.white, align: "center", valign: "middle" });
    note(s, ["User-provided HH Shooting Brief, pp. 10-14.", "Idea interpretation is an original synthesis for this proposal."]);
  }

  // 11 — Worldview laws
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "FIVE WORLDVIEW LAWS", "五条法则，成为所有创意的判断标准", 11, true);
    const laws = [
      ["01", "自然有自己的规则，不是背景", "镜头必须尊重真实天气与地形"],
      ["02", "未知是一种条件，不是敌人", "戏剧来自判断与应变，不来自征服"],
      ["03", "专业是准备，不是逞强", "功能必须在必要动作里被证明"],
      ["04", "远方依靠系统，不依靠英雄", "Crew、装备与经验共同完成旅程"],
      ["05", "旅程留下记录，不留下奖杯", "每一次出发都成为下一次的经验"],
    ];
    addText(s, "LAW", 64, 218, 90, 22, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "BELIEF", 160, 218, 430, 22, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "CREATIVE TEST", 724, 218, 430, 22, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    for (let i = 0; i < laws.length; i++) {
      const y = 260 + i * 70;
      addLine(s, 64, y - 10, 1100, "#2C4652", 1);
      addText(s, laws[i][0], 64, y, 70, 28, { fontSize: 17, bold: true, color: i === 3 ? C.acid : C.red, typeface: FONT_SANS });
      addText(s, laws[i][1], 160, y, 470, 36, { fontSize: 22, bold: true, color: C.white });
      addText(s, laws[i][2], 724, y, 430, 36, { fontSize: 19, color: "#B8C8CE" });
    }
    note(s, ["Five laws are an original creative governance system derived from all user-provided materials."]);
  }

  // 12 — Symbolic system
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "SYMBOLIC SYSTEM", "五个符号，把抽象理念变成可见的世界", 12, false,
      "它们会反复出现在人物动作、产品细节、声音、KV 与传播物料里。");
    addLine(s, 102, 402, 1076, C.line, 2);
    const symbols = [
      ["WIND", "风", "世界的规则"],
      ["CREW", "同伴", "人与人的信任"],
      ["GEAR", "装备", "专业的能力"],
      ["LIGHT", "光", "生命力与继续"],
      ["LOGBOOK", "航海志", "经验被记录与传承"],
    ];
    for (let i = 0; i < symbols.length; i++) {
      const x = 64 + i * 224;
      addShape(s, x + 62, 380, 44, 44, i === 4 ? C.red : C.navy, "none", 0, 22);
      addText(s, symbols[i][0], x, 250, 168, 26, { fontSize: 14, bold: true, color: i === 4 ? C.red : C.sea, typeface: FONT_SANS, align: "center" });
      addText(s, symbols[i][1], x, 296, 168, 40, { fontSize: 28, bold: true, color: C.navy, align: "center" });
      addText(s, symbols[i][2], x, 458, 168, 64, { fontSize: 18, color: C.sea, align: "center", lineSpacing: 1.18 });
    }
    addText(s, "风提出问题 → Crew 作出判断 → 装备提供能力 → 光证明仍然 Alive → 航海志把经验留给下一次出发", 64, 582, 1120, 40, { fontSize: 20, bold: true, color: C.red, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 19-28, 54-57.", "HH150周年《航海志》世界观深化创意文档_V2.docx, sections 3, 5 and 10."]);
  }

  // 13 — Narrative device / logbook
  {
    const s = newSlide(p);
    addImage(s, A.logbook, 0, 0, W, H, "Concept visual: the new expedition logbook");
    addShape(s, 0, 0, 565, H, "#071925/90");
    addTag(s, "NARRATIVE DEVICE", M, 62, { width: 190 });
    addText(s, "THE\nLOGBOOK", M, 136, 450, 150, { fontSize: 65, bold: true, color: C.white, typeface: FONT_SANS, lineSpacing: 0.84 });
    addText(s, "《航海志》", M, 308, 430, 52, { fontSize: 38, bold: true, color: C.paper, typeface: FONT_SERIF });
    addLine(s, M, 390, 330, C.red, 4);
    addText(s, "世界观回答我们相信什么；《航海志》负责让它被看见。\n一本在 150 周年正式启用、未来可由真实 Crew 持续续写的远行记录。", M, 420, 440, 116, { fontSize: 21, color: "#D3DEE2", lineSpacing: 1.28 });
    addText(s, "CHAPTER 01 · SEA\nCHAPTER 02 · TRAIL\nNEXT CHAPTER · OPEN", M, 578, 340, 62, { fontSize: 15, bold: true, color: "#8FA6B0", typeface: FONT_SANS, lineSpacing: 1.25 });
    note(s, ["HH150周年《航海志》世界观深化创意文档_V2.docx.", "OpenAI ImageGen concept visual; provenance recorded in source-notes.txt."]);
  }

  // 07 — Platform mechanics
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "PLATFORM MECHANIC", "每一页都必须记录四件真实发生的事", 7, false, "环境不是布景，产品不是陈列，人物不是摆拍。");
    const labels = ["WHERE", "WHO", "DECISION", "PROOF"];
    const big = ["真实环境", "真实 Crew", "真实判断", "真实产品证据"];
    const sub = ["风、浪、岩壁、溪流", "专业者与新加入者平等协作", "等待、应变、取舍与继续", "功能在必要动作中自然出现"];
    const colors = [C.navy, C.sea, "#456A64", C.red];
    for (let i = 0; i < 4; i++) {
      const x = M + i * 290;
      addShape(s, x, 256, 255, 286, i === 3 ? "#FFF4F2" : C.paper, "none", 0, 0);
      addText(s, labels[i], x + 22, 280, 210, 24, { fontSize: 13, bold: true, color: colors[i], typeface: FONT_SANS });
      addText(s, String(i + 1).padStart(2, "0"), x + 22, 326, 96, 56, { fontSize: 42, bold: true, color: colors[i], typeface: FONT_SANS });
      addText(s, big[i], x + 22, 396, 212, 48, { fontSize: 27, bold: true, color: C.navy });
      addText(s, sub[i], x + 22, 462, 212, 66, { fontSize: 18, color: C.sea, lineSpacing: 1.2 });
    }
    note(s, ["HH150周年《航海志》世界观深化创意文档_V2.docx.", "Editorial synthesis based on user-provided brief."]);
  }

  // 08 — Architecture
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "ANNUAL ARCHITECTURE", "同一本《航海志》，两章解决两件不同的事", 8, true);
    addLine(s, 180, 396, 920, "#47606C", 2);
    addLine(s, 610, 396, 60, C.red, 7);
    addShape(s, 122, 278, 356, 230, "#102A38", "#35505D", 1, 22);
    addShape(s, 802, 278, 356, 230, "#102A38", "#35505D", 1, 22);
    addShape(s, 558, 330, 164, 132, C.red, "none", 0, 66);
    addText(s, "01", 150, 300, 60, 44, { fontSize: 30, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "SAILING", 150, 350, 280, 42, { fontSize: 32, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "讲人和人\n建立航海专业根基与 Crew 信任", 150, 414, 278, 68, { fontSize: 20, color: "#B9C7CD", lineSpacing: 1.2 });
    addText(s, "02", 830, 300, 60, 44, { fontSize: 30, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "TRAIL", 830, 350, 280, 42, { fontSize: 32, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "讲人和装备\n建立 Deer Racer 专业产品认知", 830, 414, 278, 68, { fontSize: 20, color: "#B9C7CD", lineSpacing: 1.2 });
    addText(s, "向更远\n赴新域", 578, 362, 124, 70, { fontSize: 25, bold: true, color: C.white, align: "center", valign: "middle", lineSpacing: 0.95 });
    addText(s, "海上形成专业 → 山野验证未来", 390, 572, 500, 42, { fontSize: 28, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 15-20, 26-29, 47-52."]);
  }

  // 09 — Ambassador role
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "AMBASSADOR ROLE", "王一博不是被削弱，而是被赋予更有用的角色", 9, false);
    addText(s, "第一位加入的 Crew", M, 238, 600, 66, { fontSize: 46, bold: true, color: C.red });
    addText(s, "他替观众进入陌生环境：学习、判断、信任，然后继续。", M, 316, 650, 46, { fontSize: 23, color: C.sea });
    const rules = [
      ["看得见", "核心时刻与关键平面保证明星识别度"],
      ["不独占", "不给虚假的单人英雄叙事，和真实专业者协作"],
      ["有变化", "从新 Crew 到能作出判断的人，完成角色弧光"],
      ["能转化", "单人素材仍为 Social / 电商 / 门店留足版本"],
    ];
    for (let i = 0; i < 4; i++) {
      const y = 226 + i * 98;
      addText(s, String(i + 1).padStart(2, "0"), 780, y, 50, 30, { fontSize: 17, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, rules[i][0], 844, y, 130, 34, { fontSize: 24, bold: true, color: C.navy });
      addText(s, rules[i][1], 844, y + 38, 345, 46, { fontSize: 18, color: C.sea, lineSpacing: 1.12 });
    }
    note(s, ["User-provided HH Shooting Brief, pp. 2-5, 17, 19.", "HH150周年《航海志》世界观深化创意文档_V2.docx, sections 4 and 9."]);
  }

  // 10 — Visual language
  {
    const s = newSlide(p);
    addImage(s, A.hands, 0, 0, W, H, "Concept visual: real crew hands working a wet winch");
    addShape(s, 0, 0, 570, H, "#071925/92");
    addTag(s, "TONE & MANNER", M, 54, { width: 180 });
    addText(s, "QUIET\nCONFIDENCE", M, 128, 440, 142, { fontSize: 58, bold: true, color: C.white, typeface: FONT_SANS, lineSpacing: 0.88 });
    addText(s, "安静的专业", M, 286, 400, 44, { fontSize: 31, bold: true, color: C.paper });
    const items = ["真实天气", "真实动作", "克制构图", "使用痕迹"];
    const defs = ["风、雨、湿度都留在画面里", "身体在解决问题，不在摆姿势", "少而好，让环境与产品自己说话", "不是奢华道具，而是长期专业积累"];
    for (let i = 0; i < 4; i++) {
      const y = 374 + i * 62;
      addText(s, items[i], M, y, 126, 26, { fontSize: 19, bold: true, color: C.red });
      addText(s, defs[i], 208, y, 300, 38, { fontSize: 17, color: "#C8D3D7" });
    }
    note(s, ["User-provided HH Shooting Brief, pp. 25, 33-42.", "OpenAI ImageGen concept visual; provenance recorded in source-notes.txt."]);
  }

  // 11 — Film 01 cover
  {
    const s = newSlide(p);
    addImage(s, A.sailing, 0, 0, W, H, "Concept visual: sailing chapter");
    addShape(s, 0, 0, W, H, "#071925/34");
    addText(s, "CHAPTER 01", M, 62, 260, 24, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "CREW", M, 116, 520, 100, { fontSize: 82, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "以海为根 · 向远而生", M, 230, 540, 52, { fontSize: 34, bold: true, color: C.white });
    addLine(s, M, 305, 300, C.red, 5);
    addText(s, "150周年品牌精神微电影\n一次真实航行，讲清 HH 的专业从何而来。", M, 340, 500, 90, { fontSize: 23, color: "#E2E8EA", lineSpacing: 1.24 });
    addText(s, "FILM 01 · SAILING", 930, 648, 270, 24, { fontSize: 14, bold: true, color: C.white, typeface: FONT_SANS, align: "right" });
    note(s, ["User-provided HH Shooting Brief, pp. 19-25.", "OpenAI ImageGen concept visual; provenance recorded in source-notes.txt."]);
  }

  // 12 — Film 01 premise
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "FILM 01 · CORE DRAMA", "风不认船长，只认 Crew", 12, false, "当风向改变，任何人都不能单独把船带远。");
    addText(s, "不是“王一博驾驶帆船”的明星大片，\n而是他加入一支真实 Crew 后，第一次把判断交给共同经验。", M, 240, 690, 120, { fontSize: 31, bold: true, color: C.navy, lineSpacing: 1.18 });
    addShape(s, 832, 236, 344, 280, C.navy, "none", 0, 22);
    addText(s, "人物关系", 864, 268, 140, 28, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "新 Crew", 864, 324, 200, 40, { fontSize: 31, bold: true, color: C.white });
    addText(s, "被纠正 → 学会观察 → 作出判断 → 获得托付", 864, 382, 270, 88, { fontSize: 21, color: "#C6D2D7", lineSpacing: 1.25 });
    addText(s, "品牌证明", M, 448, 160, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "HH 的专业不是一个人的英勇，\n而是一套让每个人都能继续前行的系统。", M, 488, 660, 88, { fontSize: 25, color: C.sea, lineSpacing: 1.2 });
    note(s, ["User-provided HH Shooting Brief, pp. 19-20.", "HH150周年《航海志》世界观深化创意文档_V2.docx, sections 4-6.", "Film line is original proposal copy, not presented as a historical quotation."]);
  }

  // 13 — Film 01 story beats
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "FILM 01 · STORY", "让真实风况成为叙事时钟", 13, false, "所有戏剧都从环境与协作里发生，不靠旁白制造宏大。");
    const beats = [
      ["01", "识风", "清晨整装。新 Crew 学会先观察风与水，再决定出发。", "静 / 生涩"],
      ["02", "等待", "风窗未到。沉默准备、检查装备，让专业感先于动作出现。", "克制 / 专注"],
      ["03", "应变", "风向改变，全员调帆。第一次看见共同判断如何发生。", "紧张 / 协作"],
      ["04", "归港", "天色转暗，纹样微光出现。不是胜利，是下一页被写下。", "平静 / 余韵"],
    ];
    for (let i = 0; i < 4; i++) {
      const x = M + i * 290;
      addText(s, beats[i][0], x, 250, 50, 34, { fontSize: 18, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, beats[i][1], x, 300, 230, 42, { fontSize: 30, bold: true, color: C.navy });
      addLine(s, x, 358, 230, i === 2 ? C.red : C.line, i === 2 ? 4 : 1);
      addText(s, beats[i][2], x, 386, 235, 116, { fontSize: 19, color: C.sea, lineSpacing: 1.22 });
      addText(s, beats[i][3], x, 528, 220, 26, { fontSize: 14, bold: true, color: i === 2 ? C.red : C.grey, typeface: FONT_SANS });
    }
    note(s, ["HH150周年《航海志》世界观深化创意文档_V2.docx, sections 5-6.", "User-provided HH Shooting Brief, pp. 19-20, 44."]);
  }

  // 14 — Film 01 product translation
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "FILM 01 · PRODUCT PROOF", "功能不做参数字卡，全部藏进必要动作", 14, true);
    const rows = [
      ["分区保暖", "清晨等待风窗时保持动作灵活与身体稳定", "等待"],
      ["HELLY TECH 防泼水", "浪花打上甲板，Crew 继续调帆而不被天气打断", "应变"],
      ["海浪 / 北极星夜光", "日照吸光，归港时在群像中低调显影", "归港"],
      ["CREW 命名", "每位船员承担一个不可替代的岗位", "全片"],
    ];
    addText(s, "PRODUCT", 72, 240, 220, 24, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "ACTION", 390, 240, 460, 24, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "BEAT", 1010, 240, 140, 24, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    for (let i = 0; i < rows.length; i++) {
      const y = 286 + i * 78;
      addLine(s, 72, y - 14, 1090, "#2C4652", 1);
      addText(s, rows[i][0], 72, y, 270, 34, { fontSize: 22, bold: true, color: C.white });
      addText(s, rows[i][1], 390, y, 530, 50, { fontSize: 19, color: "#C8D4D8", lineSpacing: 1.15 });
      addText(s, rows[i][2], 1010, y, 140, 32, { fontSize: 18, bold: true, color: i === 2 ? C.acid : C.paper, align: "right" });
    }
    addText(s, "仅露出艺术家联名 Crew Jacket；产品细节以官方样衣与品牌审批为准。", 72, 616, 960, 28, { fontSize: 16, color: "#92A7B0" });
    note(s, ["User-provided HH Shooting Brief, pp. 20, 22.", "智能纪要：品牌150周年拍摄项目规划, pp. 2-3, 6-7."]);
  }

  // 15 — Nightglow signature
  {
    const s = newSlide(p);
    addImage(s, A.night, 0, 0, W, H, "Concept visual: crew returning at blue hour with subtle nightglow textile");
    addShape(s, 0, 0, 560, H, "#071925/82");
    addText(s, "SIGNATURE FRAME", M, 66, 240, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "白天吸收湖面上的光，\n入夜把光还给 Crew。", M, 132, 450, 132, { fontSize: 44, bold: true, color: C.white, lineSpacing: 1.05 });
    addLine(s, M, 310, 250, C.red, 4);
    addText(s, "不是科技炫技，而是一次从功能到情感的自然转译。\n这将成为航海篇最可识别、最可获奖、也最能延展 KV 的画面。", M, 342, 430, 126, { fontSize: 21, color: "#D2DDE1", lineSpacing: 1.26 });
    addText(s, "视觉方向示意 · 非最终成片", M, 622, 330, 22, { fontSize: 13, bold: true, color: "#91A6AF" });
    note(s, ["HH150周年《航海志》世界观深化创意文档_V2.docx, section 5.", "OpenAI ImageGen concept visual; product treatment and talent likeness subject to official approvals."]);
  }

  // 16 — Film 01 output
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "FILM 01 · RELEASE SYSTEM", "把一次发布，做成 150 周年上映事件", 16, false);
    const cols = [
      ["TEASER", "制造悬念", "风、绳索、纹样、Crew 名单\n15s / 6s / 静帧"],
      ["PREMIERE", "正式上映", "品牌微电影完整版\n主 KV / 红毯与首映空间内容"],
      ["SUSTAIN", "续写 150 年", "Crew 人物短片\n航海纪录 / 艺术家纹样 / 产品切片"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 64 + i * 386;
      addText(s, cols[i][0], x, 250, 300, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, cols[i][1], x, 296, 320, 46, { fontSize: 31, bold: true, color: C.navy });
      addLine(s, x, 356, 320, i === 1 ? C.red : C.line, i === 1 ? 4 : 1);
      addText(s, cols[i][2], x, 388, 320, 110, { fontSize: 20, color: C.sea, lineSpacing: 1.25 });
    }
    addShape(s, 64, 558, 1120, 58, C.navy, "none", 0, 0);
    addText(s, "评奖目标：影片完成度按品牌电影，而不是常规产品 TVC 的标准建立。", 90, 574, 1068, 28, { fontSize: 21, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 47-49.", "智能纪要：品牌150周年拍摄项目规划, p. 3."]);
  }

  // 17 — Film 02 cover
  {
    const s = newSlide(p);
    addImage(s, A.trail, 0, 0, W, H, "Concept visual: trail chapter in the Dolomites");
    addShape(s, 0, 0, W, H, "#071925/30");
    addText(s, "CHAPTER 02", M, 62, 260, 24, { fontSize: 15, bold: true, color: C.acid, typeface: FONT_SANS });
    addText(s, "DEAR\nRACER", M, 116, 520, 164, { fontSize: 76, bold: true, color: C.white, typeface: FONT_SANS, lineSpacing: 0.84 });
    addText(s, "情寄山野 · 跑探新境", M, 318, 540, 50, { fontSize: 33, bold: true, color: C.white });
    addLine(s, M, 392, 300, C.acid, 5);
    addText(s, "Deer Racer 核心产品片\n让每一段地形，都替产品回信。", M, 426, 500, 80, { fontSize: 23, color: "#E2E8EA", lineSpacing: 1.22 });
    addText(s, "FILM 02 · TRAIL RUNNING", 890, 648, 310, 24, { fontSize: 14, bold: true, color: C.white, typeface: FONT_SANS, align: "right" });
    note(s, ["User-provided HH Shooting Brief, pp. 26-33.", "OpenAI ImageGen concept visual; provenance recorded in source-notes.txt."]);
  }

  // 18 — Film 02 idea
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "FILM 02 · CREATIVE IDEA", "“写给每位 Racer 的信”，由地形来写", 18, false);
    addText(s, "鞋不自我介绍。\n它在每一次落地、转向、上坡和湿滑中，留下答案。", M, 230, 720, 118, { fontSize: 36, bold: true, color: C.navy, lineSpacing: 1.15 });
    addShape(s, 838, 226, 342, 296, C.navy, "none", 0, 22);
    addText(s, "DEAR RACER,", 870, 260, 280, 34, { fontSize: 22, bold: true, color: C.acid, typeface: FONT_SANS });
    addText(s, "你不必先相信参数。\n\n让碎石、湿地、陡坡和夜色，\n一一替它作证。", 870, 320, 274, 170, { fontSize: 22, color: C.white, lineSpacing: 1.2 });
    addText(s, "叙事主语：跑者的身体感受\n视觉主语：鞋与地形的接触", M, 474, 620, 74, { fontSize: 23, color: C.sea, lineSpacing: 1.3 });
    note(s, ["User-provided HH Shooting Brief, pp. 26-28.", "智能纪要：品牌150周年拍摄项目规划, pp. 2-3."]);
  }

  // 19 — Four proofs
  {
    const s = newSlide(p);
    addImage(s, A.shoe, 0, 0, W, H, "Concept visual: trail shoe contacting wet limestone");
    addShape(s, 0, 0, 620, H, "#071925/88");
    addText(s, "FOUR PROOFS", M, 60, 280, 24, { fontSize: 14, bold: true, color: C.acid, typeface: FONT_SANS });
    addText(s, "轻 · 弹 · 稳 · 抓", M, 110, 500, 64, { fontSize: 48, bold: true, color: C.white });
    const rows = [
      ["轻", "长上坡后仍能保持步频"],
      ["弹", "落地后迅速回到下一步"],
      ["稳", "横切与急转中锁定脚步"],
      ["抓", "湿岩、碎石与坡面建立牵引"],
    ];
    for (let i = 0; i < 4; i++) {
      const y = 224 + i * 82;
      addText(s, rows[i][0], M, y, 50, 42, { fontSize: 29, bold: true, color: C.acid });
      addText(s, rows[i][1], 142, y + 2, 390, 50, { fontSize: 20, color: "#D3DEE2", lineSpacing: 1.15 });
      addLine(s, 142, y + 60, 340, "#405763", 1);
    }
    addText(s, "视觉方向示意 · 鞋型以官方样品为准", M, 620, 420, 22, { fontSize: 13, bold: true, color: "#93A7B0" });
    note(s, ["User-provided HH Shooting Brief, pp. 27-28.", "OpenAI ImageGen concept visual; prototype form is illustrative and subject to official product assets."]);
  }

  // 20 — Camera logic
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "FILM 02 · CAMERA LOGIC", "先拍鞋与地形，再拍人物与情绪", 20, false, "把专业性拍成身体可以感到的证据，而不是实验室解释。");
    const left = [
      ["脚步优先", "落地、蹬伸、刹停、转向"],
      ["路面分层", "湿岩 / 碎石 / 林道 / 坡面"],
      ["同位替模", "高难度段落保证动作专业与艺人安全"],
    ];
    for (let i = 0; i < left.length; i++) {
      const y = 246 + i * 108;
      addText(s, String(i + 1).padStart(2, "0"), M, y, 46, 28, { fontSize: 16, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, left[i][0], 126, y, 200, 38, { fontSize: 27, bold: true, color: C.navy });
      addText(s, left[i][1], 126, y + 44, 350, 40, { fontSize: 19, color: C.sea });
    }
    addShape(s, 610, 236, 556, 330, C.paper, "none", 0, 22);
    addText(s, "镜头占比建议", 642, 270, 240, 30, { fontSize: 21, bold: true, color: C.navy });
    const bars = [
      ["鞋 × 地形", 0.46, C.red],
      ["动作 × 身体", 0.34, C.sea],
      ["人物 × 情绪", 0.20, C.grey2],
    ];
    for (let i = 0; i < bars.length; i++) {
      const y = 332 + i * 72;
      addText(s, bars[i][0], 642, y, 180, 28, { fontSize: 18, bold: true, color: C.ink });
      addShape(s, 820, y + 3, 300, 20, "#D9DEDC", "none", 0, 10);
      addShape(s, 820, y + 3, 300 * bars[i][1], 20, bars[i][2], "none", 0, 10);
      addText(s, `${Math.round(bars[i][1] * 100)}%`, 1084, y - 2, 48, 28, { fontSize: 15, bold: true, color: C.navy, typeface: FONT_SANS, align: "right" });
    }
    addText(s, "比例为创意剪辑方向，最终随导演脚本与素材质量调整。", 642, 516, 470, 28, { fontSize: 15, color: C.grey });
    note(s, ["User-provided HH Shooting Brief, pp. 27-28.", "智能纪要：品牌150周年拍摄项目规划, pp. 2-3, 6-7.", "Shot-ratio recommendation is an original creative proposal assumption."]);
  }

  // 21 — Film 02 release system
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "FILM 02 · RELEASE SYSTEM", "把产品首发变成一段从银幕到赛道的证明", 21, true);
    const phases = [
      ["TEASER", "一句开场 + 四种地形", "上市前 3-4 周"],
      ["LAUNCH", "完整产品片 + 王一博发起", "2027 年 4 月"],
      ["PROVE", "真实跑者试穿 / 专业媒体 / 赛事", "持续验证"],
      ["SUSTAIN", "四卖点切片 / 门店 / 社群回信", "长效沉淀"],
    ];
    addLine(s, 110, 404, 1054, "#3A515C", 2);
    for (let i = 0; i < 4; i++) {
      const x = 88 + i * 292;
      addShape(s, x, 376, 48, 48, i === 1 ? C.acid : C.red, "none", 0, 24);
      addText(s, String(i + 1), x, 386, 48, 24, { fontSize: 16, bold: true, color: i === 1 ? C.navy : C.white, typeface: FONT_SANS, align: "center", valign: "middle" });
      addText(s, phases[i][0], x, 252, 210, 30, { fontSize: 15, bold: true, color: i === 1 ? C.acid : C.red, typeface: FONT_SANS });
      addText(s, phases[i][1], x, 296, 230, 60, { fontSize: 22, bold: true, color: C.white, lineSpacing: 1.15 });
      addText(s, phases[i][2], x, 456, 210, 30, { fontSize: 16, color: "#A7BAC2" });
    }
    addText(s, "传播目标：王一博建立第一印象，真实 Racer 完成专业证明。", M, 580, 1120, 44, { fontSize: 25, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 50-52."]);
  }

  // 22 — Eight looks overview
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "STILLS SYSTEM", "8套 Look 不做八次换装，而做一本完整场景志", 22, false);
    addImage(s, A.brief16, 688, 218, 490, 300, "Source product matrix from client brief", { fit: "cover", crop: { left: 0.02, top: 0.05, right: 0.02, bottom: 0.04 }, radius: 18 });
    addText(s, "SEA · 3 LOOKS", M, 238, 380, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "01  Sailing Heritage\n04  H2ALIVE\n07  HH Stotte", M, 278, 430, 122, { fontSize: 27, bold: true, color: C.navy, lineSpacing: 1.35 });
    addText(s, "TRAIL · 5 LOOKS", M, 446, 380, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "02 / 03  徒步\n05 / 08  越野跑\n06  溪流", M, 486, 430, 110, { fontSize: 25, bold: true, color: C.navy, lineSpacing: 1.3 });
    addText(s, "同一视觉语法：真实环境、克制机能、安静自信；\n不同产品证据：每套 Look 只解决一个首要沟通点。", 688, 548, 490, 66, { fontSize: 19, color: C.sea, lineSpacing: 1.2 });
    note(s, ["User-provided HH Shooting Brief, p. 16.", "User-provided source slide shown as a compact reference matrix."]);
  }

  // 23 — Sailing looks
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "STILLS · SEA", "三套航海 Look，三种距离品牌历史的方式", 23, false);
    const cols = [
      ["LOOK 01", "SAILING HERITAGE", "艺术家联名 Crew Jacket", "群像协作 · 夜光纹样 · 防泼水", "品牌根基"],
      ["LOOK 04", "H2ALIVE", "SALT 航海夹克", "赛事级防护 · 海岸街头 · 动态混搭", "经典重塑"],
      ["LOOK 07", "HH STOTTE", "航海生活方式", "双向拉链 · 多口袋 · UPF50+", "日常延伸"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 64 + i * 386;
      addText(s, cols[i][0], x, 230, 160, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, cols[i][1], x, 274, 330, 38, { fontSize: 27, bold: true, color: C.navy, typeface: FONT_SANS });
      addText(s, cols[i][2], x, 328, 320, 40, { fontSize: 21, bold: true, color: C.sea });
      addLine(s, x, 386, 320, i === 0 ? C.red : C.line, i === 0 ? 4 : 1);
      addText(s, cols[i][3], x, 414, 320, 70, { fontSize: 19, color: C.ink, lineSpacing: 1.2 });
      addText(s, cols[i][4], x, 526, 170, 30, { fontSize: 17, bold: true, color: C.red });
    }
    addShape(s, 64, 590, 1120, 38, C.navy);
    addText(s, "拍法：不是码头度假；是准备、操控、抵御与收工之后留下的真实痕迹。", 86, 598, 1076, 24, { fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
    note(s, ["User-provided HH Shooting Brief, pp. 21-25, 35-41."]);
  }

  // 24 — Outdoor looks
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "STILLS · OUTDOOR", "五套山野 Look，用天气与路面建立差异", 24, false);
    const rows = [
      ["LOOK 02", "HIKING", "防风 / 山系探索", "抓住 Weather Window"],
      ["LOOK 03", "HIKING", "层搭 / 真实天气变化", "走出完美天气"],
      ["LOOK 05", "TRAIL", "轻量 / 夜跑 / 可打包", "Healing in Bergen"],
      ["LOOK 06", "CREEK", "防晒 / 快干 / 溪流", "Micro Adventure"],
      ["LOOK 08", "TRAIL", "鞋服完整专业场景", "Deer Racer Hero"],
    ];
    addText(s, "LOOK", 64, 230, 130, 20, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "SCENE", 228, 230, 160, 20, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "PRODUCT PROOF", 470, 230, 260, 20, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "VISUAL HOOK", 860, 230, 280, 20, { fontSize: 13, bold: true, color: C.red, typeface: FONT_SANS });
    for (let i = 0; i < rows.length; i++) {
      const y = 272 + i * 70;
      addLine(s, 64, y - 10, 1120, C.line, 1);
      addText(s, rows[i][0], 64, y, 130, 28, { fontSize: 18, bold: true, color: C.navy, typeface: FONT_SANS });
      addText(s, rows[i][1], 228, y, 170, 28, { fontSize: 19, bold: true, color: i >= 2 ? C.red : C.sea, typeface: FONT_SANS });
      addText(s, rows[i][2], 470, y, 330, 34, { fontSize: 19, color: C.ink });
      addText(s, rows[i][3], 860, y, 300, 34, { fontSize: 19, bold: true, color: C.navy });
    }
    addText(s, "红线：不拍旅游明信片，不拍无运动逻辑的松弛摆姿；每个画面先成立为真实户外，再成立为时尚。", 64, 612, 1080, 30, { fontSize: 18, bold: true, color: C.red });
    note(s, ["User-provided HH Shooting Brief, pp. 29-33, 35-42."]);
  }

  // 25 — 40-frame system
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "CAPTURE SYSTEM", "每套 5 张，不是重复构图，而是完成一次传播闭环", 25, true);
    const shots = [
      ["01", "HERO", "人物 × 环境\n主 KV / POP"],
      ["02", "ACTION", "真实运动瞬间\nSocial / PR"],
      ["03", "FUNCTION", "功能在动作中出现\n电商 / 详情页"],
      ["04", "DETAIL", "面料 / 鞋 / 配件\n切片 / 门店"],
      ["05", "PORTRAIT", "人物情绪与识别度\n粉丝 / CRM"],
    ];
    for (let i = 0; i < 5; i++) {
      const x = 64 + i * 224;
      addText(s, shots[i][0], x, 240, 50, 26, { fontSize: 15, bold: true, color: i === 0 ? C.acid : C.red, typeface: FONT_SANS });
      addText(s, shots[i][1], x, 282, 200, 34, { fontSize: 23, bold: true, color: C.white, typeface: FONT_SANS });
      addLine(s, x, 336, 180, i === 0 ? C.acid : "#3B5360", i === 0 ? 4 : 1);
      addText(s, shots[i][2], x, 362, 190, 82, { fontSize: 18, color: "#C4D1D6", lineSpacing: 1.22 });
    }
    addText(s, "8 LOOKS × 5 SHOTS = 40 CORE STILLS", 64, 520, 840, 50, { fontSize: 38, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "横版、竖版与安全裁切区在取景时一次完成，避免上线阶段靠后期硬裁。", 64, 584, 900, 34, { fontSize: 20, color: "#AFC0C8" });
    note(s, ["User-provided HH Shooting Brief, p. 16.", "智能纪要：品牌150周年拍摄项目规划, p. 2.", "Five-shot taxonomy is an original production recommendation."]);
  }

  // 26 — Locations
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "LOCATION AS STORY", "两个场地不是风景替换，而是两种专业证据", 26, false);
    addShape(s, 64, 222, 536, 350, C.paper, "none", 0, 20);
    addShape(s, 680, 222, 536, 350, "#E8ECEA", "none", 0, 20);
    addText(s, "LAKE GARDA", 92, 250, 270, 28, { fontSize: 17, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "风，迫使 Crew 作出判断", 92, 298, 430, 48, { fontSize: 30, bold: true, color: C.navy });
    addText(s, "真实 M32 操控\n码头准备与 Crew 分工\n山、湖、风况与蓝调时刻\n红线：绝不拍成欧洲度假胜地", 92, 372, 430, 140, { fontSize: 20, color: C.sea, lineSpacing: 1.35 });
    addText(s, "DOLOMITES", 708, 250, 270, 28, { fontSize: 17, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "地形，迫使装备给出答案", 708, 298, 440, 48, { fontSize: 30, bold: true, color: C.navy });
    addText(s, "连续动作与技术坡面\n湿岩、碎石、林道、溪流切换\n晨昏与低光反光测试\n红线：绝不拍成旅游宣传片", 708, 372, 430, 140, { fontSize: 20, color: C.sea, lineSpacing: 1.35 });
    addText(s, "场地确认前必须完成：风况 / 船位 / 许可 / 日照 / 天气备选 / 医疗与救援线路实勘。", 64, 614, 1120, 30, { fontSize: 18, bold: true, color: C.red, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 43-45.", "智能纪要：品牌150周年拍摄项目规划, pp. 3, 5-7."]);
  }

  // 27 — Two talent days
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "TWO TALENT DAYS", "用双机组并行，把 16 小时留给真正不可替代的镜头", 27, false);
    const dayData = [
      ["DAY 01", "GARDA · SAILING", "A组：Film 01 主叙事 + Look 01\nB组：船体 / 产品 / 空镜 / 替身走位\n穿插：Look 04 / 07 核心艺人镜头\n收口：蓝调归港 + 夜光纹样"],
      ["DAY 02", "DOLOMITES · OUTDOOR", "A组：Film 02 艺人可控跑段 + Hero Portrait\nB组：替模高难度动作 + 产品微距\n穿插：Look 02 / 03 / 05 / 06 / 08\n收口：低光跑动 + 反光细节"],
    ];
    for (let i = 0; i < 2; i++) {
      const x = 64 + i * 576;
      addText(s, dayData[i][0], x, 236, 160, 28, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, dayData[i][1], x, 282, 470, 42, { fontSize: 29, bold: true, color: C.navy, typeface: FONT_SANS });
      addLine(s, x, 344, 500, i === 0 ? C.sea : C.red, 4);
      addText(s, dayData[i][2], x, 374, 490, 186, { fontSize: 20, color: C.sea, lineSpacing: 1.42 });
    }
    addShape(s, 64, 584, 1076, 44, C.navy);
    addText(s, "原则：艺人只做他不可替代的识别与情绪；替模、产品、空镜和技术动作全部提前预演。", 84, 594, 1036, 24, { fontSize: 18, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 43-45.", "智能纪要：品牌150周年拍摄项目规划, pp. 3, 5-7.", "Day structure is an original feasibility recommendation; exact call sheets remain TBC after location recce and product fitting."]);
  }

  // 28 — KV layout demo
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "KV LAYOUT DEMO", "同一主视觉，一次完成横版、方版与竖版安全构图", 28, true, "以下为版式逻辑示意；最终人物、产品与联名纹样以官方素材为准。");
    // landscape
    addImage(s, A.sailing, 64, 242, 500, 282, "Landscape KV mockup", { radius: 16 });
    addShape(s, 64, 242, 230, 282, "#071925/82", "none", 0, 16);
    addText(s, "向更远\n赴新域", 84, 278, 190, 82, { fontSize: 32, bold: true, color: C.white, lineSpacing: 0.95 });
    addText(s, "THE LOGBOOK", 84, 452, 180, 22, { fontSize: 12, bold: true, color: C.red, typeface: FONT_SANS });
    // square
    addImage(s, A.night, 612, 242, 282, 282, "Square KV mockup", { radius: 16, crop: { left: 0.22, top: 0, right: 0.22, bottom: 0 } });
    addShape(s, 612, 402, 282, 122, "#071925/78", "none", 0, 16);
    addText(s, "CREW", 632, 420, 242, 42, { fontSize: 28, bold: true, color: C.white, typeface: FONT_SANS });
    addText(s, "以海为根 · 向远而生", 632, 470, 242, 24, { fontSize: 15, bold: true, color: C.red });
    // vertical
    addImage(s, A.trail, 944, 222, 208, 330, "Vertical KV mockup", { radius: 16, crop: { left: 0.28, top: 0, right: 0.28, bottom: 0 } });
    addShape(s, 944, 222, 208, 130, "#071925/78", "none", 0, 16);
    addText(s, "DEAR\nRACER", 964, 244, 168, 76, { fontSize: 28, bold: true, color: C.white, typeface: FONT_SANS, lineSpacing: 0.88 });
    addText(s, "16:9", 64, 554, 100, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "1:1 / 4:5", 612, 554, 120, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "9:16", 944, 574, 100, 24, { fontSize: 14, bold: true, color: C.red, typeface: FONT_SANS });
    note(s, ["User-provided HH Shooting Brief, p. 43 (KV layout demo requirement).", "OpenAI ImageGen concept visuals; provenance recorded in source-notes.txt."]);
  }

  // 29 — Content engine
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "CONTENT ENGINE", "一次拍摄，留下两条可以持续一年的内容链", 29, false);
    addText(s, "SAILING", 64, 224, 180, 28, { fontSize: 17, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "150周年首映", 64, 266, 250, 38, { fontSize: 28, bold: true, color: C.navy });
    addText(s, "Teaser → 品牌微电影 → Crew纪录 → 艺术家故事 → 门店放映", 64, 320, 1080, 40, { fontSize: 22, color: C.sea });
    addLine(s, 64, 388, 1120, C.line, 1);
    addText(s, "TRAIL", 64, 424, 180, 28, { fontSize: 17, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "Deer Racer 首发", 64, 466, 290, 38, { fontSize: 28, bold: true, color: C.navy });
    addText(s, "一句来信 → 产品片 → 四卖点切片 → 专业跑者实证 → 赛事 / 门店 / 社群", 64, 520, 1080, 40, { fontSize: 22, color: C.sea });
    addShape(s, 64, 594, 1120, 38, C.red);
    addText(s, "共同资产：THE LOGBOOK 每年开启新章节，让 150 周年之后仍有内容可写。", 84, 602, 1080, 24, { fontSize: 18, bold: true, color: C.white, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 47-52.", "HH150周年《航海志》世界观深化创意文档_V2.docx, section 3."]);
  }

  // 30 — Deliverables
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "DELIVERABLE ECOSYSTEM", "交付不只按文件计数，更按传播场景反推", 30, false);
    metric(s, "40", "张 8 Look 核心平面\n横竖版安全构图", 64, 240, 230, false, C.red);
    metric(s, "02", "支核心影片\n品牌精神 + 产品认知", 334, 240, 240, false, C.red);
    metric(s, "3×", "Teaser / Launch / Sustain\n多阶段剪辑架构", 624, 240, 260, false, C.red);
    metric(s, "1", "套可续写内容平台\nTHE LOGBOOK", 924, 240, 240, false, C.red);
    addLine(s, 64, 430, 1120, C.line, 1);
    const items = ["门店 POP / 电商详情", "品牌 Social / 代言人 Social", "PR / 媒体 / 会员内容", "赛事 / 首映 / 线下体验"];
    for (let i = 0; i < 4; i++) {
      const x = 64 + i * 280;
      addText(s, items[i], x, 478, 245, 54, { fontSize: 21, bold: true, color: C.navy, align: "center", valign: "middle" });
    }
    addText(s, "最终尺寸、条数与字幕语言在 Creative PPM 前锁定；所有原始素材按长期复用标准归档。", 64, 598, 1120, 28, { fontSize: 18, color: C.sea, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp. 16, 47-52.", "智能纪要：品牌150周年拍摄项目规划, pp. 2-3."]);
  }

  // 31 — Timeline and budget control
  {
    const s = newSlide(p);
    s.background.fill = C.navy;
    titleBlock(s, "EXECUTION CONTROL", "200万预算，要在屏幕上看得见", 31, true);
    const steps = [
      ["8/17", "FIRST PITCH"],
      ["8/24", "SECOND PITCH"],
      ["8/26", "AGENCY AWARD"],
      ["8/27-9/23", "DEVELOP / RECCE"],
      ["9/25 TBC", "CREATIVE PPM"],
      ["10/24-31", "ITALY SHOOT"],
    ];
    addLine(s, 92, 350, 1060, "#3B5360", 2);
    for (let i = 0; i < steps.length; i++) {
      const x = 76 + i * 190;
      addShape(s, x, 334, 32, 32, i >= 4 ? C.red : "#55717E", "none", 0, 16);
      addText(s, steps[i][0], x - 20, 262, 150, 28, { fontSize: 15, bold: true, color: i >= 4 ? C.red : C.white, typeface: FONT_SANS });
      addText(s, steps[i][1], x - 20, 388, 165, 54, { fontSize: 16, bold: true, color: "#B8C8CE", typeface: FONT_SANS, lineSpacing: 1.05 });
    }
    addText(s, "报价框架", 64, 512, 160, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "创作团队 / 意大利本地制作 / 船与场地 / 差旅与保险 / 后期与版本 / 天气机动", 64, 552, 1090, 34, { fontSize: 23, bold: true, color: C.white });
    addText(s, "原则：先保导演、摄影、真实行动与天气备份；每一笔费用都对应一个最终可见画面或交付风险。", 64, 612, 1088, 28, { fontSize: 17, color: "#AFC0C8" });
    note(s, ["User-provided HH Shooting Brief, pp. 43-45.", "智能纪要：品牌150周年拍摄项目规划, pp. 5-7.", "Budget categories are a quoting framework; detailed supplier pricing remains to be produced."]);
  }

  // 32 — Team and governance
  {
    const s = newSlide(p);
    s.background.fill = C.white;
    titleBlock(s, "TEAM & GOVERNANCE", "一套创意中枢，三条专业生产线", 32, false);
    addVLine(s, 640, 308, 58, C.red, 3);
    addLine(s, 252, 366, 776, C.line, 2);
    addVLine(s, 252, 366, 32, C.line, 2);
    addVLine(s, 646, 366, 32, C.line, 2);
    addVLine(s, 1040, 366, 32, C.line, 2);
    addShape(s, 498, 226, 284, 82, C.navy, "none", 0, 18);
    addText(s, "CREATIVE CORE", 520, 246, 240, 28, { fontSize: 18, bold: true, color: C.red, typeface: FONT_SANS, align: "center" });
    addText(s, "ECD / Strategy / Producer", 520, 278, 240, 22, { fontSize: 15, color: C.white, typeface: FONT_SANS, align: "center" });
    const teams = [
      ["FILM UNIT", "Director / DP / AD\nMarine + Trail action"],
      ["STILLS UNIT", "Photographer / Digitech\n8 Look output control"],
      ["FIELD SAFETY", "M32 skipper / Alpine guide\nMedical / weather lead"],
    ];
    for (let i = 0; i < 3; i++) {
      const x = 112 + i * 394;
      addShape(s, x, 398, 280, 144, C.paper, "none", 0, 18);
      addText(s, teams[i][0], x + 24, 424, 232, 30, { fontSize: 20, bold: true, color: C.red, typeface: FONT_SANS, align: "center" });
      addText(s, teams[i][1], x + 24, 470, 232, 58, { fontSize: 17, color: C.sea, align: "center", lineSpacing: 1.2 });
    }
    addText(s, "关键审批门：脚本锁定 → 样衣与艺人试装 → 风况与安全复核 → PPM → 每日素材复盘 → 精剪与 KV 同步审片", 64, 594, 1120, 40, { fontSize: 17, bold: true, color: C.navy, align: "center" });
    note(s, ["User-provided HH Shooting Brief, p. 43 (Team Structure requirement).", "Team roles are an original recommended governance structure."]);
  }

  // 33 — Why this wins
  {
    const s = newSlide(p);
    s.background.fill = C.paper;
    titleBlock(s, "WHY THIS WINS", "这套方案同时回答世界观、品牌、创意与执行", 33, false);
    const claims = [
      ["世界观", "用“进入而非征服”建立一套能长期指导品牌的上位理念"],
      ["品牌上", "把150年从历史数字变成可持续的信任资产"],
      ["创意上", "一本《航海志》统一两支片与8套Look，但不牺牲各自任务"],
      ["执行上", "真实环境、双机组、替模与多版本从第一天就被写进方案"],
    ];
    for (let i = 0; i < 4; i++) {
      const y = 216 + i * 92;
      addText(s, String(i + 1).padStart(2, "0"), 64, y, 60, 40, { fontSize: 22, bold: true, color: C.red, typeface: FONT_SANS });
      addText(s, claims[i][0], 154, y, 180, 40, { fontSize: 28, bold: true, color: C.navy });
      addText(s, claims[i][1], 390, y, 770, 50, { fontSize: 22, color: C.sea, lineSpacing: 1.12 });
      addLine(s, 154, y + 60, 1006, C.line, 1);
    }
    addText(s, "最终结果：市场先因王一博走近，再因 HH 的专业留下。", 64, 610, 1120, 38, { fontSize: 29, bold: true, color: C.red, align: "center" });
    note(s, ["Editorial synthesis based on all six user-provided source documents."]);
  }

  // 34 — Closing
  {
    const s = newSlide(p);
    addImage(s, A.afterglow, 0, 0, W, H, "Concept visual: runners looking toward the next ridge at afterglow");
    addShape(s, 0, 0, W, H, "#071925/30");
    addShape(s, 710, 0, 570, H, "#071925/80");
    addText(s, "THE NEXT PAGE", 764, 76, 390, 26, { fontSize: 15, bold: true, color: C.red, typeface: FONT_SANS });
    addText(s, "下一段 150 年，\n从已知地图之外开始。", 764, 152, 430, 150, { fontSize: 47, bold: true, color: C.white, lineSpacing: 1.08 });
    addLine(s, 764, 350, 300, C.red, 5);
    addText(s, "HELLY HANSEN\nTHE LOGBOOK《航海志》", 764, 388, 420, 84, { fontSize: 25, bold: true, color: C.paper, lineSpacing: 1.18 });
    addText(s, "#向更远 赴新域#", 764, 540, 360, 42, { fontSize: 29, bold: true, color: C.red });
    addText(s, "让我们写下第一页。", 764, 604, 360, 32, { fontSize: 21, color: C.white });
    note(s, ["OpenAI ImageGen concept visual; provenance recorded in source-notes.txt.", "Closing line is original proposal copy."]);
  }

  // Export previews, layouts and deck.
  for (const [index, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(`${BUILD}/rendered/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${BUILD}/rendered/${stem}.layout.json`, await layout.text());
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(`${BUILD}/rendered/deck-montage.webp`, new Uint8Array(await montage.arrayBuffer()));
  const file = await PresentationFile.exportPptx(p);
  await file.save(`${OUT}/HH150周年_航海志_完整世界观创意提案_V2.pptx`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
