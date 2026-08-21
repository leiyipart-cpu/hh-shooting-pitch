import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const W = 1280;
const H = 720;
const ROOT = "/Users/leiyi/Documents/ChatGPT/HH拍摄pitch";
const BUILD = `${ROOT}/tmp_hh_pitch/deck_build_v4`;
const OUT = `${ROOT}/output/HH_2027SS_自在天地间_8LOOK统一优化提案_V2.pptx`;

const C = {
  olive: "#58572F",
  oliveDark: "#34351F",
  moss: "#77734B",
  fog: "#E9E9E3",
  paper: "#F2F1EC",
  white: "#FFFFFF",
  black: "#171914",
  red: "#FF1515",
  grey: "#7D7F75",
  line: "#C6C6BC",
};

const FONT_CN = "PingFang SC";
const FONT_EN = "Impact";
const FONT_BODY = "Avenir Next";

const P = {
  ref1: "/Users/leiyi/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_quxdxdtw6hg921_b667/temp/RWTemp/2026-08/a03da6a1b3b13d5b61e64457b271073e/9ade166d5007291926bf1f9b6ef6a035.png",
  ref2: "/Users/leiyi/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_quxdxdtw6hg921_b667/temp/RWTemp/2026-08/a03da6a1b3b13d5b61e64457b271073e/ff6b4a5a81eeaf08498084c5ed523b3d.png",
  ref3: "/Users/leiyi/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_quxdxdtw6hg921_b667/temp/RWTemp/2026-08/a03da6a1b3b13d5b61e64457b271073e/7bb33db253eaab6661f0c5745b6b4927.png",
  sailing: `${ROOT}/assets/generated/sailing-crew-hero.png`,
  hands: `${ROOT}/assets/generated/sailing-hands.png`,
  night: `${ROOT}/assets/generated/sailing-nightglow.png`,
  trail: `${ROOT}/assets/generated/trail-hero.png`,
  shoe: `${ROOT}/assets/generated/trail-shoe-grip.png`,
  afterglow: `${ROOT}/assets/generated/trail-afterglow.png`,
  logbook: `${ROOT}/assets/generated/logbook-hero.png`,
  brief16: `${ROOT}/tmp_hh_pitch/renders/brief/page-16.jpg`,
  brief20: `${ROOT}/tmp_hh_pitch/renders/brief/page-20.jpg`,
  brief21: `${ROOT}/tmp_hh_pitch/renders/brief/page-21.jpg`,
  brief22: `${ROOT}/tmp_hh_pitch/renders/brief/page-22.jpg`,
  brief23: `${ROOT}/tmp_hh_pitch/renders/brief/page-23.jpg`,
  brief24: `${ROOT}/tmp_hh_pitch/renders/brief/page-24.jpg`,
  brief30: `${ROOT}/tmp_hh_pitch/renders/brief/page-30.jpg`,
  brief31: `${ROOT}/tmp_hh_pitch/renders/brief/page-31.jpg`,
  brief32: `${ROOT}/tmp_hh_pitch/renders/brief/page-32.jpg`,
  brief35: `${ROOT}/tmp_hh_pitch/renders/brief/page-35.jpg`,
};

function mime(path) {
  return path.toLowerCase().endsWith(".jpg") || path.toLowerCase().endsWith(".jpeg") ? "image/jpeg" : "image/png";
}

function rect(slide, x, y, w, h, fill, name) {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

function line(slide, x, y, w, fill = C.red, width = 4, name) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x, top: y, width: w, height: 0 },
    fill: "none",
    line: { style: "solid", fill, width },
  });
}

function vline(slide, x, y, h, fill = C.red, width = 4, name) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x, top: y, width: 0, height: h },
    fill: "none",
    line: { style: "solid", fill, width },
  });
}

function text(slide, value, x, y, w, h, o = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    name: o.name,
    position: { left: x, top: y, width: w, height: h },
    fill: o.fill ?? "none",
    line: { style: "solid", fill: o.lineFill ?? "none", width: o.lineWidth ?? 0 },
  });
  s.text = value;
  s.text.style = {
    fontSize: o.fontSize ?? 22,
    bold: o.bold ?? false,
    color: o.color ?? C.black,
    alignment: o.align ?? "left",
    verticalAlignment: o.valign ?? "top",
    typeface: o.typeface ?? FONT_CN,
    lineSpacing: o.lineSpacing ?? 1.05,
    autoFit: o.autoFit ?? "none",
    insets: o.insets ?? { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return s;
}

function image(slide, asset, x, y, w, h, alt, fit = "cover") {
  return slide.images.add({
    blob: asset.bytes,
    contentType: asset.contentType,
    alt,
    fit,
    position: { left: x, top: y, width: w, height: h },
  });
}

function note(slide, sources) {
  slide.speakerNotes.textFrame.setText(`[Sources]\n${sources.map((s) => `- ${s}`).join("\n")}`);
}

const SLIDE_NUMBERS = new WeakMap();

function chrome(slide, n, dark = false, section = "HELLY HANSEN 27SS") {
  const actualN = SLIDE_NUMBERS.get(slide) ?? n;
  const color = dark ? C.white : C.black;
  text(slide, "HH  ×  VCHINA", 1030, 25, 190, 30, { fontSize: 20, bold: true, color, typeface: FONT_EN, align: "right" });
  text(slide, section.toUpperCase(), 48, 676, 360, 18, { fontSize: 11, bold: true, color: dark ? "#FFFFFF/75" : C.grey, typeface: FONT_BODY });
  text(slide, String(actualN).padStart(2, "0"), 1168, 676, 52, 18, { fontSize: 11, bold: true, color: dark ? "#FFFFFF/75" : C.grey, typeface: FONT_BODY, align: "right" });
}

function title(slide, en, cn, n, dark = false, sub) {
  const color = dark ? C.white : C.oliveDark;
  text(slide, en.toUpperCase(), 48, 46, 920, 48, { fontSize: 42, bold: true, color, typeface: FONT_EN, lineSpacing: 0.9 });
  line(slide, 48, 108, 98, C.red, 5);
  text(slide, cn, 48, 128, 1040, 62, { fontSize: 40, bold: true, color, lineSpacing: 1.0 });
  if (sub) text(slide, sub, 48, 198, 1060, 52, { fontSize: 19, color: dark ? "#FFFFFF/82" : C.grey, typeface: FONT_BODY, lineSpacing: 1.15 });
  chrome(slide, n, dark);
}

function photoBg(slide, asset, overlay = "#11140E/36") {
  image(slide, asset, 0, 0, W, H, "Atmospheric outdoor campaign visual");
  if (overlay) rect(slide, 0, 0, W, H, overlay);
}

function newSlide(p) {
  const s = p.slides.add();
  SLIDE_NUMBERS.set(s, p.slides.items.length);
  return s;
}

async function main() {
  await fs.mkdir(`${BUILD}/rendered`, { recursive: true });
  const loaded = await Promise.all(Object.entries(P).map(async ([key, path]) => [key, { path, bytes: await fs.readFile(path), contentType: mime(path) }]));
  const A = Object.fromEntries(loaded);
  const p = Presentation.create({ slideSize: { width: W, height: H } });

  // 01–03 supplied visual direction, retained as exact opening pages.
  for (const [idx, key] of ["ref1", "ref2", "ref3"].entries()) {
    const s = newSlide(p);
    image(s, A[key], 0, 0, W, H, `User-provided visual reference page ${idx + 1}`);
    note(s, [`User-provided visual reference: ${P[key]}.`]);
  }

  // 04 Cultural collision
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#F4F3ED/82"); chrome(s, 4, false, "WORLDVIEW");
    text(s, "THE CULTURAL COLLISION", 48, 44, 760, 48, { fontSize: 42, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "一北一东，两种自然观在此相遇", 48, 126, 850, 54, { fontSize: 38, bold: true, color: C.oliveDark });
    line(s, 48, 204, 110, C.red, 5);
    text(s, "FRILUFTSLIV", 48, 260, 520, 76, { fontSize: 66, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "自由地生活在自然中", 52, 340, 460, 40, { fontSize: 25, bold: true, color: C.olive });
    text(s, "×", 610, 280, 60, 70, { fontSize: 58, bold: true, color: C.red, typeface: FONT_EN, align: "center" });
    text(s, "天人合一", 760, 262, 430, 72, { fontSize: 54, bold: true, color: C.oliveDark, align: "right" });
    text(s, "人在天地之中，与万物共行", 710, 342, 480, 40, { fontSize: 25, bold: true, color: C.olive, align: "right" });
    text(s, "北欧给出一种生活方式，东方给出一种相处之道。", 48, 520, 1140, 56, { fontSize: 30, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["Norwegian Ministry of Climate and Environment, Friluftsliv White Paper (2015–2016).", "张岱年：《中国哲学中“天人合一”思想的剖析》，北京大学学报，1985年。", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 05 Friluftsliv breakdown
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "FRI · LUFT · LIV", "Friluftsliv：自由地生活在自然之中", 5, false, "不是一次远征，而是一种深植挪威日常的自然生活哲学。");
    const xs = [48, 438, 828];
    const en = ["FRI", "LUFT", "LIV"];
    const cn = ["自由", "空气与户外", "生活，也是生命"];
    const body = ["暂时离开目标与压力，在自然中重新获得身心的自由。", "走进开放空气，亲身感受风、雨、温度与天气。", "自然不是偶尔抵达的目的地，而是持续发生的生活方式。"];
    for (let i = 0; i < 3; i++) {
      text(s, en[i], xs[i], 286, 320, 96, { fontSize: 82, bold: true, color: i === 1 ? C.red : C.oliveDark, typeface: FONT_EN });
      text(s, cn[i], xs[i], 386, 330, 42, { fontSize: 28, bold: true, color: C.oliveDark });
      line(s, xs[i], 446, 280, C.line, 2);
      text(s, body[i], xs[i], 476, 310, 94, { fontSize: 19, color: C.grey, lineSpacing: 1.2 });
    }
    note(s, ["Norwegian government cultural introduction to Friluftsliv: Fri / Luft / Liv.", "User-approved worldview copy developed in this project."]);
  }

  // 06 Tianren Heyi
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "HEAVEN · HUMAN · ONE", "天人合一：人在天地运行之中", 6, false, "不是消极服从自然，而是理解规律，在其中找到恰当的行动方式。");
    const items = [
      ["天", "自然自有秩序", "风有方向，水有流势，山有地形。"],
      ["人", "自然的参与者", "人不凌驾于自然，也不站在自然之外。"],
      ["合一", "理解、回应、共行", "观察变化，因势行动，与天地保持连接。"],
    ];
    items.forEach((it, i) => {
      const y = 280 + i * 116;
      text(s, it[0], 58, y, 170, 72, { fontSize: i === 2 ? 48 : 62, bold: true, color: i === 1 ? C.red : C.oliveDark, align: "center" });
      line(s, 266, y + 38, 90, C.red, 3);
      text(s, it[1], 392, y, 360, 42, { fontSize: 28, bold: true, color: C.oliveDark });
      text(s, it[2], 760, y + 3, 440, 54, { fontSize: 20, color: C.grey, lineSpacing: 1.15 });
    });
    note(s, ["张岱年：《中国哲学中“天人合一”思想的剖析》，北京大学学报，1985年。", "Stanford Encyclopedia of Philosophy: Daoism."]);
  }

  // 07 synthesis
  {
    const s = newSlide(p); photoBg(s, A.logbook, "#1A1C13/66"); chrome(s, 7, true, "WORLDVIEW");
    text(s, "NORDIC LIFE. EASTERN WISDOM.", 48, 46, 900, 46, { fontSize: 40, bold: true, color: C.white, typeface: FONT_EN });
    line(s, 48, 112, 100, C.red, 5);
    text(s, "北欧给出生活方式，\n东方给出相处之道。", 48, 152, 840, 130, { fontSize: 48, bold: true, color: C.white, lineSpacing: 1.05 });
    text(s, "HELLY HANSEN 给出将它付诸实践的专业能力。", 48, 330, 980, 56, { fontSize: 31, bold: true, color: C.white });
    text(s, "共生，是世界观。\n专业，是方法。\n自在与鲜活，是结果。", 744, 436, 460, 150, { fontSize: 26, bold: true, color: C.white, align: "right", lineSpacing: 1.25 });
    note(s, ["User-approved strategic synthesis developed in this project.", "OpenAI ImageGen concept visual; source file in project assets/generated/logbook-hero.png."]);
  }

  // 08 Brand role
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#F1EFE8/82"); chrome(s, 8, false, "BRAND ROLE");
    text(s, "PROFESSIONALISM MAKES FREEDOM POSSIBLE", 48, 44, 1020, 48, { fontSize: 40, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "专业，让自在发生", 48, 160, 780, 76, { fontSize: 58, bold: true, color: C.oliveDark });
    line(s, 48, 260, 120, C.red, 5);
    text(s, "真正的自在，不是自然为人让路，\n而是人有能力在自然中找到自己的路。", 48, 310, 840, 120, { fontSize: 32, bold: true, color: C.oliveDark, lineSpacing: 1.18 });
    text(s, "专业，不是把自然挡在身体之外；\n而是让人拥有置身其中的底气。", 650, 500, 540, 96, { fontSize: 25, bold: true, color: C.olive, align: "right", lineSpacing: 1.2 });
    note(s, ["User-approved worldview and brand-role copy.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 09 Architecture
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "ONE WORLDVIEW · TWO ACTIONS", "一个世界观，统领2027春夏全部内容", 9, false);
    const ys = [264, 346, 428, 510];
    const left = ["文化根源", "世界观大主题", "品牌方法", "年度行动"];
    const right = ["Friluftsliv × 天人合一", "自在天地间 · LIVE FREE IN THE ELEMENTS", "Trust by Professional · 专业，让自在发生", "向更远 · 赴新域"];
    ys.forEach((y, i) => {
      text(s, left[i], 48, y, 200, 42, { fontSize: 18, bold: true, color: C.grey });
      line(s, 260, y + 20, 130, i === 1 ? C.red : C.line, i === 1 ? 4 : 2);
      text(s, right[i], 420, y - 5, 760, 54, { fontSize: i === 1 ? 32 : 26, bold: true, color: i === 1 ? C.red : C.oliveDark, typeface: i === 0 ? FONT_EN : FONT_CN });
    });
    note(s, ["User-approved campaign architecture based on HH 27SS Shooting Brief."]);
  }

  // 10 Looks chapter divider
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#13160F/48"); chrome(s, 10, true, "8 LOOKS · ONE UNIVERSE");
    text(s, "8 LOOKS · 3 RELATIONSHIPS", 48, 50, 1050, 76, { fontSize: 68, bold: true, color: C.white, typeface: FONT_EN });
    line(s, 48, 152, 118, C.red, 6);
    text(s, "8套Look，不是8个产品答案", 48, 196, 920, 66, { fontSize: 46, bold: true, color: C.white });
    text(s, "它们是“自在天地间”世界观下，\n人与自然建立关系的8种方式。", 48, 286, 790, 108, { fontSize: 30, bold: true, color: C.white, lineSpacing: 1.18 });
    text(s, "海，是来处", 48, 510, 300, 46, { fontSize: 29, bold: true, color: C.white });
    text(s, "野，是日常", 426, 510, 300, 46, { fontSize: 29, bold: true, color: C.white });
    text(s, "地势，是节奏", 812, 510, 360, 46, { fontSize: 29, bold: true, color: C.white });
    line(s, 48, 578, 1120, "#FFFFFF/45", 2);
    text(s, "不是把自然变成背景，而是让每套Look成为人与自然建立关系的证据。", 48, 606, 1080, 38, { fontSize: 21, bold: true, color: C.white });
    note(s, ["User-approved ‘自在天地间’ worldview developed in this project.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 11 Three relationships overview
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "ONE UNIVERSE · THREE RELATIONSHIPS", "一套世界观，三种与自然相处的方式", 11, false);
    const cards = [
      { x: 48, en: "SEA AS ORIGIN", cn: "海，是来处", nums: "01 · 04 · 07", body: "从品牌诞生的航海根源，\n转译为当代风格与日常生活。", color: C.oliveDark },
      { x: 444, en: "NATURE AS LIFE", cn: "野，是日常", nums: "02 · 03 · 06", body: "天气、森林与溪流，\n不是远方，而是每日生活。", color: C.red },
      { x: 840, en: "TERRAIN AS RHYTHM", cn: "地势，是节奏", nums: "05 · 08", body: "坡度提出问题，\n脚步与装备持续回应。", color: C.oliveDark },
    ];
    cards.forEach((c, i) => {
      rect(s, c.x, 270, 348, 318, i === 1 ? "#FFFFFF" : "#E2E2DA");
      text(s, `0${i + 1}`, c.x + 24, 292, 54, 36, { fontSize: 19, bold: true, color: i === 1 ? C.red : C.grey, typeface: FONT_EN });
      text(s, c.en, c.x + 24, 344, 300, 46, { fontSize: 31, bold: true, color: c.color, typeface: FONT_EN });
      text(s, c.cn, c.x + 24, 404, 300, 44, { fontSize: 28, bold: true, color: C.oliveDark });
      line(s, c.x + 24, 466, 86, i === 1 ? C.red : C.olive, 4);
      text(s, c.nums, c.x + 24, 492, 300, 28, { fontSize: 17, bold: true, color: C.grey, typeface: FONT_BODY });
      text(s, c.body, c.x + 24, 534, 300, 54, { fontSize: 18, bold: true, color: C.olive, lineSpacing: 1.18 });
    });
    text(s, "从来处，到生活，再到行动节奏——8套Look形成一条完整的品牌关系链。", 48, 622, 1140, 36, { fontSize: 22, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["User-provided HH 27SS Ambassador Shooting Brief, p.18 (8-look list).", "User-approved ‘自在天地间’ worldview developed in this project."]);
  }

  // 12 Sea as origin
  {
    const s = newSlide(p); photoBg(s, A.sailing, "#07100D/61"); chrome(s, 12, true, "LOOKS 01 · 04 · 07");
    text(s, "SEA AS ORIGIN", 48, 42, 760, 58, { fontSize: 52, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "海，是来处", 48, 120, 580, 62, { fontSize: 44, bold: true, color: C.white });
    line(s, 48, 204, 108, C.red, 5);
    text(s, "HH以海为生。三套Look不是三个航海造型，\n而是150年航海基因的三次转译。", 48, 236, 770, 80, { fontSize: 24, bold: true, color: C.white, lineSpacing: 1.18 });
    const seaLooks = [
      ["01", "历史", "经典航海150周年艺术家联名", "以帆、北极星与Crew，读懂共同方向"],
      ["04", "当代", "航海潮流150周年限定", "让专业遗产进入当代风格语境"],
      ["07", "日常", "航海生活方式", "让航海基因离开甲板，进入日常"],
    ];
    seaLooks.forEach((it, i) => {
      const y = 370 + i * 82;
      text(s, it[0], 48, y, 56, 34, { fontSize: 22, bold: true, color: i === 0 ? C.red : C.white, typeface: FONT_EN });
      text(s, it[1], 122, y, 78, 34, { fontSize: 19, bold: true, color: C.white });
      text(s, it[2], 220, y, 360, 34, { fontSize: 20, bold: true, color: C.white });
      text(s, it[3], 620, y, 540, 34, { fontSize: 18, bold: true, color: "#FFFFFF/82" });
      line(s, 48, y + 48, 1112, "#FFFFFF/35", 1);
    });
    text(s, "从航海诞生，到当代风格，再到日常生活。", 48, 628, 1110, 34, { fontSize: 24, bold: true, color: C.white, align: "right" });
    note(s, ["User-provided HH Shooting Brief, pp.19–25 and look list.", "OpenAI ImageGen concept visual; source file in project assets/generated/sailing-crew-hero.png."]);
  }

  // 13 Nature as life
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#F0EFE8/76"); chrome(s, 13, false, "LOOKS 02 · 03 · 06");
    text(s, "NATURE AS LIFE", 48, 42, 760, 58, { fontSize: 52, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "野，是日常", 48, 120, 580, 62, { fontSize: 44, bold: true, color: C.oliveDark });
    line(s, 48, 204, 108, C.red, 5);
    text(s, "Friluftsliv不是偶尔远征，\n而是自由地生活在自然中。", 48, 236, 650, 80, { fontSize: 26, bold: true, color: C.oliveDark, lineSpacing: 1.18 });
    const natureLooks = [
      ["02", "天气", "徒步1", "在真实天气中保持判断与从容"],
      ["03", "日常", "徒步2", "从远征叙事回到日常自然生活"],
      ["06", "流动", "溯溪", "回应水势，在流动中寻找路径"],
    ];
    natureLooks.forEach((it, i) => {
      const x = 48 + i * 396;
      rect(s, x, 372, 348, 200, i === 1 ? "#FFFFFF/84" : "#E2E2D8/86");
      text(s, it[0], x + 22, 390, 52, 32, { fontSize: 21, bold: true, color: i === 1 ? C.red : C.oliveDark, typeface: FONT_EN });
      text(s, it[1], x + 84, 390, 90, 32, { fontSize: 20, bold: true, color: C.oliveDark });
      text(s, it[2], x + 22, 438, 290, 40, { fontSize: 27, bold: true, color: C.oliveDark });
      line(s, x + 22, 492, 70, i === 1 ? C.red : C.olive, 4);
      text(s, it[3], x + 22, 516, 300, 44, { fontSize: 18, bold: true, color: C.olive, lineSpacing: 1.15 });
    });
    text(s, "自然不是偶尔抵达的远方，而是可以每天发生的生活。", 48, 618, 1138, 38, { fontSize: 24, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["User-provided HH Shooting Brief, pp.30–32 (Hiking and Stream looks).", "Norwegian Friluftsliv cultural concept used in the approved campaign worldview.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 14 Terrain as rhythm
  {
    const s = newSlide(p); photoBg(s, A.trail, "#0C110C/55"); chrome(s, 14, true, "LOOKS 05 · 08");
    text(s, "TERRAIN AS RHYTHM", 48, 42, 880, 58, { fontSize: 52, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "地势，是节奏", 48, 120, 620, 62, { fontSize: 44, bold: true, color: C.white });
    line(s, 48, 204, 108, C.red, 5);
    text(s, "山野不是等待被征服的赛道。\n坡度、碎石、水面与身体，共同决定每一步。", 48, 236, 820, 86, { fontSize: 25, bold: true, color: C.white, lineSpacing: 1.2 });
    rect(s, 48, 386, 518, 184, "#11150F/68");
    text(s, "05", 74, 408, 62, 34, { fontSize: 24, bold: true, color: C.red, typeface: FONT_EN });
    text(s, "感知", 148, 408, 100, 34, { fontSize: 20, bold: true, color: C.white });
    text(s, "越野跑", 74, 458, 220, 42, { fontSize: 30, bold: true, color: C.white });
    text(s, "让脚步回应地势与身体节奏", 74, 520, 420, 34, { fontSize: 20, bold: true, color: "#FFFFFF/82" });
    rect(s, 614, 386, 518, 184, "#11150F/68");
    text(s, "08", 640, 408, 62, 34, { fontSize: 24, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "能力", 714, 408, 100, 34, { fontSize: 20, bold: true, color: C.white });
    text(s, "越野跑场景", 640, 458, 260, 42, { fontSize: 30, bold: true, color: C.white });
    text(s, "以专业装备进入新的运动领域", 640, 520, 420, 34, { fontSize: 20, bold: true, color: "#FFFFFF/82" });
    text(s, "与野同步 · 赴新域", 48, 618, 1080, 38, { fontSize: 26, bold: true, color: C.white, align: "right" });
    note(s, ["User-provided HH Shooting Brief, pp.27–29 (Dearacer product communication).", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-hero.png."]);
  }

  // 11 visual tone
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#10130D/58"); chrome(s, 11, true, "VISUAL LANGUAGE");
    text(s, "REAL ELEMENTS. REAL MOVEMENT.", 48, 44, 900, 48, { fontSize: 42, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "不制造景观，让真实自然成为叙事力量", 48, 130, 850, 58, { fontSize: 38, bold: true, color: C.white });
    const words = ["真实天气", "真实动作", "极简构图", "北欧疏离", "专业高端"];
    words.forEach((w, i) => text(s, w, 50 + i * 236, 532, 206, 48, { fontSize: 24, bold: true, color: i === 0 ? C.red : C.white, align: "center" }));
    text(s, "自然不做背景；风、浪、地势与水流都是行动的一部分。", 48, 612, 900, 42, { fontSize: 22, color: C.white, bold: true });
    note(s, ["User-provided HH 27SS Ambassador Shooting Brief, pp.35–40 (Visual Guideline).", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 12 sailing divider
  {
    const s = newSlide(p); photoBg(s, A.sailing, "#09100D/40"); chrome(s, 12, true, "FILM 01 · SAILING HERITAGE");
    text(s, "SAIL AS ONE", 48, 54, 800, 84, { fontSize: 80, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "与海同舟", 48, 170, 520, 82, { fontSize: 60, bold: true, color: C.white });
    line(s, 48, 282, 120, C.red, 6);
    text(s, "向远，从来不是一个人的抵达。", 48, 332, 850, 60, { fontSize: 34, bold: true, color: C.white });
    text(s, "FILM 01 · 向更远", 48, 604, 420, 28, { fontSize: 18, bold: true, color: C.white, typeface: FONT_BODY });
    note(s, ["User-provided HH Shooting Brief, pp.19–25.", "OpenAI ImageGen concept visual; source file in project assets/generated/sailing-crew-hero.png."]);
  }

  // 13 sailing idea
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "THE CREW IS THE HERO", "主角不是某一个人，而是整个Crew", 13, false);
    text(s, "海上没有个人英雄。", 48, 276, 580, 70, { fontSize: 48, bold: true, color: C.oliveDark });
    text(s, "风决定条件，海浪改变航程，北极星确认方向；\n每一名Crew，都在自己的位置上完成判断、配合与承担。", 48, 366, 670, 118, { fontSize: 24, color: C.olive, lineSpacing: 1.28 });
    line(s, 792, 276, 340, C.red, 5);
    text(s, "王一博不是带领所有人的船长。\n他进入真实Crew，\n成为航行中不可或缺的一员。", 792, 314, 410, 170, { fontSize: 30, bold: true, color: C.oliveDark, lineSpacing: 1.22 });
    note(s, ["User-provided HH Shooting Brief, pp.19–20 (Crew spirit and ambassador role).", "User-approved Sailing concept: 与海同舟."]);
  }

  // 14 two meanings of same boat
  {
    const s = newSlide(p); photoBg(s, A.sailing, "#F0EFE9/78"); chrome(s, 14, false, "FILM 01 · IDEA");
    text(s, "WITH EACH OTHER. WITH THE SEA.", 48, 44, 900, 48, { fontSize: 42, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "“与海同舟”的两层关系", 48, 136, 680, 54, { fontSize: 38, bold: true, color: C.oliveDark });
    text(s, "人与人同舟", 48, 284, 430, 52, { fontSize: 34, bold: true, color: C.red });
    text(s, "Crew彼此信任、共同判断、共同承担。", 48, 350, 500, 58, { fontSize: 23, bold: true, color: C.oliveDark });
    text(s, "人与自然同舟", 680, 284, 430, 52, { fontSize: 34, bold: true, color: C.red });
    text(s, "不对抗海洋，而是理解风浪、借势成航。", 680, 350, 500, 58, { fontSize: 23, bold: true, color: C.oliveDark });
    text(s, "自在不是一个人在海上的自由，\n而是一群人在真实风浪中共同前进的从容。", 280, 510, 720, 88, { fontSize: 29, bold: true, color: C.oliveDark, align: "center", lineSpacing: 1.2 });
    note(s, ["User-approved Sailing subtheme explanation.", "OpenAI ImageGen concept visual; source file in project assets/generated/sailing-crew-hero.png."]);
  }

  // 15 artist motifs
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "ART BECOMES ACTION", "艺术家图案，不是装饰，而是一套航海图腾", 15, false);
    image(s, A.brief20, 48, 252, 540, 316, "Sailing TVC key message and artist collaboration visual", "contain");
    const rows = [
      ["帆", "承载HH 150年的航海故事"],
      ["北极星", "所有Crew共同确认的方向"],
      ["夜光海浪", "保存白昼的光，在未知中继续指引"],
    ];
    rows.forEach((r, i) => {
      const y = 270 + i * 102;
      text(s, r[0], 654, y, 180, 44, { fontSize: 28, bold: true, color: i === 1 ? C.red : C.oliveDark });
      text(s, r[1], 850, y + 2, 350, 58, { fontSize: 20, bold: true, color: C.olive, lineSpacing: 1.15 });
      line(s, 654, y + 72, 546, C.line, 1);
    });
    note(s, ["User-provided HH Shooting Brief, p.20 (sail and North Star motif meanings).", "User-provided HH Shooting Brief, p.22 (Nordic artist collaboration and glow-wave treatment)."]);
  }

  // 16 collective voyage
  {
    const s = newSlide(p); photoBg(s, A.hands, "#11130E/54"); chrome(s, 16, true, "FILM 01 · CORE DRAMA");
    text(s, "FURTHER, MADE TOGETHER", 48, 46, 900, 48, { fontSize: 44, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "远方，由所有人共同完成", 48, 142, 900, 62, { fontSize: 44, bold: true, color: C.white });
    line(s, 48, 232, 120, C.red, 6);
    text(s, "艺术家完成图案的第一笔，\n风与海留下下一笔，\nCrew用行动完成整件作品。", 48, 288, 620, 166, { fontSize: 31, bold: true, color: C.white, lineSpacing: 1.2 });
    text(s, "王一博，也是其中的一笔。", 748, 516, 440, 54, { fontSize: 28, bold: true, color: C.white, align: "right" });
    note(s, ["User-approved Sailing concept synthesis.", "OpenAI ImageGen concept visual; source file in project assets/generated/sailing-hands.png."]);
  }

  // 17 sailing story
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "FOUR MOVEMENTS", "一次航行，四次关系推进", 17, false);
    const steps = [
      ["01", "入船", "王一博进入Crew，学习位置与规则。"],
      ["02", "遇风", "天气改变，个人动作必须连接他人。"],
      ["03", "同向", "帆、绳索、判断与信任形成一个系统。"],
      ["04", "向远", "全体Crew完成关键调帆，共同驶向远方。"],
    ];
    steps.forEach((st, i) => {
      const x = 48 + i * 300;
      text(s, st[0], x, 276, 90, 50, { fontSize: 34, bold: true, color: i === 2 ? C.red : C.oliveDark, typeface: FONT_EN });
      text(s, st[1], x, 342, 240, 50, { fontSize: 31, bold: true, color: C.oliveDark });
      line(s, x, 410, 238, i === 2 ? C.red : C.line, i === 2 ? 4 : 2);
      text(s, st[2], x, 444, 240, 100, { fontSize: 19, color: C.grey, lineSpacing: 1.2 });
    });
    text(s, "所有关键动作，都必须依赖上一个人的动作，并触发下一个人的回应。", 48, 602, 1140, 36, { fontSize: 23, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["Creative narrative developed from HH Shooting Brief, pp.19–20 and user-approved Crew direction."]);
  }

  // 18 camera principles
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "SHOOT THE RELATIONSHIP", "镜头不追逐英雄，镜头捕捉关系", 18, false);
    const items = [
      ["不以单人Hero Shot开场", "先让观众看见整艘船如何运转。"],
      ["动作必须彼此连接", "手、眼神、绳索、帆面形成连续反应。"],
      ["高潮属于整个Crew", "关键调帆或转向由所有人共同完成。"],
      ["自然始终在场", "风、浪、光线与天气推动情节，而非装饰画面。"],
    ];
    items.forEach((it, i) => {
      const y = 270 + i * 84;
      text(s, `0${i + 1}`, 48, y, 60, 34, { fontSize: 20, bold: true, color: i === 0 ? C.red : C.oliveDark, typeface: FONT_EN });
      text(s, it[0], 128, y - 2, 390, 40, { fontSize: 25, bold: true, color: C.oliveDark });
      text(s, it[1], 580, y, 610, 48, { fontSize: 19, color: C.grey, lineSpacing: 1.15 });
      line(s, 128, y + 54, 1062, C.line, 1);
    });
    note(s, ["User-approved director/camera principles for de-emphasizing individual heroism."]);
  }

  // 19 sailing looks
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "THREE SAILING EXPRESSIONS", "三套航海Look，完成从历史到当代生活", 19, false);
    const imgs = [A.brief22, A.brief23, A.brief24];
    const labels = ["01 · SAILING HERITAGE", "04 · H2ALIVE", "07 · HH STØTTE"];
    const desc = ["艺术家联名 × Crew精神", "150周年限定 × 航海潮流", "功能设计 × 航海生活方式"];
    imgs.forEach((im, i) => {
      const x = 48 + i * 396;
      image(s, im, x, 252, 360, 238, `Sailing look ${i + 1} reference`, "cover");
      text(s, labels[i], x, 512, 360, 32, { fontSize: 18, bold: true, color: C.oliveDark, typeface: FONT_BODY });
      text(s, desc[i], x, 556, 360, 54, { fontSize: 20, bold: true, color: i === 0 ? C.red : C.olive });
    });
    note(s, ["User-provided HH Shooting Brief, pp.22–24 (Sailing Heritage, H2Alive, HH Støtte)."]);
  }

  // 20 sailing ending
  {
    const s = newSlide(p); photoBg(s, A.night, "#06100D/40"); chrome(s, 20, true, "FILM 01 · END FRAME");
    text(s, "WITH THE SEA, AS ONE CREW", 48, 52, 920, 50, { fontSize: 44, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "与海同舟", 48, 196, 560, 86, { fontSize: 68, bold: true, color: C.white });
    text(s, "向更远", 48, 296, 560, 70, { fontSize: 52, bold: true, color: C.red });
    text(s, "自在天地间", 48, 558, 420, 52, { fontSize: 31, bold: true, color: C.white });
    text(s, "HELLY HANSEN 150", 48, 618, 420, 28, { fontSize: 18, bold: true, color: C.white, typeface: FONT_BODY });
    note(s, ["User-approved Sailing endline.", "OpenAI ImageGen concept visual; source file in project assets/generated/sailing-nightglow.png."]);
  }

  // 21 trail divider
  {
    const s = newSlide(p); photoBg(s, A.trail, "#0B0F0A/40"); chrome(s, 21, true, "FILM 02 · DEARACER");
    text(s, "IN STEP WITH THE WILD", 48, 54, 1040, 84, { fontSize: 76, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "与野同步", 48, 176, 560, 82, { fontSize: 60, bold: true, color: C.white });
    line(s, 48, 286, 120, C.red, 6);
    text(s, "不是让山野适应速度，\n而是找到与山野同步的那一步。", 48, 336, 770, 110, { fontSize: 34, bold: true, color: C.white, lineSpacing: 1.15 });
    text(s, "FILM 02 · 赴新域", 48, 606, 420, 28, { fontSize: 18, bold: true, color: C.white, typeface: FONT_BODY });
    note(s, ["User-provided HH Shooting Brief, pp.26–33.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-hero.png."]);
  }

  // 22 Trail premise
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "THE WILD SETS THE QUESTION", "山野不断提问，脚步持续回应", 22, false);
    text(s, "坡度、岩石、泥土、溪流、风与身体状态，\n每一刻都在重新定义下一步。", 48, 286, 650, 106, { fontSize: 32, bold: true, color: C.oliveDark, lineSpacing: 1.22 });
    line(s, 48, 426, 110, C.red, 5);
    text(s, "真正的越野，不是把自然变成标准赛道；\n而是在不断变化的地形中，保持感知与行动。", 48, 472, 720, 100, { fontSize: 24, color: C.olive, lineSpacing: 1.25 });
    text(s, "王一博不是唯一的Racer\n他代表每一个正在寻找\n自己节奏的人。", 790, 310, 400, 166, { fontSize: 28, bold: true, color: C.oliveDark, align: "right", lineSpacing: 1.22 });
    note(s, ["User-approved Dearacer worldview: 与野同步.", "User-provided HH Shooting Brief, pp.26–29."]);
  }

  // 23 sync
  {
    const s = newSlide(p); photoBg(s, A.trail, "#F2F0E8/77"); chrome(s, 23, false, "FILM 02 · IDEA");
    text(s, "SYNC IS NOT SLOWING DOWN", 48, 46, 900, 48, { fontSize: 44, bold: true, color: C.oliveDark, typeface: FONT_EN });
    text(s, "同步，不是降低速度", 48, 146, 730, 64, { fontSize: 46, bold: true, color: C.oliveDark });
    text(s, "它意味着跑者不断感知环境，\n在每一次落脚之间作出新的判断。", 48, 254, 650, 94, { fontSize: 29, bold: true, color: C.oliveDark, lineSpacing: 1.2 });
    const cues = ["看见地势", "听见呼吸", "调整重心", "回应地面"];
    cues.forEach((c, i) => text(s, c, 48 + i * 290, 500, 250, 54, { fontSize: 27, bold: true, color: i === 3 ? C.red : C.oliveDark, align: "center" }));
    text(s, "自在，发生在身体与山野找到共同节奏的那一刻。", 180, 604, 920, 38, { fontSize: 26, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["User-approved Dearacer concept explanation.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-hero.png."]);
  }

  // 24 Dear Racer reinterpretation
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "DEAR RACER IS AN OPEN CALL", "Dear Racer，不必被拍成一封具体的信", 24, false);
    text(s, "它不是一种格式，\n而是HH对所有跑者的一次公开称呼。", 48, 286, 620, 122, { fontSize: 36, bold: true, color: C.oliveDark, lineSpacing: 1.2 });
    line(s, 48, 446, 110, C.red, 5);
    text(s, "不限定手写、信封或逐字旁白。\n可以是一段由呼吸、脚步、碎石、水流与风共同组成的开放宣言。", 48, 492, 710, 100, { fontSize: 22, color: C.olive, lineSpacing: 1.25 });
    text(s, "DEAR RACER,\nFIND YOUR STEP\nIN THE WILD.", 820, 294, 370, 180, { fontSize: 46, bold: true, color: C.red, typeface: FONT_EN, align: "right", lineSpacing: 0.95 });
    note(s, ["User-provided HH Shooting Brief, p.26: ‘一封写给每一位Racer的信’仅作为创意启发，不限定表现形式。", "User-approved reinterpretation of Dear Racer as an open call."]);
  }

  // 25 Trail story
  {
    const s = newSlide(p); s.background.fill = C.paper; title(s, "FOUR TERRAIN RESPONSES", "不是穿过山野，而是与山野形成同一个运动系统", 25, false);
    const steps = [
      ["01", "自然先动", "风过草木，水绕石块，地形先建立节奏。"],
      ["02", "身体进入", "呼吸、步幅与重心开始回应环境。"],
      ["03", "装备连接", "轻、弹、稳、抓把变化转化为下一步。"],
      ["04", "同步发生", "跑者不再穿越景观，而与山野共同运动。"],
    ];
    steps.forEach((st, i) => {
      const x = 48 + i * 300;
      text(s, st[0], x, 278, 90, 50, { fontSize: 34, bold: true, color: i === 3 ? C.red : C.oliveDark, typeface: FONT_EN });
      text(s, st[1], x, 342, 240, 50, { fontSize: 29, bold: true, color: C.oliveDark });
      line(s, x, 410, 238, i === 3 ? C.red : C.line, i === 3 ? 4 : 2);
      text(s, st[2], x, 444, 240, 100, { fontSize: 19, color: C.grey, lineSpacing: 1.2 });
    });
    text(s, "每一次地形变化，都是一次产品功能被真实验证的机会。", 48, 604, 1140, 36, { fontSize: 23, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["Creative narrative developed from HH Shooting Brief and user-approved 与野同步 direction."]);
  }

  // 26 product behavior
  {
    const s = newSlide(p); photoBg(s, A.shoe, "#0E110C/50"); chrome(s, 26, true, "FILM 02 · PRODUCT");
    text(s, "THE SHOE BECOMES THE CONNECTION", 48, 44, 980, 48, { fontSize: 42, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "Dearacer，让脚步与地面建立真实连接", 48, 126, 920, 58, { fontSize: 38, bold: true, color: C.white });
    const features = [
      ["轻", "保持节奏"],
      ["弹", "转化力量"],
      ["稳", "应对变化"],
      ["抓", "连接地面"],
    ];
    features.forEach((f, i) => {
      const x = 52 + i * 292;
      text(s, f[0], x, 478, 100, 70, { fontSize: 56, bold: true, color: i === 3 ? C.red : C.white, align: "center" });
      text(s, f[1], x + 100, 500, 160, 40, { fontSize: 21, bold: true, color: C.white });
    });
    note(s, ["User-provided HH Shooting Brief, pp.27–29 (Dearacer product communication).", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-shoe-grip.png."]);
  }

  // 27 trail camera
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#10120D/45"); chrome(s, 27, true, "FILM 02 · CAMERA");
    text(s, "SHOOT THE RESPONSE", 48, 44, 840, 48, { fontSize: 44, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "镜头拍的不是速度，而是每一步如何发生", 48, 132, 930, 62, { fontSize: 40, bold: true, color: C.white });
    const list = ["地形先出现，人物后进入", "脚步、呼吸、重心形成节奏", "鞋底与岩石、泥土、水面真实接触", "王一博的专业状态来自应对，而非摆拍"];
    list.forEach((it, i) => {
      text(s, `0${i + 1}`, 50, 276 + i * 72, 54, 30, { fontSize: 18, bold: true, color: i === 2 ? C.red : C.white, typeface: FONT_EN });
      text(s, it, 124, 272 + i * 72, 800, 40, { fontSize: 24, bold: true, color: C.white });
    });
    note(s, ["User-approved Trail Running camera direction.", "OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png."]);
  }

  // 29 unified system
  {
    const s = newSlide(p); s.background.fill = C.fog; title(s, "ONE UNIVERSE · TWO CHAPTERS", "两支TVC，是“自在天地间”的两种行动方式", 29, false);
    text(s, "与海同舟", 48, 272, 500, 64, { fontSize: 48, bold: true, color: C.oliveDark });
    text(s, "SAIL AS ONE", 48, 342, 420, 42, { fontSize: 34, bold: true, color: C.red, typeface: FONT_EN });
    text(s, "共同判断 · 彼此信任 · 向更远", 48, 418, 500, 48, { fontSize: 23, bold: true, color: C.olive });
    text(s, "自在源于共同承担", 48, 502, 500, 54, { fontSize: 29, bold: true, color: C.oliveDark });
    vline(s, 620, 272, 284, C.red, 3);
    text(s, "与野同步", 706, 272, 500, 64, { fontSize: 48, bold: true, color: C.oliveDark, align: "right" });
    text(s, "IN STEP WITH THE WILD", 706, 342, 500, 42, { fontSize: 34, bold: true, color: C.red, typeface: FONT_EN, align: "right" });
    text(s, "感知地形 · 回应变化 · 赴新域", 706, 418, 500, 48, { fontSize: 23, bold: true, color: C.olive, align: "right" });
    text(s, "自在源于彼此同步", 706, 502, 500, 54, { fontSize: 29, bold: true, color: C.oliveDark, align: "right" });
    text(s, "无论向远，还是赴新域，真正的自由都发生在人与自然建立关系之后。", 110, 612, 1060, 36, { fontSize: 25, bold: true, color: C.oliveDark, align: "center" });
    note(s, ["User-approved final campaign architecture."]);
  }

  // 30 close
  {
    const s = newSlide(p); photoBg(s, A.afterglow, "#15170F/52"); chrome(s, 30, true, "HELLY HANSEN 2027 SS");
    text(s, "LIVE FREE IN THE ELEMENTS", 48, 58, 1080, 70, { fontSize: 64, bold: true, color: C.white, typeface: FONT_EN });
    text(s, "自在天地间", 48, 196, 620, 92, { fontSize: 70, bold: true, color: C.white });
    line(s, 48, 326, 132, C.red, 7);
    text(s, "向更远 · 赴新域", 48, 380, 720, 70, { fontSize: 48, bold: true, color: C.white });
    text(s, "专业，让自在发生。", 48, 548, 520, 48, { fontSize: 28, bold: true, color: C.white });
    text(s, "STAY & FEEL ALIVE", 48, 608, 560, 38, { fontSize: 30, bold: true, color: C.white, typeface: FONT_EN });
    note(s, ["OpenAI ImageGen concept visual; source file in project assets/generated/trail-afterglow.png.", "User-approved final theme and annual communication line."]);
  }

  const inspect = await p.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 200000 });
  await fs.writeFile(`${BUILD}/deck.inspect.ndjson`, inspect.ndjson);
  for (const [i, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(i + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(`${BUILD}/rendered/${stem}.png`, new Uint8Array(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${BUILD}/rendered/${stem}.layout.json`, await layout.text());
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(`${BUILD}/rendered/montage.webp`, new Uint8Array(await montage.arrayBuffer()));
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(OUT);
  console.log(`Wrote ${OUT}`);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
