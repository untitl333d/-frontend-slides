// Builds "Tobacco Marketing on International Markets" in the same visual style as the
// MiCA seminar deck (presentations/mica-seminar/build.js on the eloquent-wozniak branch).
// Usage: node build.js out.pptx   (PREVIEW=1 draws grey boxes where photos go, for QA only)
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const JSZip = require("jszip");
const fs = require("fs");
const Fi = require("react-icons/fi");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const PREVIEW = process.env.PREVIEW === "1";

const HEX = {
  navy: "0E2841", red: "C00000", grey: "7F7F7F", light: "D1D1D1",
  blue: "0080CC", ink: "1A1A1A", tint: "F3F4F6", white: "FFFFFF",
};

const THEME = {
  name: "Corvinus Seminar",
  headFontFace: "Aptos",
  bodyFontFace: "Aptos",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.tint,
    accent1: HEX.navy, accent2: HEX.red, accent3: HEX.grey, accent4: HEX.light,
    accent5: HEX.blue, accent6: "404040", hlink: HEX.blue, folHlink: HEX.grey,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Tobacco Marketing on International Markets";
pres.author = "Amir Kenenbaev";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const NAVY = C.accent1, RED = C.accent2, GREY = C.accent3, LIGHT = C.accent4, BLUE = C.accent5, INK = C.text1, TINT = C.background2, WHITE = C.background1;
const BODY = C.accent6;

const TOPIC = "Tobacco Marketing on International Markets";
const L = 1.0;
const R = 12.58;
const W = R - L;

// ---------- photo slots (shared by layouts and preview boxes) ----------
const QUIZ_W = 2.55, QUIZ_GAP = (W - 4 * QUIZ_W) / 3;
const QUIZ = ["A", "B", "C", "D"].map((k, i) => ({ name: "pack" + k, label: "Insert photo: pack " + k, x: L + i * (QUIZ_W + QUIZ_GAP), y: 1.85, w: QUIZ_W, h: 3.9 }));
const REV_W = (W - 0.9) / 4;
const REVEAL = ["USA", "Hungary", "Kyrgyzstan", "Australia"].map((k, i) => ({ name: "photo" + k, label: "Insert photo: " + k + " pack", x: L + i * (REV_W + 0.3), y: 1.45, w: REV_W, h: 2.25 }));
const SIDE = { name: "picture", label: "Insert picture", x: 9.35, y: 1.55, w: 3.23, h: 5.1 };

// ---------- layouts ----------
const headerObjects = [
  { rect: { x: 0.55, y: 0, w: 0.3, h: 1.08, fill: { color: BLUE }, line: { color: BLUE, width: 0 } } },
  { placeholder: { options: { name: "topic", type: "body", x: L, y: 0.2, w: 9.6, h: 0.36, fontSize: 14, bold: true, color: GREY, valign: "middle", align: "left", margin: 0 }, text: "" } },
  { placeholder: { options: { name: "title", type: "title", x: L, y: 0.52, w: 11.2, h: 0.6, fontSize: 26, bold: true, color: NAVY, valign: "middle", align: "left", margin: 0 }, text: "" } },
  { text: { text: TOPIC, options: { x: L, y: 6.98, w: 8, h: 0.3, fontSize: 10, color: GREY, margin: 0 } } },
];
const picPh = (p) => ({ placeholder: { options: { name: p.name, type: "pic", x: p.x, y: p.y, w: p.w, h: p.h, color: GREY, fontSize: 12 }, text: p.label } });
const slideNumber = { x: 11.98, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: GREY, align: "right" };
pres.defineSlideMaster({ title: "CONTENT", background: { color: WHITE }, objects: headerObjects, slideNumber });
pres.defineSlideMaster({ title: "CONTENT_PICTURE", background: { color: WHITE }, objects: [...headerObjects, picPh(SIDE)], slideNumber });
pres.defineSlideMaster({ title: "PACK_QUIZ", background: { color: WHITE }, objects: [...headerObjects, ...QUIZ.map(picPh)], slideNumber });
pres.defineSlideMaster({ title: "PACK_REVEAL", background: { color: WHITE }, objects: [...headerObjects, ...REVEAL.map(picPh)], slideNumber });
pres.defineSlideMaster({ title: "TITLE", background: { color: NAVY }, objects: [] });

// ---------- helpers ----------
async function icon(Comp, color = "#FFFFFF") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: 256, strokeWidth: 1.75 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function addContent(topic, title, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master });
  s.addText(topic, { placeholder: "topic" });
  s.addText(title, { placeholder: "title" });
  return s;
}

function card(s, x, y, w, h, opts = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opts.fill || WHITE }, line: { color: opts.line || LIGHT, width: opts.lineW || 1 }, objectName: opts.name });
}

function txt(s, text, o) {
  s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontSize: 16, color: INK, ...o });
}

function source(s, text, x = L, y = 6.62, w = W) {
  txt(s, text, { x, y, w, h: 0.28, fontSize: 10, color: GREY, italic: true });
}

// QA-only stand-ins for empty photo placeholders (LibreOffice does not draw empty placeholders)
function previewSlots(s, slots) {
  if (!PREVIEW) return;
  slots.forEach((p) => s.addText(p.label, { isTextBox: true, shape: pres.shapes.RECTANGLE, x: p.x, y: p.y, w: p.w, h: p.h, fill: { color: "E7E9EE" }, line: { color: LIGHT, width: 1, dashType: "dash" }, fontSize: 12, color: GREY, align: "center", valign: "middle" }));
}

const chartBase = () => ({
  catAxisLabelColor: "404040", valAxisLabelColor: "7F7F7F",
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
  dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: "1A1A1A", dataLabelFontBold: true,
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: "0E2841",
  valGridLine: { color: "E5E5E5", size: 0.75 }, catGridLine: { style: "none" },
  catAxisLineShow: true, catAxisLineColor: "D1D1D1", valAxisLineShow: false,
  showLegend: false, showValue: true,
});

const bullets = (s, items, o) => s.addText(
  items.flatMap(([lead, rest], i) => [
    { text: lead + " ", options: { bold: true, color: NAVY, bullet: true } },
    { text: rest, options: { color: BODY, breakLine: i < items.length - 1 } },
  ]),
  { isTextBox: true, margin: 0, valign: "top", fontSize: 16, paraSpaceAfter: 10, ...o },
);
(async () => {
  const I = {
    user: await icon(Fi.FiUsers), target: await icon(Fi.FiTarget), wind: await icon(Fi.FiWind),
    box: await icon(Fi.FiPackage), tag: await icon(Fi.FiTag), store: await icon(Fi.FiShoppingBag), mega: await icon(Fi.FiSpeaker),
    candy: await icon(Fi.FiDroplet), usb: await icon(Fi.FiSmartphone), phone: await icon(Fi.FiInstagram), coin: await icon(Fi.FiDollarSign),
  };
  const iconCircle = (s, img, x, y, d = 0.62, fill = NAVY) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };
  const banner = (s, runs, y, h = 0.85, fontSize = 20) => s.addText(runs, { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y, w: W, h, fill: { color: NAVY }, color: WHITE, fontSize, bold: true, valign: "middle", margin: 16 });

  // ===== 1. Title =====
  {
    const s = pres.addSlide({ masterName: "TITLE" });
    txt(s, "International Marketing  ·  Individual presentation", { x: 0.8, y: 1.15, w: 8.4, h: 0.4, fontSize: 14, color: LIGHT, bold: true, charSpacing: 1 });
    txt(s, "Tobacco Marketing on International Markets", { x: 0.8, y: 1.7, w: 8.2, h: 2.3, fontSize: 40, bold: true, color: WHITE });
    txt(s, "How to sell a product nobody lets you advertise", { x: 0.8, y: 4.15, w: 8.0, h: 0.9, fontSize: 22, color: LIGHT });
    txt(s, "Corvinus University of Budapest  ·  Fall 2026", { x: 0.8, y: 6.35, w: 7.4, h: 0.4, fontSize: 14, color: LIGHT });
    s.addShape(pres.shapes.OVAL, { x: 9.95, y: 2.0, w: 2.0, h: 2.0, fill: { color: "1B3A5C" }, line: { color: WHITE, width: 1.5 }, objectName: "Photo Amir Kenenbaev" });
    txt(s, "Amir Kenenbaev", { x: 9.65, y: 4.2, w: 2.6, h: 0.4, fontSize: 15, color: WHITE, bold: true, align: "center" });
    s.addNotes("(~15 s) Good afternoon. I'm Amir, and my topic is tobacco marketing on international markets. But before I say anything about marketing, I want to show you some packs.\n\nPhoto circle: right-click the circle > Format Shape > Fill > Picture fill to insert a portrait.");
  }

  // ===== 2. Quiz =====
  {
    const s = addContent("Prologue", "Guess the country", "PACK_QUIZ");
    txt(s, "Four packs, four countries. What is different?", { x: L, y: 1.3, w: W, h: 0.4, fontSize: 18, color: BODY });
    previewSlots(s, QUIZ);
    QUIZ.forEach((p, i) => {
      s.addText("ABCD"[i], { isTextBox: true, shape: pres.shapes.OVAL, x: p.x, y: 5.95, w: 0.5, h: 0.5, fill: { color: RED }, color: WHITE, bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0 });
      txt(s, "Which country?", { x: p.x + 0.65, y: 6.03, w: p.w - 0.65, h: 0.36, fontSize: 15, color: GREY });
    });
    s.addNotes(
      "(~45 s) Here are four cigarette packs from four countries. Which country is each one from? And what is different about them? Take a few seconds and shout out your guesses.\n\n" +
      "PHOTOS TO INSERT (ideally the same brand, e.g. Marlboro, so only the rules differ):\n" +
      "A = USA: full brand design, small text-only warning on the side.\n" +
      "B = Hungary (EU): plain pack + 65% picture warning in Hungarian.\n" +
      "C = Kyrgyzstan: pack with the 50% picture warning in Kyrgyz/Russian (Eurasian Economic Union rules).\n" +
      "D = Australia: plain drab dark brown pack, 75% warning on the front.\n" +
      "Click the picture icon in each box to insert a photo; it fills and crops automatically.");
  }

  // ===== 3. Reveal =====
  {
    const s = addContent("Prologue", "Same cigarettes, four different rulebooks", "PACK_REVEAL");
    previewSlots(s, REVEAL);
    [
      ["A  ·  USA", "Text only", "Small warning on the side. The brand owns the front."],
      ["B  ·  Hungary", "65%", "Picture warnings and a plain pack since 2022."],
      ["C  ·  Kyrgyzstan", "50%", "Picture warnings under Eurasian Economic Union rules."],
      ["D  ·  Australia", "75%", "The world’s first plain packs (2012)."],
    ].forEach(([h, n, b], i) => {
      const x = REVEAL[i].x, w = REVEAL[i].w;
      txt(s, h, { x, y: 3.95, w, h: 0.35, fontSize: 16, bold: true, color: RED });
      txt(s, n, { x, y: 4.35, w, h: 0.7, fontSize: 36, bold: true, color: NAVY });
      txt(s, b, { x, y: 5.15, w, h: 0.9, fontSize: 15, color: BODY });
    });
    source(s, "Warning size = share of the pack front. Sources: Canadian Cancer Society (2025); Tobacco Control Laws (2025).");
    s.addNotes("(~45 s) A is the USA: the brand still owns the pack. B is Hungary, the pack you see in every Nemzeti Dohánybolt. C is Kyrgyzstan: half of the pack is a picture warning, under the common rules of the Eurasian Economic Union. D is Australia, the first country with plain packs. Same product, often the same company, but four completely different rulebooks. And that is the story I want to tell today: how the industry keeps selling when, one by one, the doors close.");
  }

  // ===== 4. The marketer's problem =====
  {
    const s = addContent("Prologue", "The marketer’s problem");
    txt(s, "How do you sell a product to 1.2 billion people when most of the world won’t let you advertise it?", { x: L, y: 1.7, w: W, h: 1.6, fontSize: 32, bold: true, color: NAVY });
    const cw = (W - 0.6) / 3;
    [
      ["1.2 bn", "people use tobacco: about 1 adult in 5", "WHO, 2025"],
      ["80%", "of them live in low- and middle-income countries", "WHO"],
      ["183", "parties to the WHO treaty that bans tobacco ads", "WHO FCTC"],
    ].forEach(([n, l, src], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 3.9, cw, 2.3, { fill: TINT, line: TINT });
      txt(s, n, { x: x + 0.35, y: 4.1, w: cw - 0.7, h: 0.85, fontSize: 44, bold: true, color: RED });
      txt(s, l, { x: x + 0.35, y: 5.0, w: cw - 0.7, h: 0.75, fontSize: 16, color: BODY });
      txt(s, src, { x: x + 0.35, y: 5.8, w: cw - 0.7, h: 0.25, fontSize: 10, color: GREY, italic: true });
    });
    s.addNotes("(~40 s) Imagine you are the marketing manager. Your product has 1.2 billion customers, most of them in developing countries. But almost every country in the world has signed a treaty that says: you may not advertise. So how do you market it? Let me tell this story in three chapters.");
  }

  // ===== 5. Doors close =====
  {
    const s = addContent("Chapter 1  ·  Closing the doors", "One by one, the doors closed");
    const ms = [
      ["2003", "WHO treaty", "Tobacco ads banned"],
      ["2009", "USA", "Flavoured cigarettes banned (not menthol)"],
      ["2012", "Australia", "World’s first plain packs"],
      ["2020", "European Union", "Menthol banned"],
      ["2026", "United Kingdom", "Nobody born after 2008 can ever buy"],
    ];
    const x0 = 2.0, x1 = 11.6, ly = 2.9, step = (x1 - x0) / (ms.length - 1);
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: LIGHT, width: 3 } });
    ms.forEach(([d, t, e], i) => {
      const cx = x0 + i * step, last = i === ms.length - 1, col = last ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.19, y: ly - 0.19, w: 0.38, h: 0.38, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x: cx - 1.1, y: ly - 1.0, w: 2.2, h: 0.55, fontSize: 26, bold: true, color: col, align: "center" });
      txt(s, t, { x: cx - 1.1, y: ly + 0.45, w: 2.2, h: 0.4, fontSize: 17, bold: true, color: col, align: "center" });
      txt(s, e, { x: cx - 1.1, y: ly + 0.9, w: 2.2, h: 0.85, fontSize: 15, color: BODY, align: "center" });
    });
    banner(s, "Every closed door pushed marketers to look for another one.", 5.55, 0.85, 20);
    s.addNotes("(~45 s) Chapter one: the doors close. In 2003 the WHO treaty bans advertising. In 2009 the US bans flavoured cigarettes, except menthol. In 2012 Australia takes the logo off the pack. In 2020 the EU bans menthol. And this year the UK passed a law that nobody born after 2008 can ever legally buy tobacco. Every time a door closed, marketers looked for the next one.");
  }

  // ===== 6. Did it work? =====
  {
    const s = addContent("Chapter 1  ·  Closing the doors", "Did it work? Teenagers stopped smoking");
    s.addChart(pres.charts.BAR, [{ name: "Smokers", labels: ["1999", "2011", "2014", "2019"], values: [28.5, 15.8, 9.2, 5.8] }], {
      ...chartBase(), x: L - 0.1, y: 1.4, w: 6.9, h: 3.85, barDir: "col", barGapWidthPct: 55,
      chartColors: ["0E2841"], showTitle: true, title: "US high-school students who smoke cigarettes, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisLabelFormatCode: "0", valAxisMaxVal: 35, valAxisMinVal: 0,
    });
    source(s, "CDC / FDA National Youth Tobacco Survey (past-30-day use).", L, 5.3, 6.8);
    const sx = 8.3, sw = R - sx;
    txt(s, "÷5", { x: sx, y: 1.6, w: sw, h: 0.8, fontSize: 44, bold: true, color: RED });
    txt(s, "five times fewer teenage smokers in 20 years", { x: sx, y: 2.4, w: sw, h: 0.7, fontSize: 16, color: BODY });
    txt(s, "−17%", { x: sx, y: 3.3, w: sw, h: 0.8, fontSize: 44, bold: true, color: RED });
    txt(s, "lower chance that a teenager smokes after the 2009 US flavour ban", { x: sx, y: 4.1, w: sw, h: 0.7, fontSize: 16, color: BODY });
    txt(s, "Courtemanche, Palmer & Pesko (2017)", { x: sx, y: 4.85, w: sw, h: 0.25, fontSize: 10, color: GREY, italic: true });
    banner(s, [{ text: "Cigarettes lost the teenagers.  " }, { text: "But nicotine didn’t.", options: { color: "F2B8B8" } }], 5.75, 0.8, 20);
    s.addNotes("(~50 s) Did it work? Look at American high-school students. In 1999, almost 3 in 10 smoked. By 2019 it was under 6 in 100: five times fewer. Researchers found that the 2009 flavour ban alone cut the chance that a teenager smokes by about 17%. So the cigarette lost the teenagers. But, and this is the twist, nicotine didn't. Before we get there, let's look at what happened to the pack.");
  }

  // ===== 7. The pack =====
  {
    const s = addContent("Chapter 1  ·  Closing the doors", "The pack became the last billboard");
    const cw = (W - 0.4) / 2;
    [
      ["European Union", "65%", "of the pack is a picture warning", ["Plain packs in France, Ireland, Hungary…", "No flavours at all, menthol included"], NAVY],
      ["United States", "Text only", "a small warning on the side", ["Graphic warnings stuck in court since 2020", "Menthol is still allowed"], RED],
    ].forEach(([h, n, sub, pts, col], i) => {
      const x = L + i * (cw + 0.4);
      s.addText(h, { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y: 1.5, w: cw, h: 0.6, fill: { color: col }, color: WHITE, bold: true, fontSize: 20, valign: "middle", margin: 16 });
      card(s, x, 2.1, cw, 2.95);
      txt(s, n, { x: x + 0.4, y: 2.35, w: cw - 0.8, h: 0.9, fontSize: 48, bold: true, color: col });
      txt(s, sub, { x: x + 0.4, y: 3.25, w: cw - 0.8, h: 0.4, fontSize: 17, color: BODY });
      s.addText(pts.map((p, j) => ({ text: p, options: { bullet: true, breakLine: j < pts.length - 1 } })), { isTextBox: true, x: x + 0.4, y: 3.95, w: cw - 0.8, h: 1.3, fontSize: 16, color: BODY, margin: 0, valign: "top", paraSpaceAfter: 8 });
    });
    txt(s, "World record: in Türkiye and Timor-Leste, 92.5% of the pack is warning.", { x: L, y: 5.4, w: W, h: 0.4, fontSize: 16, color: GREY, italic: true });
    source(s, "Sources: EU Directive 2014/40/EU; US FDA; Canadian Cancer Society (2025).");
    s.addNotes("(~40 s) When there is no TV, no billboard and no magazine, the pack itself is the last billboard. That's why regulators fight over it. In the EU, two-thirds of the pack is a warning, and in Hungary the logo is gone too. In the US, the brand still owns the pack: just a small text warning on the side, and the picture warnings have been stuck in court since 2020. The record holders are Türkiye and Timor-Leste, where 92.5% of the pack is warning.");
  }

  // ===== 8. Flavours of the world =====
  {
    const s = addContent("Chapter 2  ·  The sweet door", "Every country has its own taste", "CONTENT_PICTURE");
    previewSlots(s, [SIDE]);
    const rows = [
      ["Kyrgyzstan", "Nasvay: local chewing tobacco, the only smokeless product allowed."],
      ["Japan", "Menthol and “ice” flavours are mainstream."],
      ["Latin America", "Capsules: click the filter, change the flavour."],
      ["USA", "Menthol: used by ~85% of Black smokers after years of targeted ads."],
    ];
    const lw = 8.0, rh = 1.1, gap = 0.17;
    rows.forEach(([h, b], i) => {
      const y = 1.55 + i * (rh + gap);
      card(s, L, y, lw, rh, { fill: i % 2 ? WHITE : TINT, line: i % 2 ? LIGHT : TINT });
      txt(s, h, { x: L + 0.3, y: y + 0.1, w: 2.1, h: rh - 0.2, fontSize: 18, bold: true, color: RED, valign: "middle" });
      txt(s, b, { x: L + 2.45, y: y + 0.1, w: lw - 2.7, h: rh - 0.2, fontSize: 16, color: BODY, valign: "middle" });
    });
    txt(s, "Flavour is the easiest way to adapt the product to a culture.", { x: L, y: 6.6, w: lw, h: 0.3, fontSize: 14, color: GREY, italic: true });
    s.addNotes("(~45 s) Chapter two: the sweet door. If you can't change the advertising, you change the product. In Kyrgyzstan the local habit is nasvay, a chewing tobacco: it is the only smokeless product the law still allows, and it has not been taxed like cigarettes. In Japan menthol is mainstream. In Latin America, capsule cigarettes let you click a ball in the filter and change the taste. And in the US, menthol was advertised so heavily in Black neighbourhoods that today about 85% of Black smokers use menthol.\n\nPicture suggestion: a capsule ('click') pack or a pack of nasvay.");
  }

  // ===== 9. Why flavours get banned =====
  {
    const s = addContent("Chapter 2  ·  The sweet door", "Why governments ban flavours");
    const lw = 6.2;
    [
      [I.wind, "They hide the harshness", "Menthol numbs the throat. The first cigarette doesn’t hurt."],
      [I.user, "They attract beginners", "Sweet tastes make perfect starter products."],
      [I.target, "They target groups", "Flavours were aimed at specific communities."],
    ].forEach(([img, h, b], i) => {
      const y = 1.55 + i * 1.6;
      card(s, L, y, lw, 1.4, { fill: TINT, line: TINT });
      iconCircle(s, img, L + 0.3, y + 0.39, 0.62);
      txt(s, h, { x: L + 1.2, y: y + 0.22, w: lw - 1.4, h: 0.4, fontSize: 19, bold: true, color: NAVY });
      txt(s, b, { x: L + 1.2, y: y + 0.66, w: lw - 1.4, h: 0.6, fontSize: 15, color: BODY });
    });
    const sx = 7.6, sw = R - sx;
    card(s, sx, 1.55, sw, 4.6, { line: RED, lineW: 1.25 });
    txt(s, "The clove war", { x: sx + 0.35, y: 1.75, w: sw - 0.7, h: 0.45, fontSize: 22, bold: true, color: RED });
    [["2009", "USA bans clove cigarettes, but keeps menthol"], ["2010", "Indonesia sues the USA at the WTO"], ["2012", "Indonesia wins: it was discrimination"]].forEach(([d, t], i) => {
      const y = 2.5 + i * 1.15;
      txt(s, d, { x: sx + 0.35, y, w: 1.0, h: 0.45, fontSize: 20, bold: true, color: NAVY });
      txt(s, t, { x: sx + 1.4, y: y + 0.03, w: sw - 1.75, h: 0.85, fontSize: 16, color: BODY });
    });
    txt(s, "Health policy became trade policy.", { x: sx + 0.35, y: 5.6, w: sw - 0.7, h: 0.35, fontSize: 15, color: GREY, italic: true });
    s.addNotes("(~50 s) Why do governments ban flavours? Menthol numbs the throat, so the first cigarette doesn't hurt. Sweet flavours make perfect starter products. And they were used to target specific groups. Here's my favourite story. In 2009 the US banned clove cigarettes, but kept menthol. Who makes clove cigarettes? Indonesia. Who makes menthol? American companies. Indonesia sued at the WTO and won in 2012. Health policy became trade policy, which is exactly why international marketers must watch regulation.");
  }

  // ===== 10. Then came the vape =====
  {
    const s = addContent("Chapter 3  ·  The new door", "Then came the vape");
    s.addChart(pres.charts.BAR, [{ name: "Vapers", labels: ["2011", "2014", "2018", "2019", "2022", "2024"], values: [1.5, 13.4, 20.8, 27.5, 14.1, 7.8] }], {
      ...chartBase(), x: L - 0.1, y: 1.4, w: 7.4, h: 4.0, barDir: "col", barGapWidthPct: 45,
      chartColors: ["0E2841", "0E2841", "0E2841", "C00000", "0E2841", "0E2841"], showTitle: true, title: "US high-school students who vape, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisLabelFormatCode: "0", valAxisMaxVal: 32, valAxisMinVal: 0,
    });
    source(s, "CDC / FDA National Youth Tobacco Survey. The 2019 survey moved to tablets, which may raise estimates.", L, 5.45, 7.3);
    const sx = 8.8, sw = R - sx;
    [["2014", "Vapes overtake cigarettes among US teens", NAVY], ["2019", "Peak: more than 1 in 4 high-schoolers vape", RED], ["2020", "FDA restricts flavoured cartridges, and use falls", NAVY]].forEach(([d, t, col], i) => {
      const y = 1.55 + i * 1.25;
      txt(s, d, { x: sx, y, w: sw, h: 0.5, fontSize: 24, bold: true, color: col });
      txt(s, t, { x: sx, y: y + 0.5, w: sw, h: 0.65, fontSize: 15, color: BODY });
    });
    banner(s, [{ text: "Today: " }, { text: "~9 in 10 young vapers use flavours.", options: { color: "F2B8B8" } }, { text: "  Worldwide, 15 million kids aged 13–15 vape." }], 5.85, 0.8, 17);
    s.addNotes("(~50 s) Chapter three: the new door. Just as teenagers stopped smoking, the vape arrived. In 2011 only 1.5% of American high-schoolers vaped. In 2014 vapes overtook cigarettes. By 2019, more than one in four high-schoolers vaped. That is more than ever smoked cigarettes in the last decade. Then regulators reacted: in 2020 the FDA restricted flavoured cartridges, and use fell to under 8% by 2024. But even today about nine in ten young vapers use flavours, and worldwide 15 million children aged 13 to 15 vape.");
  }

  // ===== 11. Why vapes speak to kids =====
  {
    const s = addContent("Chapter 3  ·  The new door", "Why vapes speak to kids");
    const cw = (W - 0.9) / 4;
    [
      [I.candy, "Flavours", "Bubblegum, mango, cotton candy"],
      [I.usb, "Design", "Colourful, looks like a USB stick, easy to hide"],
      [I.phone, "Social media", "Influencers and TikTok, not TV"],
      [I.coin, "Price", "Cheap disposables sold at every kiosk"],
    ].forEach(([img, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 3.0);
      iconCircle(s, img, x + 0.35, 1.85, 0.75, i === 0 ? RED : NAVY);
      txt(s, h, { x: x + 0.35, y: 2.8, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: i === 0 ? RED : NAVY });
      txt(s, b, { x: x + 0.35, y: 3.3, w: cw - 0.7, h: 1.1, fontSize: 16, color: BODY });
    });
    s.addText([
      { text: "$462 million", options: { bold: true, fontSize: 30, color: WHITE } },
      { text: "   paid by JUUL (2023) to US states over marketing to young people", options: { fontSize: 17, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.85, w: W, h: 1.0, fill: { color: NAVY }, valign: "middle", margin: 16 });
    txt(s, "“A candy-flavoured trap.”  — WHO Director-General, 2024", { x: L, y: 6.05, w: W, h: 0.4, fontSize: 16, color: RED, italic: true, bold: true });
    s.addNotes("(~50 s) Why did vapes reach kids so fast? Four reasons, and they are pure marketing. Flavours: bubblegum, mango, cotton candy. Design: colourful, looks like a USB stick, easy to hide from parents and teachers. Social media: influencers and TikTok, where tobacco ad bans were never written for. And price: cheap disposables at every kiosk. The companies say vapes are for adult smokers who want to quit. But JUUL paid 462 million dollars to six US states over marketing to young people, and the head of the WHO called it a 'candy-flavoured trap'.");
  }

  // ===== 12. 4Ps =====
  {
    const s = addContent("What marketers learned", "Marketing without advertising: the 4Ps");
    const cw = (W - 0.9) / 4;
    [
      [I.box, "Product", "New formats: capsules, vapes, pouches"],
      [I.tag, "Price", "Single sticks and cheap disposables"],
      [I.store, "Place", "Kiosks, warungs, duty-free, online"],
      [I.mega, "Promotion", "Sponsorships, influencers, the pack"],
    ].forEach(([img, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 3.0, { fill: TINT, line: TINT });
      iconCircle(s, img, x + 0.35, 1.85, 0.75);
      txt(s, h, { x: x + 0.35, y: 2.8, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.35, y: 3.3, w: cw - 0.7, h: 1.1, fontSize: 16, color: BODY });
    });
    txt(s, "Big Tobacco now owns the new door too", { x: L, y: 4.9, w: W, h: 0.4, fontSize: 18, bold: true, color: NAVY });
    txt(s, [{ text: "41.5%", options: { bold: true, color: RED, fontSize: 30 } }, { text: "  of Philip Morris revenue now comes from smoke-free products (IQOS, ZYN)", options: { color: BODY, fontSize: 16 } }], { x: L, y: 5.35, w: W, h: 0.9, valign: "middle" });
    source(s, "PMI 2025 full-year results.");
    s.addNotes("(~40 s) So what did marketers learn? When mass media is closed, the other three Ps do the work. Product: new formats. Price: single sticks and cheap disposables. Place: kiosks and online. Promotion: sponsorships, influencers and the pack. And the big companies are not losing this game: Philip Morris already earns 41.5% of its revenue from heated tobacco and nicotine pouches.");
  }

  // ===== 13. Epilogue =====
  {
    const s = addContent("Epilogue", "The story in three chapters");
    const cw = (W - 0.6) / 3;
    [
      ["1", "The doors closed", "Ad bans and plain packs: teenage smoking fell five times."],
      ["2", "The sweet door", "Flavours adapted the product to every culture, until they were banned."],
      ["3", "The new door", "Vapes reached kids through flavour, design and social media."],
    ].forEach(([n, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 2.6);
      txt(s, n, { x: x + 0.35, y: 1.7, w: 1, h: 0.75, fontSize: 40, bold: true, color: RED });
      txt(s, h, { x: x + 0.35, y: 2.5, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.35, y: 3.0, w: cw - 0.7, h: 1.0, fontSize: 15, color: BODY });
    });
    txt(s, "“Tobacco marketing didn’t disappear. It moved: into the pack, the flavour and the phone.”", { x: L, y: 4.3, w: W, h: 0.85, fontSize: 22, bold: true, color: RED, italic: true });
    s.addText([
      { text: "Question for you", options: { bold: true, fontSize: 14, color: LIGHT, breakLine: true } },
      { text: "Kyrgyzstan banned nicotine vapes completely in 2025. Should others follow, or are plain packs and no flavours enough?", options: { fontSize: 20, bold: true, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.3, w: W, h: 1.4, fill: { color: NAVY }, valign: "middle", margin: 18, paraSpaceAfter: 4 });
    s.addNotes("(~40 s) So, three chapters. First, the doors closed, and teenagers stopped smoking. Second, flavours became the sweet door, until governments closed that too. Third, vapes opened a new door straight into teenagers' phones. My conclusion: tobacco marketing didn't disappear, it moved. And my question for you: since July 2025 Kyrgyzstan has banned the import, sale and use of nicotine vapes completely. Should other countries follow, or are plain packs and no flavours enough? Thank you.");
  }

  // ===== 14. Bibliography =====
  {
    const s = addContent("Sources", "Bibliography");
    const refs = [
      "Canadian Cancer Society (2025). Cigarette Package Health Warnings: International Status Report, 9th ed.",
      "CDC (2015). Tobacco use among middle and high school students, 2011–2014. MMWR 64(14).",
      "CDC (2019). Tobacco product use among middle and high school students, 2019. MMWR 68(SS-12).",
      "CDC / FDA (2024). Tobacco product use among middle and high school students, 2024. MMWR 73(41).",
      "Courtemanche, C., Palmer, M., Pesko, M. (2017). Influence of the flavored cigarette ban on adolescent tobacco use. American Journal of Preventive Medicine.",
      "Directive 2014/40/EU (Tobacco Products Directive). Official Journal of the EU.",
      "FDA (2026). 2025 National Youth Tobacco Survey findings.",
      "GOV.UK (2026). Tobacco and Vapes Bill becomes law, 29 April 2026.",
      "Keegan, W. J., & Green, M. C. (2020). Global Marketing, 10th ed. Pearson.",
      "Philip Morris International (2026). 2025 Fourth-Quarter & Full-Year Results.",
      "Tobacco Control Laws, Campaign for Tobacco-Free Kids (2025). Kyrgyzstan: packaging and labeling.",
      "24.kg (2024). E-cigarettes banned in Kyrgyzstan: law to come into force in July 2025.",
      "WHO (2024). Hooking the next generation: how the tobacco industry captures young customers.",
      "WHO (2025). Global report on trends in prevalence of tobacco use 2000–2024.",
      "WTO (2012). US — Clove Cigarettes (DS406); New York AG (2023). JUUL settlement.",
    ];
    const half = Math.ceil(refs.length / 2);
    const colW = (W - 0.4) / 2;
    [refs.slice(0, half), refs.slice(half)].forEach((col, i) => {
      s.addText(col.map((r, j) => ({ text: r, options: { breakLine: j < col.length - 1 } })), {
        isTextBox: true, x: L + i * (colW + 0.4), y: 1.45, w: colW, h: 5.35, fontSize: 13, color: INK, margin: 0, valign: "top", paraSpaceAfter: 8,
      });
    });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);

  // pptxgenjs writes picture placeholders without type="pic"; add it so PowerPoint shows the picture icon
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const names = Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts)\/slide(Layout)?\d+\.xml$/.test(n));
  for (const n of names) {
    const xml = await zip.file(n).async("string");
    const fixed = xml.replace(/<p:ph\s+idx="(\d+)"\s+hasCustomPrompt="1"\s*\/>/g, '<p:ph type="pic" idx="$1" hasCustomPrompt="1"/>');
    if (fixed !== xml) zip.file(n, fixed);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);
})();
