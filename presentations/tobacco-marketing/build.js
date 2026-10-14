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
    txt(s, "How to market a product nobody lets you advertise", { x: 0.8, y: 4.15, w: 8.0, h: 0.9, fontSize: 22, color: LIGHT });
    txt(s, "Corvinus University of Budapest  ·  Fall 2026", { x: 0.8, y: 6.35, w: 7.4, h: 0.4, fontSize: 14, color: LIGHT });
    s.addShape(pres.shapes.OVAL, { x: 9.95, y: 2.0, w: 2.0, h: 2.0, fill: { color: "1B3A5C" }, line: { color: WHITE, width: 1.5 }, objectName: "Photo Amir Kenenbaev" });
    txt(s, "Amir Kenenbaev", { x: 9.65, y: 4.2, w: 2.6, h: 0.4, fontSize: 15, color: WHITE, bold: true, align: "center" });
    s.addNotes("(~15 s) Good afternoon. I'm Amir, and my topic is tobacco marketing on international markets: not the business side, but the marketing itself. Before I say anything, I want to show you some packs.\n\nPhoto circle: right-click the circle > Format Shape > Fill > Picture fill to insert a portrait.");
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
    const s = addContent("Prologue", "Same cigarettes, four ways to brand a pack", "PACK_REVEAL");
    previewSlots(s, REVEAL);
    [
      ["A  ·  USA", "100%", "The whole front is still brand: logo, colours, design."],
      ["B  ·  Hungary", "35%", "No logo, no colours: only the name in a plain font."],
      ["C  ·  Kyrgyzstan", "50%", "Half the front for the brand, half for a picture warning."],
      ["D  ·  Australia", "25%", "The first plain packs (2012): name only, drab brown."],
    ].forEach(([h, n, b], i) => {
      const x = REVEAL[i].x, w = REVEAL[i].w;
      txt(s, h, { x, y: 3.95, w, h: 0.35, fontSize: 16, bold: true, color: RED });
      txt(s, n, { x, y: 4.35, w, h: 0.7, fontSize: 36, bold: true, color: NAVY });
      txt(s, b, { x, y: 5.15, w, h: 0.9, fontSize: 15, color: BODY });
    });
    source(s, "Share of the pack front left to the brand after the health warning. Sources: Canadian Cancer Society (2025); Tobacco Control Laws (2025).");
    s.addNotes("(~45 s) Here are the answers, and look at them as a marketer: how much of the pack is still yours? A is the USA: the whole front is still brand. B is Hungary, the pack you see in every Nemzeti Dohánybolt: no logo, no colours, just the name. C is Kyrgyzstan: half for the brand, half for the warning. D is Australia, the first plain pack: a quarter of the front, in the least attractive colour researchers could find. Same cigarettes, four completely different marketing toolboxes. That is today's story: how tobacco marketing keeps working when, one by one, its tools are taken away.");
  }

  // ===== 4. The golden age =====
  {
    const s = addContent("Prologue", "The golden age of tobacco advertising");
    const cw = (W - 0.6) / 3;
    [
      ["1954", "The Marlboro Man", "A filter cigarette once sold to women as “Mild as May” became a symbol of the rugged cowboy."],
      ["1968", "Formula 1", "Gold Leaf Lotus started it. Marlboro, Camel and others followed on the fastest cars in the world."],
      ["1988", "Joe Camel", "A cartoon camel. 91% of 6-year-olds recognised him, as many as knew Mickey Mouse."],
    ].forEach(([d, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.5, cw, 3.55);
      txt(s, d, { x: x + 0.35, y: 1.7, w: cw - 0.7, h: 0.75, fontSize: 36, bold: true, color: RED });
      txt(s, h, { x: x + 0.35, y: 2.55, w: cw - 0.7, h: 0.45, fontSize: 21, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.35, y: 3.1, w: cw - 0.7, h: 1.8, fontSize: 16, color: BODY });
    });
    banner(s, [{ text: "Then the ads were banned.  " }, { text: "How do you market a product you can’t advertise?", options: { color: "F2B8B8" } }], 5.4, 0.95, 20);
    source(s, "Fischer et al., JAMA (1991); Marlboro Man: Leo Burnett for Philip Morris.", L, 6.5);
    s.addNotes("(~50 s) To understand tobacco marketing today, remember how powerful it used to be. In 1954 the Leo Burnett agency took Marlboro, a filter cigarette sold to women with the slogan 'Mild as May', and turned it into the Marlboro Man, one of the most successful repositionings in marketing history. From 1968, when Gold Leaf sponsored Team Lotus, cigarette brands were painted on Formula 1 cars. And in 1988 came Joe Camel, a cartoon: a 1991 study found that 91% of six-year-olds recognised him, as many as recognised Mickey Mouse. Then all of this was banned. So how do you market a product you're not allowed to advertise?");
  }

  // ===== 5. Doors close =====
  {
    const s = addContent("Chapter 1  ·  Closing the doors", "One by one, marketers lost their tools");
    const ms = [
      ["2003", "WHO treaty", "Lost: ads, promotion, sponsorship"],
      ["2009", "USA", "Lost: sweet flavours"],
      ["2012", "Australia", "Lost: the logo on the pack"],
      ["2020", "European Union", "Lost: menthol"],
      ["2025", "Kyrgyzstan", "Lost: the whole vape category"],
      ["2026", "United Kingdom", "Lost: a whole generation of customers"],
    ];
    const x0 = 1.9, x1 = 11.7, ly = 2.9, step = (x1 - x0) / (ms.length - 1);
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: LIGHT, width: 3 } });
    ms.forEach(([d, t, e], i) => {
      const cx = x0 + i * step, last = i === ms.length - 1, col = last ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.19, y: ly - 0.19, w: 0.38, h: 0.38, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x: cx - 0.95, y: ly - 1.0, w: 1.9, h: 0.55, fontSize: 26, bold: true, color: col, align: "center" });
      txt(s, t, { x: cx - 0.95, y: ly + 0.45, w: 1.9, h: 0.4, fontSize: 16, bold: true, color: col, align: "center" });
      txt(s, e, { x: cx - 0.95, y: ly + 0.9, w: 1.9, h: 0.85, fontSize: 14, color: BODY, align: "center" });
    });
    banner(s, "Every lost tool pushed marketers to look for another one.", 5.55, 0.85, 20);
    s.addNotes("(~45 s) Chapter one: the marketing toolbox is taken away, piece by piece. In 2003 the WHO treaty bans advertising, promotion and sponsorship. In 2009 the US bans sweet flavours in cigarettes. In 2012 Australia takes the logo off the pack. In 2020 the EU bans menthol. In 2025 Kyrgyzstan bans nicotine vapes completely. And this year the UK passed a law that anyone born after 2008 can never legally buy tobacco: a whole generation of customers gone. Every time a tool was lost, marketers looked for another one.");
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
    s.addNotes("(~50 s) Did it work? Look at American high-school students. In 1999, almost 3 in 10 smoked. By 2019 it was under 6 in 100: five times fewer. Researchers found that the 2009 flavour ban alone cut the chance that a teenager smokes by about 17%. Without ads, cartoons and sweet flavours, cigarettes lost the teenagers. But, and this is the twist, nicotine didn't. Before we get there, let's look at the last marketing tool cigarettes still have: the pack.");
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
      ["Middle East", "Flavoured shisha: apple, grape, mint."],
    ];
    const lw = 8.0, rh = 1.1, gap = 0.17;
    rows.forEach(([h, b], i) => {
      const y = 1.55 + i * (rh + gap);
      card(s, L, y, lw, rh, { fill: i % 2 ? WHITE : TINT, line: i % 2 ? LIGHT : TINT });
      txt(s, h, { x: L + 0.3, y: y + 0.1, w: 2.1, h: rh - 0.2, fontSize: 18, bold: true, color: RED, valign: "middle" });
      txt(s, b, { x: L + 2.45, y: y + 0.1, w: lw - 2.7, h: rh - 0.2, fontSize: 16, color: BODY, valign: "middle" });
    });
    txt(s, "When you can’t change the ad, you change the product.", { x: L, y: 6.6, w: lw, h: 0.3, fontSize: 14, color: GREY, italic: true });
    s.addNotes("(~40 s) Chapter two: the sweet door. If you can't change the advertising, you change the product, and flavour is the easiest way to fit a local culture. In Kyrgyzstan the local habit is nasvay, a chewing tobacco: the only smokeless product the law still allows. In Japan menthol and 'ice' flavours are mainstream. In Latin America, capsule cigarettes let you click a ball in the filter and change the taste: product innovation as entertainment. In the Middle East it's flavoured shisha.\n\nPicture suggestion: a capsule ('click') pack or a pack of nasvay.");
  }

  // ===== 9. Menthol: targeted marketing =====
  {
    const s = addContent("Chapter 2  ·  The sweet door", "Why flavours work, and why they get banned");
    const lw = 6.2;
    [
      [I.wind, "They hide the harshness", "Menthol numbs the throat. The first cigarette doesn’t hurt."],
      [I.user, "They attract beginners", "Sweet tastes make perfect starter products."],
      [I.target, "They target groups", "One flavour can be built for one community."],
    ].forEach(([img, h, b], i) => {
      const y = 1.55 + i * 1.6;
      card(s, L, y, lw, 1.4, { fill: TINT, line: TINT });
      iconCircle(s, img, L + 0.3, y + 0.39, 0.62);
      txt(s, h, { x: L + 1.2, y: y + 0.22, w: lw - 1.4, h: 0.4, fontSize: 19, bold: true, color: NAVY });
      txt(s, b, { x: L + 1.2, y: y + 0.66, w: lw - 1.4, h: 0.6, fontSize: 15, color: BODY });
    });
    const sx = 7.6, sw = R - sx;
    card(s, sx, 1.55, sw, 4.6, { line: RED, lineW: 1.25 });
    txt(s, "Case: menthol in the USA", { x: sx + 0.35, y: 1.75, w: sw - 0.7, h: 0.45, fontSize: 21, bold: true, color: RED });
    s.addText([
      { text: "Ads in Black magazines", options: { bullet: true, breakLine: true } },
      { text: "Sponsorship: the Kool Jazz Festival", options: { bullet: true, breakLine: true } },
      { text: "Free samples in neighbourhoods", options: { bullet: true } },
    ], { isTextBox: true, x: sx + 0.35, y: 2.4, w: sw - 0.7, h: 1.6, fontSize: 16, color: BODY, margin: 0, valign: "top", paraSpaceAfter: 10 });
    txt(s, "~85%", { x: sx + 0.35, y: 4.1, w: sw - 0.7, h: 0.75, fontSize: 40, bold: true, color: NAVY });
    txt(s, "of Black American smokers use menthol today (FDA)", { x: sx + 0.35, y: 4.85, w: sw - 0.7, h: 0.7, fontSize: 15, color: BODY });
    s.addNotes("(~50 s) Why do flavours work so well? Menthol numbs the throat, so the first cigarette doesn't hurt. Sweet flavours make perfect starter products. And a flavour can be built for one specific community. The clearest case is menthol in the United States. From the 1950s, companies advertised menthol brands in Black magazines, sponsored events like the Kool Jazz Festival, and handed out free samples in Black neighbourhoods. Today about 85% of Black American smokers use menthol. That is targeted marketing at its most effective, and exactly why governments started banning flavours.");
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
    txt(s, "The Joe Camel playbook, moved to a new channel: the phone.", { x: L, y: 6.05, w: W, h: 0.4, fontSize: 16, color: RED, italic: true, bold: true });
    s.addNotes("(~50 s) Why did vapes reach kids so fast? Four reasons, and every one of them is a marketing decision. Flavours: bubblegum, mango, cotton candy. Design: colourful, looks like a USB stick, easy to hide from parents and teachers. Social media: influencers and TikTok, a channel the old ad bans were never written for. And price: cheap disposables at every kiosk. The companies say vapes are for adult smokers who want to quit. But JUUL paid 462 million dollars to US states over marketing to young people, and the WHO calls it a 'candy-flavoured trap'. It is the Joe Camel playbook, just moved to the phone.");
  }

  // ===== 12. 4Ps =====
  {
    const s = addContent("What marketers learned", "Marketing without advertising: the 4Ps");
    const cw = (W - 0.9) / 4;
    [
      [I.box, "Product", "New formats: capsules, vapes, pouches"],
      [I.tag, "Price", "Single sticks and cheap disposables"],
      [I.store, "Place", "Kiosks, shop displays, duty-free, online"],
      [I.mega, "Promotion", "Sponsorships, influencers, the pack"],
    ].forEach(([img, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 3.0, { fill: TINT, line: TINT });
      iconCircle(s, img, x + 0.35, 1.85, 0.75);
      txt(s, h, { x: x + 0.35, y: 2.8, w: cw - 0.7, h: 0.45, fontSize: 20, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.35, y: 3.3, w: cw - 0.7, h: 1.1, fontSize: 16, color: BODY });
    });
    txt(s, "Where promotion moved", { x: L, y: 4.85, w: W, h: 0.4, fontSize: 18, bold: true, color: NAVY });
    const chip = (t, x, y, w, fill, color) => s.addText(t, { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y, w, h: 0.6, fill: { color: fill }, line: { color: fill, width: 0 }, fontSize: 17, bold: true, color, align: "center", valign: "middle", margin: 0 });
    const cwid = 1.55, gap = 0.12;
    ["TV", "Billboards", "Magazines"].forEach((t, i) => chip(t, L + i * (cwid + gap), 5.45, cwid, TINT, GREY));
    const ax = L + 3 * cwid + 2 * gap + 0.2;
    s.addShape(pres.shapes.LINE, { x: ax, y: 5.75, w: 0.65, h: 0, line: { color: RED, width: 2.5, endArrowType: "triangle" } });
    ["The pack", "The shop", "The phone"].forEach((t, i) => chip(t, ax + 0.85 + i * (cwid + gap), 5.45, cwid, NAVY, WHITE));
    s.addNotes("(~40 s) So what did marketers learn? When mass media is closed, the other three Ps do the work. Product: new formats like capsules, vapes and pouches. Price: single sticks and cheap disposables that fit a teenager's pocket money. Place: kiosks, shop displays, duty-free and online. And promotion didn't disappear, it moved: from TV, billboards and magazines to the pack, the shop and the phone.");
  }

  // ===== 13. Epilogue =====
  {
    const s = addContent("Epilogue", "The story in three chapters");
    const cw = (W - 0.6) / 3;
    [
      ["1", "The tools were taken", "Ads, logos and sponsorship disappeared, and teenage smoking fell five times."],
      ["2", "The sweet door", "Flavours fitted the product to every culture and every target group."],
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
      { text: "Where is the line between marketing a legal product and marketing to children?", options: { fontSize: 20, bold: true, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.3, w: W, h: 1.4, fill: { color: NAVY }, valign: "middle", margin: 18, paraSpaceAfter: 4 });
    s.addNotes("(~40 s) So, three chapters. First, the marketing tools were taken away, and teenagers stopped smoking. Second, flavours became the sweet door, fitting the product to every culture and every target group. Third, vapes opened a new door straight into teenagers' phones. My conclusion: tobacco marketing didn't disappear, it moved. And my question for you: where is the line between marketing a legal product and marketing to children? Thank you.");
  }

  // ===== 14. Bibliography =====
  {
    const s = addContent("Sources", "Bibliography");
    const refs = [
      "Campaign for Tobacco-Free Kids. Tobacco Company Marketing to African Americans (fact sheet).",
      "Canadian Cancer Society (2025). Cigarette Package Health Warnings: International Status Report, 9th ed.",
      "CDC (2015, 2019, 2024). Tobacco product use among middle and high school students. MMWR.",
      "Courtemanche, C., Palmer, M., Pesko, M. (2017). Influence of the flavored cigarette ban on adolescent tobacco use. Am. J. of Preventive Medicine.",
      "Directive 2014/40/EU (Tobacco Products Directive). Official Journal of the EU.",
      "FDA (2026). 2025 National Youth Tobacco Survey findings.",
      "Fischer, P. M. et al. (1991). Brand logo recognition by children aged 3 to 6 years: Mickey Mouse and Old Joe the Camel. JAMA 266(22).",
      "GOV.UK (2026). Tobacco and Vapes Bill becomes law, 29 April 2026.",
      "Keegan, W. J., & Green, M. C. (2020). Global Marketing, 10th ed. Pearson.",
      "NPR (2022). How the tobacco industry targeted Black Americans with menthol smokes.",
      "Tobacco Control Laws (2025). Kyrgyzstan: packaging and labeling; 24.kg (2024). E-cigarettes banned in Kyrgyzstan from July 2025.",
      "WHO (2024). Hooking the next generation: how the tobacco industry captures young customers.",
      "WHO (2025). Global report on trends in prevalence of tobacco use 2000–2024.",
      "New York Attorney General (2023). JUUL settlement with six states and DC.",
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
