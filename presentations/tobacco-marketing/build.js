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
const REVEAL = ["USA", "Hungary", "Indonesia", "Australia"].map((k, i) => ({ name: "photo" + k, label: "Insert photo: " + k + " pack", x: L + i * (REV_W + 0.3), y: 1.45, w: REV_W, h: 2.25 }));
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
    eye: await icon(Fi.FiEye), user: await icon(Fi.FiUsers), target: await icon(Fi.FiTarget),
    box: await icon(Fi.FiPackage), tag: await icon(Fi.FiTag), store: await icon(Fi.FiShoppingBag), mega: await icon(Fi.FiSpeaker),
    globe: await icon(Fi.FiGlobe), wind: await icon(Fi.FiWind), alert: await icon(Fi.FiAlertTriangle),
  };
  const iconCircle = (s, img, x, y, d = 0.62, fill = NAVY) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };

  // ===== 1. Title =====
  {
    const s = pres.addSlide({ masterName: "TITLE" });
    txt(s, "International Marketing  ·  Individual presentation", { x: 0.8, y: 1.15, w: 8.4, h: 0.4, fontSize: 14, color: LIGHT, bold: true, charSpacing: 1 });
    txt(s, "Tobacco Marketing on International Markets", { x: 0.8, y: 1.7, w: 8.2, h: 2.3, fontSize: 40, bold: true, color: WHITE });
    txt(s, "One product, many rulebooks: how the industry sells where it can’t advertise", { x: 0.8, y: 4.15, w: 8.0, h: 0.9, fontSize: 22, color: LIGHT });
    txt(s, "Corvinus University of Budapest  ·  Fall 2026", { x: 0.8, y: 6.35, w: 7.4, h: 0.4, fontSize: 14, color: LIGHT });
    s.addShape(pres.shapes.OVAL, { x: 9.95, y: 2.0, w: 2.0, h: 2.0, fill: { color: "1B3A5C" }, line: { color: WHITE, width: 1.5 }, objectName: "Photo Amir Kenenbaev" });
    txt(s, "Amir Kenenbaev", { x: 9.65, y: 4.2, w: 2.6, h: 0.4, fontSize: 15, color: WHITE, bold: true, align: "center" });
    s.addNotes("(~20 s) Good afternoon. My name is Amir Kenenbaev, and today I'll talk about tobacco marketing on international markets. Before I explain anything, let's start with a small game.\n\nPhoto circle: right-click the circle > Format Shape > Fill > Picture fill to insert a portrait.");
  }

  // ===== 2. Quiz =====
  {
    const s = addContent("Warm-up", "Guess the country: what is different about these packs?", "PACK_QUIZ");
    txt(s, "Look at the warning size, the pictures, the logo and colours, and the flavour.", { x: L, y: 1.3, w: W, h: 0.4, fontSize: 16, color: BODY });
    previewSlots(s, QUIZ);
    QUIZ.forEach((p, i) => {
      const k = "ABCD"[i];
      s.addText(k, { isTextBox: true, shape: pres.shapes.OVAL, x: p.x, y: 5.95, w: 0.5, h: 0.5, fill: { color: RED }, color: WHITE, bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0 });
      txt(s, "Which country?", { x: p.x + 0.65, y: 6.03, w: p.w - 0.65, h: 0.36, fontSize: 15, color: GREY });
    });
    s.addNotes(
      "(~1 min) Here are four cigarette packs from four countries. Two questions: which country is each pack from, and what is different about the packaging? Take 20 seconds and shout out your guesses.\n\n" +
      "PHOTOS TO INSERT (ideally the same brand, e.g. Marlboro, so only the rules differ):\n" +
      "A = USA: full brand design, small text-only Surgeon General warning on the side.\n" +
      "B = Hungary (EU): plain pack + 65% picture-and-text warning in Hungarian.\n" +
      "C = Indonesia: kretek pack (e.g. Dji Sam Soe / Gudang Garam) with a 40% picture warning.\n" +
      "D = Australia: plain drab dark brown pack, 75% warning on the front.\n" +
      "Click the picture icon in each box to insert a photo; it fills and crops automatically.");
  }

  // ===== 3. Reveal =====
  {
    const s = addContent("Warm-up: the answer", "Same product, four different rulebooks", "PACK_REVEAL");
    previewSlots(s, REVEAL);
    const cols = [
      ["A  ·  USA", "Text only", "Small Surgeon General warning on the side. The front is pure brand. Graphic warnings (FDA, 2020) are still blocked in court."],
      ["B  ·  Hungary (EU)", "65%", "Picture + text warning on front and back. Plain packaging for all brands since 2022. No menthol."],
      ["C  ·  Indonesia", "40%", "Picture warning, below the WHO-recommended 50%. Never signed the WHO tobacco treaty. Clove kretek culture."],
      ["D  ·  Australia", "75% / 90%", "World’s first plain packs (2012) in drab dark brown, nicknamed “the ugliest colour in the world”."],
    ];
    cols.forEach(([h, n, b], i) => {
      const x = REVEAL[i].x, w = REVEAL[i].w;
      txt(s, h, { x, y: 3.9, w, h: 0.35, fontSize: 15, bold: true, color: RED });
      txt(s, n, { x, y: 4.27, w, h: 0.6, fontSize: 30, bold: true, color: NAVY });
      txt(s, b, { x, y: 4.95, w, h: 1.5, fontSize: 13, color: BODY });
    });
    source(s, "Warning size = share of the front and back of the pack. Sources: Canadian Cancer Society (2025); TPackSS (2023); Australian Dept. of Health.");
    s.addNotes("(~1 min) Here are the answers. A is the USA: the brand still owns the front of the pack. B is Hungary: 65% warnings and, since 2022, plain packaging, which is what you see in every Nemzeti Dohánybolt. C is Indonesia, home of clove kretek cigarettes, with only 40% warnings. D is Australia, the first country with plain packs in 2012, with a colour chosen because smokers found it the least attractive. Same product, same companies, four completely different rulebooks. That is exactly what international marketing is about: the environment shapes the marketing mix.");
  }

  // ===== 4. Why tobacco is a unique case =====
  {
    const s = addContent("Why this topic", "Why tobacco is a unique international marketing case");
    const stats = [
      ["1.2 bn", "tobacco users worldwide in 2024, about 1 adult in 5", "WHO, 2025"],
      ["80%", "of them live in low- and middle-income countries", "WHO"],
      ["100 m+", "people vape, including 15 million children aged 13–15", "WHO, 2025"],
    ];
    stats.forEach(([n, l, src], i) => {
      const y = 1.45 + i * 1.5;
      txt(s, n, { x: L, y, w: 2.6, h: 0.75, fontSize: 36, bold: true, color: RED });
      txt(s, l, { x: L + 2.75, y: y + 0.12, w: 3.8, h: 0.7, fontSize: 15, color: BODY });
      txt(s, src, { x: L + 2.75, y: y + 0.85, w: 3.8, h: 0.25, fontSize: 10, color: GREY, italic: true });
    });
    const dx = 7.75, dw = R - dx;
    txt(s, "Who makes the world’s cigarettes (volume, 2025)", { x: dx, y: 1.45, w: dw, h: 0.35, fontSize: 14, bold: true, color: NAVY, align: "center" });
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Volume share", labels: ["China National Tobacco", "Big Four (PMI, BAT, JTI, Imperial)", "Others"], values: [47, 34, 19] }], {
      x: dx + 0.6, y: 1.8, w: dw - 1.2, h: 2.6, holeSize: 64, chartColors: ["C00000", "0E2841", "D1D1D1"], showLegend: false, showValue: false, showPercent: false, showTitle: false, dataBorder: { pt: 1, color: "FFFFFF" },
    });
    txt(s, "47%", { x: dx, y: 2.75, w: dw, h: 0.7, fontSize: 28, bold: true, color: RED, align: "center", valign: "middle" });
    [["China National Tobacco", "47%: sold almost only inside China", RED], ["Big Four (PMI, BAT, JTI, Imperial)", "34%: the real international marketers", NAVY], ["Others", "19%", LIGHT]].forEach(([a, b, col], i) => {
      const y = 4.55 + i * 0.42;
      s.addShape(pres.shapes.RECTANGLE, { x: dx + 0.1, y: y + 0.08, w: 0.2, h: 0.2, fill: { color: col }, line: { color: col, width: 0 } });
      txt(s, [{ text: a + "  ", options: { bold: true, color: NAVY } }, { text: b, options: { color: BODY } }], { x: dx + 0.45, y, w: dw - 0.45, h: 0.36, fontSize: 13 });
    });
    source(s, "Chart: TobaccoTactics citing Euromonitor (2025).", dx + 0.1, 5.85, dw - 0.1);
    s.addText([
      { text: "The paradox: ", options: { bold: true, color: NAVY } },
      { text: "tobacco is legal, but most countries have banned the “P” for Promotion. Western sales shrink, so growth moves to Asia, Africa and the Middle East.", options: { color: BODY } },
    ], { isTextBox: true, x: L, y: 6.1, w: W, h: 0.75, fontSize: 14, margin: 12, valign: "middle", fill: { color: TINT } });
    s.addNotes("(~1 min) Why is tobacco such an interesting case? 1.2 billion people use it, and 80% of them live in low- and middle-income countries, so this industry is international by necessity. China National Tobacco makes almost half of all cigarettes but sells them at home. So when we say 'international tobacco marketing', we really mean four companies: Philip Morris International, BAT, Japan Tobacco and Imperial. And they face a paradox: the product is legal, but in most countries they are not allowed to promote it.");
  }

  // ===== 5. Rules of the game =====
  {
    const s = addContent("Political & legal environment", "The rules of the game");
    const ms = [
      ["2003", "WHO tobacco treaty (FCTC)", "Today 183 parties; ad bans are a core duty"],
      ["2012", "Australia: plain packs", "The world’s first"],
      ["2016", "Uruguay beats Philip Morris", "Investment court backs big warnings"],
      ["2020", "WTO backs Australia", "Plain packs are legal under trade law"],
      ["2020", "EU bans menthol", "All flavours now banned in the EU"],
      ["2026", "UK generational ban", "No sales ever to anyone born after 2008"],
    ];
    const x0 = 1.9, x1 = 11.7, ly = 2.6, step = (x1 - x0) / (ms.length - 1);
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: LIGHT, width: 3 } });
    ms.forEach(([d, t, e], i) => {
      const cx = x0 + i * step, last = i === ms.length - 1, col = last ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.17, y: ly - 0.17, w: 0.34, h: 0.34, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x: cx - 0.95, y: ly - 0.9, w: 1.9, h: 0.45, fontSize: 19, bold: true, color: col, align: "center" });
      txt(s, t, { x: cx - 0.95, y: ly + 0.38, w: 1.9, h: 0.75, fontSize: 14, bold: true, color: col, align: "center" });
      txt(s, e, { x: cx - 0.95, y: ly + 1.12, w: 1.9, h: 0.8, fontSize: 12, color: BODY, align: "center" });
    });
    const bw = (W - 0.3) / 2;
    [
      ["Outside the treaty", "USA (signed, never ratified) and Indonesia (never signed): two of the most important markets for the industry."],
      ["Lesson for marketers", "The industry tried to fight packaging rules with trade and investment law, and lost every major case."],
    ].forEach(([h, b], i) => {
      const x = L + i * (bw + 0.3);
      card(s, x, 4.75, bw, 1.75, { fill: TINT, line: TINT });
      txt(s, h, { x: x + 0.3, y: 4.92, w: bw - 0.6, h: 0.38, fontSize: 17, bold: true, color: i ? RED : NAVY });
      txt(s, b, { x: x + 0.3, y: 5.35, w: bw - 0.6, h: 1.05, fontSize: 14, color: BODY });
    });
    s.addNotes("(~1 min) The rules come mainly from one treaty: the WHO Framework Convention on Tobacco Control from 2003, with 183 parties. It requires bans on advertising, promotion and sponsorship. Then came packaging: Australia introduced plain packs in 2012. Philip Morris sued Uruguay and lost in 2016; several countries took Australia to the WTO and lost in 2020. The EU banned menthol in 2020, and this year the UK passed a generational ban. Two big exceptions: the USA and Indonesia never joined the treaty, and we'll see both again when we talk about flavours.");
  }

  // ===== 6. EU vs USA packaging =====
  {
    const s = addContent("Packaging", "The pack is the last billboard: EU vs USA");
    const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0E2841" } } });
    const rl = (t) => ({ text: t, options: { bold: true, color: "0E2841", fill: { color: "F3F4F6" } } });
    const rows = [
      [hd(""), hd("European Union"), hd("United States")],
      [rl("Health warning"), "Picture + text, 65% of front and back", "Text only, small, on the side panel"],
      [rl("Graphic warnings"), "Mandatory since 2016 (EU directive)", "FDA rule from 2020 still blocked by courts"],
      [rl("Plain packaging"), "France, Ireland, Hungary, Belgium and others", "None: the brand owns the pack"],
      [rl("Flavours"), "All banned, including menthol (2020)", "Banned in 2009, except menthol"],
      [rl("Advertising"), "Banned in press, radio, online, sponsorship", "TV and radio banned since 1971; print limited"],
    ];
    s.addTable(rows, { x: L, y: 1.45, w: W, colW: [2.3, 4.64, 4.64], rowH: 0.5, fontSize: 14, color: "404040", valign: "middle", border: { type: "solid", pt: 1, color: "D1D1D1" }, margin: 0.1 });
    txt(s, "World records", { x: L, y: 4.75, w: 4, h: 0.32, fontSize: 14, bold: true, color: GREY });
    const cw = (W - 0.6) / 3;
    [
      ["92.5%", "Türkiye and Timor-Leste: the largest warnings in the world"],
      ["140", "countries now require picture warnings on packs"],
      ["Every stick", "Canada and Australia print warnings on each cigarette"],
    ].forEach(([n, l], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 5.1, cw, 1.4);
      txt(s, n, { x: x + 0.3, y: 5.2, w: cw - 0.6, h: 0.6, fontSize: 28, bold: true, color: RED });
      txt(s, l, { x: x + 0.3, y: 5.8, w: cw - 0.6, h: 0.6, fontSize: 13, color: BODY });
    });
    source(s, "Sources: EU Directive 2014/40/EU; US FDA; Canadian Cancer Society, Cigarette Package Health Warnings (2025).");
    s.addNotes("(~1 min) When advertising is banned, the pack becomes the last billboard, and that is why regulators fight over it. In the EU, 65% of the front and back is a picture-and-text warning, and many countries, including Hungary, add plain packaging. In the US it is the opposite: only a small text warning on the side, and the FDA's graphic warnings from 2020 are still stuck in court. Flavours too: the EU banned everything including menthol; the US banned flavours in 2009 but kept menthol. Elsewhere it goes even further: in Türkiye 92.5% of the pack is warning, and Canada and Australia now print warnings on every single cigarette.");
  }

  // ===== 7. A world of flavours =====
  {
    const s = addContent("Product adaptation", "A world of flavours", "CONTENT_PICTURE");
    previewSlots(s, [SIDE]);
    const items = [
      ["Indonesia", "Clove kretek cigarettes are national culture, so PMI bought local Sampoerna (2005) instead of pushing Marlboro."],
      ["Japan & Philippines", "Menthol and “ice” variants are mainstream products, not a niche."],
      ["Latin America & Asia", "Capsule cigarettes: you click a ball in the filter and the flavour changes."],
      ["Middle East", "Flavoured shisha (waterpipe) tobacco: apple, mint, grape."],
      ["USA", "Menthol is used by nearly 85% of Black smokers, after decades of targeted marketing (FDA)."],
      ["European Union", "No flavours at all since 2020: one taste fits all."],
    ];
    const cw = 3.85, ch = 1.6, gap = 0.3;
    items.forEach(([h, b], i) => {
      const x = L + (i % 2) * (cw + gap), y = 1.5 + Math.floor(i / 2) * (ch + 0.2);
      card(s, x, y, cw, ch, { fill: i === 5 ? TINT : WHITE, line: i === 5 ? TINT : LIGHT });
      txt(s, h, { x: x + 0.25, y: y + 0.15, w: cw - 0.5, h: 0.35, fontSize: 16, bold: true, color: i === 5 ? NAVY : RED });
      txt(s, b, { x: x + 0.25, y: y + 0.55, w: cw - 0.5, h: 0.95, fontSize: 13.5, color: BODY });
    });
    s.addNotes("(~1 min) Taste is local, so the product itself is adapted. In Indonesia clove kretek cigarettes are part of the culture; when Philip Morris entered, it bought the local leader Sampoerna instead of pushing Marlboro. In Japan and the Philippines menthol is mainstream. In Latin America and Asia, capsule cigarettes let you click a flavour ball in the filter: product innovation as entertainment. In the Middle East it's flavoured shisha. In the US, menthol was marketed so heavily to Black communities that today about 85% of Black smokers use menthol. And in the EU: no flavours at all.\n\nPicture suggestion: a capsule ('click') pack or an Indonesian kretek pack.");
  }

  // ===== 8. Why flavours get banned =====
  {
    const s = addContent("Product adaptation", "Why flavours get banned");
    const lw = 6.1;
    txt(s, "Flavours change the first experience, and the first experience decides who starts.", { x: L, y: 1.35, w: W, h: 0.4, fontSize: 16, color: BODY });
    const reasons = [
      [I.wind, "Hide the harshness", "Menthol cools and numbs the throat, so smoke is easier to inhale."],
      [I.user, "Attract beginners", "Sweet tastes and capsules appeal to young people: perfect starter products."],
      [I.target, "Target groups", "Flavours were used to reach specific communities, e.g. menthol in Black US neighbourhoods."],
    ];
    reasons.forEach(([img, h, b], i) => {
      const y = 2.0 + i * 1.5;
      card(s, L, y, lw, 1.3, { fill: TINT, line: TINT });
      iconCircle(s, img, L + 0.25, y + 0.33, 0.62);
      txt(s, h, { x: L + 1.1, y: y + 0.17, w: lw - 1.3, h: 0.36, fontSize: 17, bold: true, color: NAVY });
      txt(s, b, { x: L + 1.1, y: y + 0.55, w: lw - 1.3, h: 0.65, fontSize: 14, color: BODY });
    });
    const tx = 7.6, tw = R - tx;
    txt(s, "Key bans and fights", { x: tx, y: 2.0, w: tw, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const ev = [
      ["2009", "USA bans all flavours, except menthol"],
      ["2012", "Indonesia wins at the WTO: banning clove but not menthol was discrimination"],
      ["2020", "EU bans menthol cigarettes"],
      ["2023", "EU extends the ban to heated tobacco"],
      ["2025", "USA withdraws its proposed menthol ban"],
    ];
    s.addShape(pres.shapes.LINE, { x: tx + 0.12, y: 2.55, w: 0, h: 3.85, line: { color: LIGHT, width: 2 } });
    ev.forEach(([d, t], i) => {
      const y = 2.45 + i * 0.83, col = i === 1 ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: tx, y: y + 0.08, w: 0.24, h: 0.24, fill: { color: col }, line: { color: WHITE, width: 1.5 } });
      txt(s, d, { x: tx + 0.45, y, w: 0.8, h: 0.4, fontSize: 17, bold: true, color: col });
      txt(s, t, { x: tx + 1.3, y: y + 0.03, w: tw - 1.3, h: 0.75, fontSize: 13, color: BODY });
    });
    s.addNotes("(~1 min) Why do governments ban flavours? Three reasons: menthol numbs the throat so the first cigarette doesn't hurt; sweet flavours attract beginners; and flavours were used to target specific groups. The most interesting fight: in 2009 the US banned clove cigarettes but kept menthol. Indonesia, the home of clove kretek, went to the WTO and won in 2012, because protecting the American flavour while banning the Indonesian one was discrimination. Health policy became trade policy. The EU banned menthol in 2020, while the US withdrew its menthol ban proposal in 2025.");
  }

  // ===== 9. 4Ps =====
  {
    const s = addContent("Marketing mix", "Marketing without advertising: the 4Ps adapted");
    const ps = [
      [I.box, "Product", "Capsules, slim packs aimed at women, kretek, heated tobacco (IQOS), nicotine pouches (ZYN)."],
      [I.tag, "Price", "Single sticks and small packs in poorer markets match daily incomes. Excise taxes shape every price tier."],
      [I.store, "Place", "Kiosks, Indonesian warungs, South African spaza shops, duty-free. Hungary: state-licensed National Tobacco Shops since 2013."],
      [I.mega, "Promotion", "Ferrari F1 “Mission Winnow” (2018–21), sports sponsorship in Indonesia, influencers who don’t disclose payment, and the pack itself."],
    ];
    const cw = (W - 0.3) / 2, ch = 2.2;
    ps.forEach(([img, h, b], i) => {
      const x = L + (i % 2) * (cw + 0.3), y = 1.5 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      iconCircle(s, img, x + 0.3, y + 0.3, 0.62, i === 3 ? RED : NAVY);
      txt(s, h, { x: x + 1.15, y: y + 0.38, w: cw - 1.45, h: 0.45, fontSize: 20, bold: true, color: i === 3 ? RED : NAVY });
      txt(s, b, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 1.1, fontSize: 15, color: BODY });
    });
    txt(s, "When mass media is closed, the product, the price and the shop do the selling.", { x: L, y: 6.45, w: W, h: 0.35, fontSize: 14, color: GREY, italic: true });
    s.addNotes("(~1 min) So how do you market without advertising? You adapt the 4Ps. Product: capsules, slim packs for women, heated tobacco, nicotine pouches. Price: in poorer markets cigarettes are sold one stick at a time, so the price fits a daily income. Place: dense informal retail like warungs and spaza shops; here in Hungary the state took control of retail in 2013 with the National Tobacco Shops, where you can't even see the products from outside. Promotion: sponsorships like Ferrari's 'Mission Winnow' and influencers on social media.");
  }

  // ===== 10. Smoke-free pivot =====
  {
    const s = addContent("The future", "The smoke-free pivot: strategy or rebranding?");
    s.addChart(pres.charts.BAR, [{ name: "Smoke-free share", labels: ["2024", "2025", "2030 target"], values: [38.7, 41.5, 67] }], {
      ...chartBase(), x: L - 0.1, y: 1.45, w: 6.4, h: 3.5, barDir: "col", barGapWidthPct: 60,
      chartColors: ["0E2841", "0E2841", "D1D1D1"], showTitle: true, title: "Philip Morris International: smoke-free share of net revenue, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisLabelFormatCode: "0", valAxisMaxVal: 80, valAxisMinVal: 0,
    });
    source(s, "PMI 2025 results; 2030 target = “more than two-thirds”.", L, 5.0, 6.3);
    const sx = 7.75, sw = R - sx;
    [["106", "markets sell PMI smoke-free products"], ["76%", "IQOS share of the heated-tobacco category"], ["Japan", "IQOS’s breakthrough market: clean, low smell, tech-friendly"]].forEach(([n, l], i) => {
      const y = 1.5 + i * 1.15;
      txt(s, n, { x: sx, y, w: sw, h: 0.6, fontSize: 30, bold: true, color: NAVY });
      txt(s, l, { x: sx, y: y + 0.6, w: sw, h: 0.4, fontSize: 14, color: BODY });
    });
    const hw = (W - 0.3) / 2;
    [["Industry view", "Fewer harmful products; PMI’s CEO says cigarettes “belong in museums”.", NAVY], ["Critics’ view", "Same nicotine, new flavours and devices, and new young users.", RED]].forEach(([h, b, col], i) => {
      const x = L + i * (hw + 0.3);
      card(s, x, 5.4, hw, 1.15, { line: col, lineW: 1.25 });
      txt(s, h, { x: x + 0.25, y: 5.5, w: hw - 0.5, h: 0.35, fontSize: 16, bold: true, color: col });
      txt(s, b, { x: x + 0.25, y: 5.88, w: hw - 0.5, h: 0.6, fontSize: 14, color: BODY });
    });
    s.addNotes("(~1 min) The industry's answer to all these rules is the so-called smoke-free pivot. Philip Morris already earns 41.5% of its revenue from heated tobacco and nicotine pouches, and wants more than two-thirds by 2030. IQOS has three quarters of the heated-tobacco category, and Japan was its breakthrough market. Is this a real strategy or just rebranding? The industry says it is moving away from cigarettes. Critics say it is the same nicotine with new flavours, new devices and new young users.");
  }

  // ===== 11. Conclusion =====
  {
    const s = addContent("Conclusion", "Conclusion & discussion");
    bullets(s, [
      ["Regulation shapes strategy", "— the macro-environment decides the marketing mix, country by country."],
      ["Global identity, local execution", "— one brand, many boxes, many flavours."],
      ["Packaging is media", "— when ads are banned, the pack, the price and the shop do the selling."],
    ], { x: L, y: 1.45, w: W, h: 1.9, fontSize: 17, paraSpaceAfter: 10 });
    txt(s, "My answer", { x: L, y: 3.45, w: 4, h: 0.32, fontSize: 14, bold: true, color: GREY });
    txt(s, "“Tobacco marketing didn’t disappear. It moved into the pack, the shop and the product.”", { x: L, y: 3.8, w: W, h: 0.95, fontSize: 24, bold: true, color: RED, italic: true });
    s.addText([
      { text: "Question for discussion", options: { bold: true, fontSize: 14, color: LIGHT, breakLine: true } },
      { text: "Should a global company follow its strictest home-market rules in every country it sells in?", options: { fontSize: 20, bold: true, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.0, w: W, h: 1.5, fill: { color: NAVY }, valign: "middle", margin: 18, paraSpaceAfter: 6 });
    s.addNotes("(~1 min) To conclude: regulation shapes strategy country by country; the brands are global but the execution is local; and when advertising is banned, the pack, the price and the shop do the selling. My answer: tobacco marketing didn't disappear, it moved. And a question for you: should a global company follow its strictest home-market rules in every country it sells in, or is it fair to use weaker rules where they exist? Thank you.");
  }

  // ===== 12. Bibliography =====
  {
    const s = addContent("Sources", "Bibliography");
    const refs = [
      "Australian Government, Department of Health (2012). Tobacco plain packaging.",
      "Canadian Cancer Society (2025). Cigarette Package Health Warnings: International Status Report, 9th ed.",
      "Directive 2014/40/EU on tobacco and related products (Tobacco Products Directive). Official Journal of the EU.",
      "GOV.UK (2026). Tobacco and Vapes Bill becomes law, 29 April 2026.",
      "ICSID (2016). Philip Morris Brands Sàrl v. Oriental Republic of Uruguay, Case No. ARB/10/7.",
      "Keegan, W. J., & Green, M. C. (2020). Global Marketing, 10th ed. Pearson.",
      "Philip Morris International (2026). 2025 Fourth-Quarter & Full-Year Results, 5 February 2026.",
      "TobaccoTactics, University of Bath (2025). Tobacco Companies.",
      "TPackSS, Johns Hopkins (2023). Indonesia: health warning label compliance.",
      "U.S. Food and Drug Administration. Menthol and other flavors in tobacco products; Cigarette health warnings.",
      "World Health Organization (2025). WHO global report on trends in prevalence of tobacco use 2000–2024 and projections 2025–2030.",
      "WHO Framework Convention on Tobacco Control (2003) and 2025 Global Progress Report.",
      "WTO (2012). United States — Measures Affecting the Production and Sale of Clove Cigarettes (DS406).",
      "WTO (2020). Australia — Tobacco Plain Packaging (DS435, DS441), Appellate Body reports.",
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
