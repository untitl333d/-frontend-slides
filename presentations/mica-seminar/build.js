const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fi = require("react-icons/fi");
const { applyTheme } = require("/root/.claude/skills/synced/31324299-bbde-44e2-af66-4158c184d70c_e35be97c-095b-463d-8384-ec50034c49cc/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

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
pres.title = "Can Europe Regulate Crypto Without Killing Innovation?";
pres.author = "Amir Kenenbaev, Aruuke Dzhumabaeva";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const NAVY = C.accent1, RED = C.accent2, GREY = C.accent3, LIGHT = C.accent4, BLUE = C.accent5, INK = C.text1, TINT = C.background2, WHITE = C.background1;

const TOPIC = "Can Europe Regulate Crypto Without Killing Innovation?";
const L = 1.0;            // left content edge (aligned with title text)
const R = 12.58;          // right content edge
const W = R - L;

// ---------- layouts ----------
const headerObjects = [
  { rect: { x: 0.55, y: 0, w: 0.3, h: 1.08, fill: { color: BLUE }, line: { color: BLUE, width: 0 } } },
  { placeholder: { options: { name: "topic", type: "body", x: L, y: 0.2, w: 9.6, h: 0.36, fontSize: 14, bold: true, color: GREY, valign: "middle", align: "left", margin: 0 }, text: "" } },
  { placeholder: { options: { name: "title", type: "title", x: L, y: 0.52, w: 10.0, h: 0.6, fontSize: 26, bold: true, color: NAVY, valign: "middle", align: "left", margin: 0 }, text: "" } },
  { text: { text: TOPIC, options: { x: L, y: 6.98, w: 8, h: 0.3, fontSize: 10, color: GREY, margin: 0 } } },
];
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: headerObjects,
  slideNumber: { x: 11.98, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: GREY, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT_PICTURE",
  background: { color: WHITE },
  objects: [
    ...headerObjects,
    { placeholder: { options: { name: "picture", type: "pic", x: 9.35, y: 1.55, w: 3.23, h: 5.1 }, text: "Insert picture" } },
  ],
  slideNumber: { x: 11.98, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: GREY, align: "right" },
});
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
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || WHITE },
    line: { color: opts.line || LIGHT, width: opts.lineW || 1 },
    objectName: opts.name,
  });
}

function txt(s, text, o) {
  s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontSize: 16, color: INK, ...o });
}

function tag(s, label, color) {
  s.addText(label, {
    isTextBox: true, shape: pres.shapes.RECTANGLE, x: 11.18, y: 0.32, w: 1.4, h: 0.46,
    fill: { color }, color: WHITE, bold: true, fontSize: 14, align: "center", valign: "middle", charSpacing: 2, margin: 0,
    objectName: "Tag " + label,
  });
}

function source(s, text, x = L, y = 6.62, w = W) {
  txt(s, text, { x, y, w, h: 0.28, fontSize: 10, color: GREY, italic: true });
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

(async () => {
  const I = {
    alert: await icon(Fi.FiAlertTriangle), user: await icon(Fi.FiUserCheck), send: await icon(Fi.FiSend),
    shield: await icon(Fi.FiShield), file: await icon(Fi.FiFileText), server: await icon(Fi.FiServer),
    coin: await icon(Fi.FiDatabase), net: await icon(Fi.FiShare2), globe: await icon(Fi.FiGlobe),
    alertNavy: await icon(Fi.FiAlertTriangle, "#C00000"),
  };
  const iconCircle = (s, img, x, y, d = 0.62, fill = NAVY) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };

  // ===== 1. Title =====
  {
    const s = pres.addSlide({ masterName: "TITLE" });
    txt(s, "Doing Business in Europe  ·  Seminar presentation", { x: 0.8, y: 1.15, w: 7.6, h: 0.4, fontSize: 14, color: LIGHT, bold: true, charSpacing: 1 });
    txt(s, "Can Europe Regulate Crypto Without Killing Innovation?", { x: 0.8, y: 1.7, w: 7.4, h: 2.3, fontSize: 40, bold: true, color: WHITE, valign: "top", fontFace: "Aptos" });
    txt(s, "MiCA two years on", { x: 0.8, y: 4.15, w: 7.4, h: 0.5, fontSize: 24, color: LIGHT });
    txt(s, "Corvinus University of Budapest  ·  Fall 2026", { x: 0.8, y: 6.35, w: 7.4, h: 0.4, fontSize: 14, color: LIGHT });
    const names = ["Amir Kenenbaev", "Aruuke Dzhumabaeva"];
    names.forEach((n, i) => {
      const cx = 8.75 + i * 2.3;
      s.addShape(pres.shapes.OVAL, { x: cx, y: 2.0, w: 2.0, h: 2.0, fill: { color: "1B3A5C" }, line: { color: WHITE, width: 1.5 }, objectName: "Photo " + n });
      txt(s, n, { x: cx - 0.3, y: 4.2, w: 2.6, h: 0.4, fontSize: 15, color: WHITE, bold: true, align: "center" });
    });
    s.addNotes("Speaker 1 (slides 1-7, ~10 min). Photo circles: right-click a circle > Format Shape > Fill > Picture fill to insert a portrait.");
  }
  const BODY = C.accent6; // softer dark grey for running text
  // bullet list where each item is [bold lead, plain text]
  const bullets = (s, items, o) => s.addText(
    items.flatMap(([lead, rest], i) => [
      { text: lead + " ", options: { bold: true, color: NAVY, bullet: true } },
      { text: rest, options: { color: BODY, breakLine: i < items.length - 1 } },
    ]),
    { isTextBox: true, margin: 0, valign: "top", fontSize: 16, paraSpaceAfter: 10, ...o },
  );

  // ===== 2. Agenda =====
  {
    const s = addContent("Structure", "Agenda");
    const items = [
      ["Why regulate crypto", "Risks for users and why rules are needed"],
      ["How MiCA works", "Token types, licences and the timeline"],
      ["The case for and against", "What MiCA improved and what it cost"],
      ["EU vs US", "Two different ways to regulate stablecoins"],
      ["How the EU is adjusting", "The review known as \u201CMiCA 2\u201D"],
      ["Verdict & business implications", "Our answer and what firms should do"],
    ];
    const cw = (W - 0.6) / 3, ch = 2.25;
    items.forEach(([t, d], i) => {
      const x = L + (i % 3) * (cw + 0.3), y = 1.6 + Math.floor(i / 3) * (ch + 0.3);
      card(s, x, y, cw, ch, { name: "Agenda card " + (i + 1) });
      txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.35, y: y + 0.25, w: 1.5, h: 0.65, fontSize: 32, bold: true, color: RED });
      txt(s, t, { x: x + 0.35, y: y + 0.95, w: cw - 0.7, h: 0.7, fontSize: 19, color: NAVY, bold: true });
      txt(s, d, { x: x + 0.35, y: y + 1.6, w: cw - 0.7, h: 0.5, fontSize: 14, color: BODY });
    });
  }

  // ===== 3. Why crypto should be regulated? =====
  {
    const s = addContent("Why regulate crypto", "Why crypto should be regulated?");
    txt(s, "Crypto is growing fast, but its users face three different kinds of risk.", { x: L, y: 1.35, w: W, h: 0.4, fontSize: 17, color: BODY });
    // risk 1: hacks (chart)
    txt(s, [{ text: "1  Hacks  ", options: { bold: true, color: RED } }, { text: "crypto stolen from exchanges and platforms, worldwide", options: { color: BODY } }], { x: L, y: 1.95, w: 6.9, h: 0.35, fontSize: 15 });
    s.addChart(pres.charts.BAR, [{ name: "Stolen funds", labels: ["2022", "2023", "2024", "2025"], values: [3.7, 1.8, 2.2, 3.4] }], {
      ...chartBase(), x: L - 0.1, y: 2.3, w: 7.0, h: 3.65, barDir: "col", barGapWidthPct: 70,
      chartColors: ["0E2841"], showTitle: false,
      dataLabelPosition: "outEnd", dataLabelFormatCode: "\\$0.0\" bn\"", valAxisLabelFormatCode: "0", valAxisMaxVal: 4.5, valAxisMinVal: 0,
    });
    source(s, "US$ bn. Source: Chainalysis (2022–2026); 2025 covers January to early December.", L, 5.95, 7.0);
    // risks 2 and 3 + adoption
    const sx = 8.35, sw = R - sx;
    const blocks = [
      ["2  Scams", "$9.3 bn", "lost by people in the US to crypto fraud in 2024, mostly fake investment schemes (+66% vs 2023)", "FBI IC3, 2025"],
      ["3  Crime", "$154 bn", "received by criminal and sanctioned addresses in 2025 \u2014 a record, but still under 1% of all crypto activity", "Chainalysis, 2026"],
    ];
    blocks.forEach(([k, n, l, src], i) => {
      const y = 1.95 + i * 1.95;
      txt(s, k, { x: sx, y, w: sw, h: 0.32, fontSize: 15, bold: true, color: RED });
      txt(s, n, { x: sx, y: y + 0.32, w: sw, h: 0.6, fontSize: 30, bold: true, color: NAVY });
      txt(s, l, { x: sx, y: y + 0.92, w: sw, h: 0.72, fontSize: 13, color: BODY });
      txt(s, src, { x: sx, y: y + 1.6, w: sw, h: 0.22, fontSize: 10, color: GREY, italic: true });
    });
    // why numbers differ
    s.addText([
      { text: "Why the numbers differ: ", options: { bold: true, color: NAVY } },
      { text: "hacks = theft from platforms (global); scams = money tricked out of individuals (US victims only); crime = all money flowing to illicit addresses. Meanwhile crypto ownership in the euro area rose from 4% (2022) to 9% (2024) \u2014 more people are exposed (ECB).", options: { color: BODY } },
    ], { isTextBox: true, x: L, y: 6.2, w: W, h: 0.65, fontSize: 13, margin: 0, valign: "top", fill: { color: TINT } });
    s.addNotes("$2.2 bn (2024) and $9.3 bn (2024) measure different things: hacks of crypto services worldwide vs. fraud losses reported by US victims to the FBI. Illicit volume is <1% of all crypto activity (Chainalysis).");
  }

  // ===== 4. The problem MiCA was built to solve =====
  {
    const s = addContent("Why regulate crypto", "The problem MiCA was built to solve");
    const cw = (W - 0.3) / 2;
    [
      ["2022", "Terra/Luna and FTX collapses", "Two of the biggest names in crypto failed within six months. Retail investors lost billions \u2014 and there was no common EU rulebook to protect them."],
      ["27", "National regimes before MiCA", "Every EU country had its own rules and no EU-wide licence existed. A firm needed separate approvals in each market, and user protection depended on the country."],
    ].forEach(([n, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 2.75);
      txt(s, n, { x: x + 0.4, y: 1.75, w: cw - 0.8, h: 0.75, fontSize: 40, bold: true, color: RED });
      txt(s, h, { x: x + 0.4, y: 2.55, w: cw - 0.8, h: 0.42, fontSize: 20, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.4, y: 3.05, w: cw - 0.8, h: 1.1, fontSize: 15, color: BODY });
    });
    txt(s, "Core dilemma", { x: L, y: 4.55, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const bw = (W - 0.9) / 2;
    s.addText("Consumer protection & stability", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.95, w: bw, h: 1.0, fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 20, align: "center", valign: "middle" });
    txt(s, "vs", { x: L + bw, y: 5.25, w: 0.9, h: 0.4, fontSize: 18, bold: true, color: GREY, align: "center" });
    s.addText("Innovation & competitiveness", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L + bw + 0.9, y: 4.95, w: bw, h: 1.0, fill: { color: RED }, color: WHITE, bold: true, fontSize: 20, align: "center", valign: "middle" });
    txt(s, "MiCA tries to do both: protect people without pushing crypto businesses out of Europe.", { x: L, y: 6.15, w: W, h: 0.4, fontSize: 16, color: BODY, italic: true });
  }

  // ===== 5. MiCA in one slide =====
  {
    const s = addContent("How MiCA works", "MiCA in one slide");
    txt(s, "MiCA (Markets in Crypto-Assets Regulation) is one rulebook for all 27 EU countries. It sorts tokens into three types and licenses the firms that offer them.", { x: L, y: 1.35, w: W, h: 0.65, fontSize: 16, color: BODY });
    const tw = 2.6, tg = 0.3, ty = 2.2, th = 2.4;
    [
      ["EMT", "E-money tokens: linked to one currency", "e.g. USDC, EURC"],
      ["ART", "Asset-referenced tokens: linked to a basket of assets", "e.g. several currencies or commodities"],
      ["Other crypto-assets", "Everything else", "e.g. BTC, ETH \u2014 white paper required"],
    ].forEach(([h, a, b], i) => {
      const x = L + i * (tw + tg);
      card(s, x, ty, tw, th, { line: NAVY, lineW: 1.25 });
      txt(s, h, { x: x + 0.22, y: ty + 0.15, w: tw - 0.44, h: 0.7, fontSize: 19, bold: true, color: NAVY });
      txt(s, [{ text: a, options: { breakLine: true, color: BODY } }, { text: b, options: { color: GREY, italic: true } }], { x: x + 0.22, y: ty + 0.9, w: tw - 0.44, h: 1.4, fontSize: 14, paraSpaceAfter: 6 });
      s.addShape(pres.shapes.LINE, { x: x + tw / 2, y: ty + th, w: 0, h: 0.35, line: { color: NAVY, width: 1.25, endArrowType: "triangle" } });
    });
    const cbw = 3 * tw + 2 * tg;
    s.addText([
      { text: "CASPs \u2014 crypto-asset service providers", options: { bold: true, fontSize: 19, breakLine: true } },
      { text: "Exchanges, brokers and custody wallets need a licence. One licence works in the whole EEA (the \u201Cpassport\u201D).", options: { fontSize: 15 } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: ty + th + 0.35, w: cbw, h: 1.3, fill: { color: NAVY }, color: WHITE, valign: "middle", margin: 14, paraSpaceAfter: 4 });
    const ox = L + cbw + 0.4, ow = R - ox;
    card(s, ox, ty, ow, th + 1.65, { fill: TINT, line: TINT });
    txt(s, "Out of scope", { x: ox + 0.3, y: ty + 0.15, w: ow - 0.6, h: 0.5, fontSize: 20, bold: true, color: NAVY });
    s.addText([{ text: "Fully decentralised DeFi", options: { bullet: true, breakLine: true } }, { text: "Most NFTs", options: { bullet: true } }], { isTextBox: true, x: ox + 0.3, y: ty + 0.8, w: ow - 0.6, h: 1.0, fontSize: 16, color: BODY, margin: 0, valign: "top", paraSpaceAfter: 8 });
    txt(s, "Not banned, but not covered by MiCA rules \u2014 for now.", { x: ox + 0.3, y: ty + 2.15, w: ow - 0.6, h: 0.9, fontSize: 14, color: GREY, italic: true });
  }

  // ===== 6. Timeline =====
  {
    const s = addContent("How MiCA works", "From adoption to full enforcement");
    const ms = [
      ["2023", "MiCA adopted", "The EU agrees on one rulebook"],
      ["June 2024", "Stablecoin rules apply", "EMT and ART issuers must comply"],
      ["Dec 2024", "CASP rules apply", "Exchanges and brokers can apply for licences"],
      ["30 June 2025", "Hungary's shortened transition ends", "Hungary chose a shorter period"],
      ["1 July 2026", "EU-wide transition ends", "Only licensed firms may operate"],
    ];
    const x0 = 2.0, x1 = 11.6, ly = 2.85;
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: LIGHT, width: 3 } });
    ms.forEach(([d, t, e], i) => {
      const cx = x0 + (i * (x1 - x0)) / 4, last = i === 4, col = last ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.17, y: ly - 0.17, w: 0.34, h: 0.34, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x: cx - 1.15, y: ly - 0.9, w: 2.3, h: 0.45, fontSize: 19, bold: true, color: col, align: "center" });
      txt(s, t, { x: cx - 1.15, y: ly + 0.38, w: 2.3, h: 0.75, fontSize: 15, color: last ? RED : NAVY, bold: true, align: "center" });
      txt(s, e, { x: cx - 1.1, y: ly + 1.12, w: 2.2, h: 0.75, fontSize: 13, color: BODY, align: "center" });
    });
    s.addText([
      { text: "What \u201Ctransition\u201D means: ", options: { bold: true, color: NAVY } },
      { text: "firms already working under national rules could keep operating while they applied for a MiCA licence. Since 1 July 2026 this grace period is over everywhere in the EU.", options: { color: BODY } },
    ], { isTextBox: true, x: L, y: 5.1, w: W, h: 0.95, fontSize: 15, margin: 14, valign: "middle", fill: { color: TINT } });
    s.addNotes("Hungary ended its transition a year early (30 June 2025); EU-wide the transition ended on 1 July 2026.");
  }

  // ===== 7. FOR =====
  {
    const s = addContent("The case for and against", "MiCA builds trust and attracts capital");
    tag(s, "FOR", NAVY);
    bullets(s, [
      ["One licence = access to the whole EEA.", "A firm licensed in one country can serve customers in all 30 EEA countries."],
      ["Banks enter.", "Société Générale–FORGE issues stablecoins under MiCA \u2014 clear rules make traditional finance comfortable."],
      ["Clear rules build trust.", "Users and investors know who is supervised and what rules apply."],
    ], { x: L, y: 1.45, w: 5.6, h: 2.6, fontSize: 15 });
    card(s, L, 4.2, 5.6, 2.35, { fill: TINT, line: TINT });
    txt(s, "+7–9%", { x: L + 0.35, y: 4.35, w: 5, h: 0.75, fontSize: 40, bold: true, color: NAVY });
    txt(s, "more VC funding per quarter (2026 study)", { x: L + 0.35, y: 5.1, w: 5.0, h: 0.35, fontSize: 15, color: NAVY, bold: true });
    txt(s, "Late-stage startups: +14–17%. Early-stage: no change \u2014 the extra money goes to bigger, older firms.", { x: L + 0.35, y: 5.47, w: 5.0, h: 0.62, fontSize: 13, color: BODY });
    txt(s, "Santos, Martins & Soares (2026), NOVA SBE", { x: L + 0.35, y: 6.15, w: 4.9, h: 0.25, fontSize: 10, color: GREY, italic: true });
    s.addChart(pres.charts.BAR, [{ name: "EURC share", labels: ["Start", "One year later"], values: [17, 41] }], {
      ...chartBase(), x: 7.0, y: 1.45, w: R - 7.0, h: 4.1, barDir: "col", barGapWidthPct: 60,
      chartColors: ["7F7F7F", "0E2841"], showTitle: true, title: "EURC share of the euro stablecoin market, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"", valAxisLabelFormatCode: "0", valAxisMaxVal: 50, valAxisMinVal: 0,
    });
    txt(s, "EURC, Circle's MiCA-compliant euro stablecoin, more than doubled its share in one year: compliance can be a competitive advantage.", { x: 7.15, y: 5.6, w: R - 7.15, h: 0.65, fontSize: 13, color: BODY });
    source(s, "Second value is approximate (~41%).", 7.15, 6.3, R - 7.15);
  }

  // ===== 8. AGAINST: market =====
  {
    const s = addContent("The case for and against", "The market shrank and concentrated");
    tag(s, "AGAINST", RED);
    const stats = [
      ["≈17%", "Only ~210 of 1,200+ registered firms got a licence by the end of the transition"],
      ["up to 80%", "of firms may leave the EU, according to the OKX Europe CEO"],
      ["−70%", "EU crypto VC compared with the 2022 peak (Coincub 2025, industry source)"],
    ];
    const cw = (W - 0.6) / 3;
    stats.forEach(([n, l], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.45, cw, 1.85);
      txt(s, n, { x: x + 0.3, y: 1.58, w: cw - 0.6, h: 0.7, fontSize: 36, bold: true, color: RED });
      txt(s, l, { x: x + 0.3, y: 2.35, w: cw - 0.6, h: 0.85, fontSize: 14, color: BODY });
    });
    s.addChart(pres.charts.BAR, [{ name: "CASPs", labels: ["Germany", "France", "Netherlands", "Cyprus", "Malta"], values: [95, 36, 28, 24, 22] }], {
      ...chartBase(), x: L - 0.1, y: 3.5, w: 7.3, h: 3.0, barDir: "bar", barGapWidthPct: 45,
      chartColors: ["0E2841"], showTitle: true, title: "Authorised CASPs by home country, 30 Sep 2026",
      dataLabelPosition: "outEnd", valAxisHidden: true, valGridLine: { style: "none" }, catAxisOrientation: "maxMin", valAxisMaxVal: 115, valAxisMinVal: 0,
    });
    source(s, "Approx. counts; ESMA interim MiCA register via public trackers (~360 CASPs in total).", L, 6.55, 7.3);
    const hx = 8.45, hw = R - hx;
    card(s, hx, 3.55, hw, 3.0, { fill: TINT, line: TINT });
    txt(s, "Where the big players licensed", { x: hx + 0.3, y: 3.7, w: hw - 0.6, h: 0.35, fontSize: 14, bold: true, color: GREY });
    [["Binance", "France"], ["Kraken & Coinbase", "Ireland"], ["OKX", "Malta"]].forEach(([a, b], i) => {
      txt(s, [{ text: a + "  \u2192  ", options: { color: BODY } }, { text: b, options: { bold: true, color: NAVY } }], { x: hx + 0.3, y: 4.15 + i * 0.48, w: hw - 0.5, h: 0.4, fontSize: 16 });
    });
    txt(s, "Thanks to the passport, one licence is enough \u2014 so licences cluster in a few countries.", { x: hx + 0.3, y: 5.65, w: hw - 0.6, h: 0.75, fontSize: 13, color: BODY, italic: true });
  }

  // ===== 9. AGAINST: stablecoins & DeFi =====
  {
    const s = addContent("The case for and against", "Stablecoins and DeFi");
    tag(s, "AGAINST", RED);
    const pts = [
      ["USDT", "Tether, the largest stablecoin, did not apply \u2192 no USDT pairs on licensed EEA exchanges from 1 July 2026."],
      ["ARTs", "No asset-referenced tokens licensed so far \u2014 the category is practically empty."],
      ["Interest ban", "Holders cannot earn interest on stablecoins. This may make euro stablecoins less attractive."],
      ["DeFi", "The most innovative segment sits outside MiCA: not banned, not regulated \u2014 legal uncertainty for builders."],
    ];
    const cw = 3.85, ch = 2.45;
    pts.forEach(([h, b], i) => {
      const x = L + (i % 2) * (cw + 0.3), y = 1.5 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      txt(s, h, { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.45, fontSize: 20, bold: true, color: RED });
      txt(s, b, { x: x + 0.3, y: y + 0.8, w: cw - 0.6, h: 1.5, fontSize: 15, color: BODY });
    });
    const dx = 9.4, dw = R - dx;
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Stablecoins", labels: ["USD and other", "EUR"], values: [99.7, 0.3] }], {
      x: dx, y: 1.45, w: dw, h: 3.0, holeSize: 68, chartColors: ["D1D1D1", "C00000"], showLegend: false, showValue: false, showPercent: false,
      showTitle: false, dataBorder: { pt: 1, color: "FFFFFF" },
    });
    txt(s, "0.3%", { x: dx, y: 2.55, w: dw, h: 0.75, fontSize: 30, bold: true, color: RED, align: "center", valign: "middle" });
    txt(s, "Euro stablecoins' share of the global stablecoin market (~$0.9 bn of ~$304 bn). The market is dominated by the dollar.", { x: dx, y: 4.6, w: dw, h: 1.2, fontSize: 13, color: BODY, align: "center" });
    source(s, "Stablecoin Beat, Sep–Oct 2026", dx, 5.85, dw);
  }

  // ===== 10. Fraud, AML & KYC =====
  {
    const s = addContent("The case for and against", "Fraud, AML & KYC: the price of trust", "CONTENT_PICTURE");
    const lw = 8.0;
    iconCircle(s, I.alert, L, 1.5, 0.5, RED);
    txt(s, [{ text: "Problem: ", options: { bold: true, color: RED } }, { text: "scams, hacks, money laundering, sanctions evasion", options: { color: BODY } }], { x: L + 0.7, y: 1.58, w: lw - 0.7, h: 0.4, fontSize: 16 });
    txt(s, "EU response", { x: L, y: 2.2, w: 4, h: 0.32, fontSize: 14, bold: true, color: GREY });
    const ms = [
      [I.user, "KYC", "Mandatory: platforms must check who their customers are"],
      [I.send, "Travel Rule", "Sender and receiver data travel with every transfer"],
      [I.shield, "AMLA", "New EU authority fighting money laundering"],
      [I.file, "DAC8", "Platforms report users' crypto data to tax offices"],
    ];
    const cw = (lw - 0.6) / 4;
    ms.forEach(([img, t, d], i) => {
      const x = L + i * (cw + 0.2);
      card(s, x, 2.6, cw, 2.2, { fill: TINT, line: TINT });
      iconCircle(s, img, x + 0.2, 2.75, 0.52);
      txt(s, t, { x: x + 0.2, y: 3.35, w: cw - 0.35, h: 0.35, fontSize: 15, bold: true, color: NAVY });
      txt(s, d, { x: x + 0.2, y: 3.72, w: cw - 0.35, h: 1.0, fontSize: 12.5, color: BODY });
    });
    const hw = (lw - 0.3) / 2;
    [["Pro-innovation view", "A clean market attracts institutional capital \u2014 banks and funds enter markets they trust.", NAVY], ["Anti-innovation view", "Compliance costs favour big players \u2014 startups struggle to pay for lawyers, audits and IT.", RED]].forEach(([h, b, col], i) => {
      const x = L + i * (hw + 0.3);
      card(s, x, 5.05, hw, 1.55, { line: col, lineW: 1.25 });
      txt(s, h, { x: x + 0.25, y: 5.18, w: hw - 0.5, h: 0.35, fontSize: 16, bold: true, color: col });
      txt(s, b, { x: x + 0.25, y: 5.56, w: hw - 0.5, h: 0.95, fontSize: 13.5, color: BODY });
    });
  }

  // ===== 11. EU vs US =====
  {
    const s = addContent("EU vs US", "EU vs US: contain risk or accelerate growth");
    const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0E2841" } } });
    const rl = (t) => ({ text: t, options: { bold: true, color: "0E2841", fill: { color: "F3F4F6" } } });
    const rows = [
      [hd(""), hd("EU (MiCA)"), hd("US (GENIUS Act, 2025)")],
      [rl("Scope"), "Whole crypto market", "Payment stablecoins only"],
      [rl("Philosophy"), "Contain risk", "Accelerate growth, support the dollar"],
      [rl("USDT"), "Already delisted on licensed venues", "Transition until ~2028"],
      [rl("Interest to holders"), "Banned", "Banned"],
    ];
    s.addTable(rows, { x: L, y: 1.5, w: W, colW: [2.9, 4.34, 4.34], rowH: 0.58, fontSize: 16, color: "404040", valign: "middle", border: { type: "solid", pt: 1, color: "D1D1D1" }, margin: 0.12 });
    txt(s, "Both sides ban interest for holders, but their goals differ: the EU wants to limit risk inside its market, while the US wants dollar stablecoins to grow worldwide.", { x: L, y: 4.6, w: W, h: 0.65, fontSize: 15, color: BODY });
    s.addText([
      { text: "Bottom line:  ", options: { bold: true } },
      { text: "Europe \u2192 compliant but smaller issuers;  US \u2192 global dollar champions" },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.45, w: W, h: 0.9, fill: { color: NAVY }, color: WHITE, fontSize: 18, valign: "middle", margin: 14 });
  }

  // ===== 12. MiCA 2 =====
  {
    const s = addContent("How the EU is adjusting", "The EU is adjusting: \u201CMiCA 2\u201D", "CONTENT_PICTURE");
    const lw = 8.0;
    txt(s, "Two years of practice showed gaps, so the Commission is reviewing the rules.", { x: L, y: 1.35, w: lw, h: 0.35, fontSize: 15, color: BODY });
    const ty = 2.4;
    s.addShape(pres.shapes.LINE, { x: L + 0.17, y: ty, w: lw - 0.34, h: 0, line: { color: LIGHT, width: 3 } });
    [[L, "May 2026", "Commission opens MiCA review consultation", NAVY, "left"], [L + lw - 3.6, "30 June 2027", "Commission report due \u2014 possibly a new proposal", RED, "right"]].forEach(([x, d, t, col, al], i) => {
      const dotX = i === 0 ? L : L + lw - 0.34;
      s.addShape(pres.shapes.OVAL, { x: dotX, y: ty - 0.17, w: 0.34, h: 0.34, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x, y: 1.85, w: 3.6, h: 0.35, fontSize: 17, bold: true, color: col, align: al });
      txt(s, t, { x, y: 2.68, w: 3.6, h: 0.65, fontSize: 14, color: BODY, align: al });
    });
    txt(s, "Under review", { x: L, y: 3.55, w: 4, h: 0.32, fontSize: 14, bold: true, color: GREY });
    const chips = ["DeFi", "Staking", "Lending", "Interest ban", "Reserve rules"];
    let cx = L;
    chips.forEach((c) => {
      const w = 0.32 + c.length * 0.13;
      s.addText(c, { isTextBox: true, shape: pres.shapes.RECTANGLE, x: cx, y: 3.92, w, h: 0.46, fill: { color: TINT }, line: { color: LIGHT, width: 1 }, fontSize: 14, color: NAVY, bold: true, align: "center", valign: "middle", margin: 0 });
      cx += w + 0.15;
    });
    txt(s, "These are the areas that MiCA left open or that businesses say limit growth.", { x: L, y: 4.48, w: lw, h: 0.35, fontSize: 13, color: BODY, italic: true });
    card(s, L, 5.0, lw, 1.5, { line: NAVY, lineW: 1.25 });
    txt(s, "Parallel proposal", { x: L + 0.3, y: 5.12, w: lw - 0.6, h: 0.35, fontSize: 17, bold: true, color: NAVY });
    txt(s, "ESMA to supervise the largest crypto firms directly. Today national authorities do this, so standards can differ between countries.", { x: L + 0.3, y: 5.5, w: lw - 0.6, h: 0.9, fontSize: 14, color: BODY });
  }

  // ===== 13. Verdict =====
  {
    const s = addContent("Verdict", "\u201CMiCA doesn\u2019t kill innovation. It filters it.\u201D");
    const cw = (W - 0.3) / 2;
    [
      ["Winners", [["Scale", "can afford compliance costs"], ["Institutions", "trust a regulated market"], ["Late-stage capital", "VC funding +14–17%"]], NAVY],
      ["Losers", [["Startups", "licence costs are high and fixed"], ["Early-stage projects", "no extra funding"], ["Offshore players", "cut off from EU users"]], RED],
    ].forEach(([h, items, col], i) => {
      const x = L + i * (cw + 0.3);
      s.addText(h, { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y: 1.5, w: cw, h: 0.55, fill: { color: col }, color: WHITE, bold: true, fontSize: 19, valign: "middle", margin: 14 });
      card(s, x, 2.05, cw, 1.95);
      bullets(s, items.map(([a, b]) => [a + " \u2014", b]), { x: x + 0.35, y: 2.25, w: cw - 0.7, h: 1.7, fontSize: 16, paraSpaceAfter: 8 });
    });
    iconCircle(s, I.alert, L, 4.3, 0.5, RED);
    txt(s, [{ text: "Biggest risk ahead: ", options: { bold: true, color: RED } }, { text: "how MiCA 2 treats DeFi. Too strict could push it out of Europe; too loose leaves users unprotected.", options: { color: BODY } }], { x: L + 0.7, y: 4.32, w: W - 0.7, h: 0.75, fontSize: 16 });
    s.addText("Europe chose trust over speed", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.4, w: W, h: 0.95, fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 24, valign: "middle", margin: 14 });
  }

  // ===== 14. Business implications =====
  {
    const s = addContent("Business implications", "What this means for business", "CONTENT_PICTURE");
    const rows = [
      [I.server, "Crypto service", "Plan a 6–12 month licence process and choose your hub country carefully \u2014 one licence covers the whole EEA."],
      [I.coin, "Stablecoin", "You need a bank or e-money licence, full reserves, and you may not pay interest to holders."],
      [I.net, "DeFi project", "Not covered yet, but follow the MiCA 2 process now \u2014 rules may follow the 2027 report."],
      [I.globe, "EU + US", "Run two separate compliance tracks: MiCA in Europe, the GENIUS Act for US stablecoins."],
    ];
    const lw = 8.0, rh = 1.05, gap = 0.17;
    rows.forEach(([img, h, b], i) => {
      const y = 1.5 + i * (rh + gap);
      card(s, L, y, lw, rh, { fill: i % 2 ? WHITE : TINT, line: i % 2 ? LIGHT : TINT });
      iconCircle(s, img, L + 0.22, y + (rh - 0.6) / 2, 0.6);
      txt(s, h, { x: L + 1.05, y: y + 0.1, w: 1.95, h: rh - 0.2, fontSize: 17, bold: true, color: NAVY, valign: "middle" });
      txt(s, b, { x: L + 3.05, y: y + 0.1, w: lw - 3.25, h: rh - 0.2, fontSize: 14, color: BODY, valign: "middle" });
    });
    txt(s, "Rule of thumb: compliance is now part of the business model, not an afterthought.", { x: L, y: 6.35, w: lw, h: 0.35, fontSize: 14, color: GREY, italic: true });
  }

  // ===== 15. Conclusion =====
  {
    const s = addContent("Conclusion", "Conclusion & discussion", "CONTENT_PICTURE");
    const lw = 8.0;
    bullets(s, [
      ["World\u2019s first comprehensive crypto regime", "\u2014 one rulebook for 27 countries."],
      ["2026 = enforcement year", "\u2014 only licensed firms may operate."],
      ["Trade-off", "\u2014 more trust and big players, but fewer small firms and less early-stage money."],
    ], { x: L, y: 1.45, w: lw, h: 1.7, fontSize: 16, paraSpaceAfter: 8 });
    txt(s, "Our answer", { x: L, y: 3.3, w: 4, h: 0.32, fontSize: 14, bold: true, color: GREY });
    txt(s, "\u201CYes \u2014 but only for those who can afford compliance\u201D", { x: L, y: 3.65, w: lw, h: 0.95, fontSize: 24, bold: true, color: RED, italic: true });
    s.addText([
      { text: "Question for discussion", options: { bold: true, fontSize: 14, color: LIGHT, breakLine: true } },
      { text: "Would you choose a MiCA-licensed exchange over a cheaper offshore one?", options: { fontSize: 20, bold: true, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.85, w: lw, h: 1.6, fill: { color: NAVY }, valign: "middle", margin: 18, paraSpaceAfter: 6 });
  }

  // ===== 16. Bibliography =====
  {
    const s = addContent("Sources", "Bibliography");
    const refs = [
      "BlockEden (2026). MiCA\u2019s July 2026 Cliff: The EU Stablecoin Delisting Map.",
      "Chainalysis (2026). The 2026 Crypto Crime Report. Chainalysis Inc.",
      "Chainalysis (2025). Crypto theft in 2025: stolen-funds update, December 2025. Chainalysis Inc.",
      "Coincub (2025). Europe Crypto Report 2025.",
      "Elliptic (2026). What the end of MiCA\u2019s transitional period means for crypto businesses. https://www.elliptic.co/blog/the-end-of-micas-transitional-period",
      "ESMA (2026). Statement on the end of transitional periods under MiCA, ESMA75-113276571-1679, 17 April 2026.",
      "ESMA (2026). Interim MiCA register: authorised CASPs (snapshot 30 September 2026).",
      "European Central Bank (2024). Study on the payment attitudes of consumers in the euro area (SPACE) \u2013 2024.",
      "European Commission (2026). Targeted consultation on the review of the Markets in Crypto-Assets Regulation, 20 May 2026.",
      "FBI Internet Crime Complaint Center (2025). 2024 Internet Crime Report. Federal Bureau of Investigation.",
      "FinTelegram (2026). Stablecoins: US Acceleration Under GENIUS vs. EU Containment Under MiCA.",
      "GENIUS Act (Guiding and Establishing National Innovation for U.S. Stablecoins Act), United States, 2025.",
      "KPMG Cyprus (2026). CASPs grandfathering period, June 2026.",
      "Latham & Watkins (2026). MiCA Review: European Commission Launches Consultation on EU Cryptoasset Regulatory Framework. JD Supra.",
      "Linklaters (2026). MiCA review asks whether EU\u2019s crypto regime is fit for the future. https://techinsights.linklaters.com/post/102n0up/mica-review-asks-whether-eus-crypto-regime-is-fit-for-the-future",
      "Regulation (EU) 2023/1114 on markets in crypto-assets (MiCA). Official Journal of the EU, L 150, 9.6.2023.",
      "Santos, F. M., Martins, P., Soares, R. F. (2026). The impact of the Markets in Crypto-Assets Regulation on venture capital funding for cryptocurrency startups in Europe. NOVA SBE. http://hdl.handle.net/10362/203414",
      "Stablecoin Beat (2026). EUR stablecoins and the MiCA era. https://stablecoinbeat.com/charts/eur-stablecoins/",
      "Tech Insider (2026). USDT Is Gone From EU Exchanges.",
    ];
    const half = Math.ceil(refs.length / 2);
    const colW = (W - 0.4) / 2;
    [refs.slice(0, half), refs.slice(half)].forEach((col, i) => {
      s.addText(col.map((r, j) => ({ text: r, options: { breakLine: j < col.length - 1, indentLevel: 0 } })), {
        isTextBox: true, x: L + i * (colW + 0.4), y: 1.45, w: colW, h: 5.35, fontSize: 12, color: INK, margin: 0, valign: "top", paraSpaceAfter: 5,
      });
    });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
