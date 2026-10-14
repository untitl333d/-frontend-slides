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

  // ===== 2. Agenda =====
  {
    const s = addContent("Structure", "Agenda");
    const items = ["Why regulate crypto", "How MiCA works", "The case for and against", "EU vs US", "How the EU is adjusting", "Verdict & business implications"];
    const cw = (W - 0.6) / 3, ch = 2.15;
    items.forEach((t, i) => {
      const x = L + (i % 3) * (cw + 0.3), y = 1.75 + Math.floor(i / 3) * (ch + 0.3);
      card(s, x, y, cw, ch, { name: "Agenda card " + (i + 1) });
      txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.35, y: y + 0.3, w: 1.5, h: 0.7, fontSize: 36, bold: true, color: RED });
      txt(s, t, { x: x + 0.35, y: y + 1.1, w: cw - 0.7, h: 0.8, fontSize: 20, color: NAVY, bold: true });
    });
  }

  // ===== 3. Why crypto should be regulated? =====
  {
    const s = addContent("Why regulate crypto", "Why crypto should be regulated?");
    s.addChart(pres.charts.BAR, [{ name: "Stolen funds", labels: ["2022", "2023", "2024", "2025"], values: [3.7, 1.8, 2.2, 3.4] }], {
      ...chartBase(), x: L - 0.1, y: 1.5, w: 7.0, h: 4.95, barDir: "col", barGapWidthPct: 70,
      chartColors: ["0E2841"], showTitle: true, title: "Funds stolen in crypto hacks, US$ bn",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "\\$0.0\" bn\"", valAxisLabelFormatCode: "0", valAxisMaxVal: 4.5, valAxisMinVal: 0,
    });
    source(s, "Source: Chainalysis (2022–2026); 2025 covers January to early December.", L, 6.5, 7.0);
    const stats = [
      ["$154 bn", "Illicit crypto volume in 2025, a record (lower-bound estimate)", "Chainalysis, 2026", RED],
      ["$9.3 bn", "US crypto fraud losses in 2024, +66% vs 2023", "FBI IC3, 2025", RED],
      ["4% → 9%", "Euro-area adults owning crypto, 2022 → 2024", "ECB SPACE study, 2024", NAVY],
    ];
    const sx = 8.45, sw = R - sx;
    stats.forEach(([n, l, src, col], i) => {
      const y = 1.6 + i * 1.65;
      if (i > 0) s.addShape(pres.shapes.LINE, { x: sx, y: y - 0.2, w: sw, h: 0, line: { color: LIGHT, width: 1 } });
      txt(s, n, { x: sx, y, w: sw, h: 0.62, fontSize: 34, bold: true, color: col });
      txt(s, l, { x: sx, y: y + 0.66, w: sw, h: 0.5, fontSize: 14, color: INK });
      txt(s, src, { x: sx, y: y + 1.2, w: sw, h: 0.25, fontSize: 10, color: GREY, italic: true });
    });
    s.addNotes("New framing slide. Hacks are volatile but never below $1.8 bn a year; illicit volume is <1% of all crypto activity (Chainalysis), yet retail exposure is rising fast (ECB: ownership more than doubled).");
  }

  // ===== 4. The problem MiCA was built to solve =====
  {
    const s = addContent("Why regulate crypto", "The problem MiCA was built to solve");
    const cw = (W - 0.3) / 2;
    [["2022", "Terra/Luna and FTX collapses", "Billions in retail losses"], ["27", "National regimes before MiCA", "No EU-wide licence"]].forEach(([n, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.6, cw, 2.35);
      txt(s, n, { x: x + 0.4, y: 1.85, w: cw - 0.8, h: 0.8, fontSize: 44, bold: true, color: RED });
      txt(s, h, { x: x + 0.4, y: 2.75, w: cw - 0.8, h: 0.45, fontSize: 22, bold: true, color: NAVY });
      txt(s, b, { x: x + 0.4, y: 3.25, w: cw - 0.8, h: 0.4, fontSize: 18, color: INK });
    });
    txt(s, "Core dilemma", { x: L, y: 4.35, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const bw = (W - 0.9) / 2;
    s.addText("Consumer protection & stability", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.8, w: bw, h: 1.35, fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 22, align: "center", valign: "middle" });
    txt(s, "vs", { x: L + bw, y: 5.25, w: 0.9, h: 0.45, fontSize: 20, bold: true, color: GREY, align: "center" });
    s.addText("Innovation & competitiveness", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L + bw + 0.9, y: 4.8, w: bw, h: 1.35, fill: { color: RED }, color: WHITE, bold: true, fontSize: 22, align: "center", valign: "middle" });
  }

  // ===== 5. MiCA in one slide =====
  {
    const s = addContent("How MiCA works", "MiCA in one slide");
    txt(s, "Three token types", { x: L, y: 1.5, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const tw = 2.6, tg = 0.3;
    [["EMT", "E-money tokens: one currency", "USDC, EURC"], ["ART", "Asset-referenced tokens: basket of assets", ""], ["Other crypto-assets", "BTC, ETH", "White paper required"]].forEach(([h, a, b], i) => {
      const x = L + i * (tw + tg);
      card(s, x, 1.95, tw, 2.25, { line: NAVY, lineW: 1.25 });
      txt(s, h, { x: x + 0.25, y: 2.12, w: tw - 0.5, h: 0.75, fontSize: 20, bold: true, color: NAVY, valign: "top" });
      txt(s, [{ text: a, options: { breakLine: !!b } }, ...(b ? [{ text: b, options: { bold: true } }] : [])], { x: x + 0.25, y: 2.95, w: tw - 0.5, h: 1.15, fontSize: 15, color: INK, paraSpaceAfter: 6 });
      // connector to CASP bar
      s.addShape(pres.shapes.LINE, { x: x + tw / 2, y: 4.2, w: 0, h: 0.45, line: { color: NAVY, width: 1.25, endArrowType: "triangle" } });
    });
    const cbw = 3 * tw + 2 * tg;
    s.addText([{ text: "CASPs", options: { bold: true, fontSize: 22, breakLine: true } }, { text: "Crypto-asset service providers: licence + EU-wide passport", options: { fontSize: 16 } }], {
      isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.65, w: cbw, h: 1.35, fill: { color: NAVY }, color: WHITE, align: "left", valign: "middle", margin: 14,
    });
    const ox = L + cbw + 0.4, ow = R - ox;
    card(s, ox, 1.95, ow, 4.05, { fill: TINT, line: TINT });
    txt(s, "Out of scope", { x: ox + 0.3, y: 2.15, w: ow - 0.6, h: 0.5, fontSize: 22, bold: true, color: NAVY });
    s.addText([{ text: "Fully decentralised DeFi", options: { bullet: true, breakLine: true } }, { text: "Most NFTs", options: { bullet: true } }], { isTextBox: true, x: ox + 0.3, y: 2.8, w: ow - 0.6, h: 1.2, fontSize: 18, color: INK, margin: 0, valign: "top", paraSpaceAfter: 8 });
  }

  // ===== 6. Timeline =====
  {
    const s = addContent("How MiCA works", "From adoption to full enforcement");
    const ms = [["2023", "MiCA adopted"], ["June 2024", "Stablecoin rules apply"], ["Dec 2024", "CASP rules apply"], ["30 June 2025", "Hungary's shortened transition ends"], ["1 July 2026", "EU-wide transition ends — only licensed firms may operate"]];
    const x0 = 2.0, x1 = 11.6, ly = 3.55;
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: LIGHT, width: 3 } });
    ms.forEach(([d, t], i) => {
      const cx = x0 + (i * (x1 - x0)) / 4, last = i === 4, col = last ? RED : NAVY;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.17, y: ly - 0.17, w: 0.34, h: 0.34, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x: cx - 1.15, y: ly - 0.95, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: col, align: "center" });
      txt(s, t, { x: cx - 1.15, y: ly + 0.4, w: 2.3, h: 1.2, fontSize: 16, color: last ? RED : INK, bold: last, align: "center" });
    });
    s.addNotes("Hungary ended its transition a year early (30 June 2025); EU-wide the transition ended on 1 July 2026.");
  }

  // ===== 7. FOR =====
  {
    const s = addContent("The case for and against", "MiCA builds trust and attracts capital");
    tag(s, "FOR", NAVY);
    s.addText([
      { text: "One licence = access to the whole EEA", options: { bullet: true, breakLine: true } },
      { text: "Banks enter: Société Générale–FORGE issues stablecoins under MiCA", options: { bullet: true } },
    ], { isTextBox: true, x: L, y: 1.6, w: 5.6, h: 1.6, fontSize: 18, color: INK, margin: 0, valign: "top", paraSpaceAfter: 12 });
    card(s, L, 3.55, 5.6, 2.75, { fill: TINT, line: TINT });
    txt(s, "+7–9%", { x: L + 0.35, y: 3.8, w: 5, h: 0.85, fontSize: 48, bold: true, color: NAVY });
    txt(s, "VC funding per quarter (2026 study)", { x: L + 0.35, y: 4.68, w: 5.0, h: 0.38, fontSize: 16, color: INK, bold: true });
    txt(s, "Late-stage: +14–17%  ·  early-stage: no change", { x: L + 0.35, y: 5.1, w: 5.0, h: 0.38, fontSize: 15, color: INK });
    txt(s, "Santos, Martins & Soares (2026), NOVA SBE", { x: L + 0.35, y: 5.65, w: 5.0, h: 0.3, fontSize: 10, color: GREY, italic: true });
    s.addChart(pres.charts.BAR, [{ name: "EURC share", labels: ["Start", "One year later"], values: [17, 41] }], {
      ...chartBase(), x: 7.0, y: 1.5, w: R - 7.0, h: 4.95, barDir: "col", barGapWidthPct: 60,
      chartColors: ["7F7F7F", "0E2841"], showTitle: true, title: "EURC share of the euro stablecoin market, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"", valAxisLabelFormatCode: "0", valAxisMaxVal: 50, valAxisMinVal: 0,
    });
    source(s, "Second value is approximate (~41%).", 7.0, 6.5, R - 7.0);
  }

  // ===== 8. AGAINST: market =====
  {
    const s = addContent("The case for and against", "The market shrank and concentrated");
    tag(s, "AGAINST", RED);
    const stats = [["≈17%", "~210 of 1,200+ registered firms got a licence by the end of the transition"], ["up to 80%", "of firms may leave the EU (OKX Europe CEO)"], ["−70%", "EU crypto VC vs the 2022 peak (Coincub 2025, industry source)"]];
    const cw = (W - 0.6) / 3;
    stats.forEach(([n, l], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 1.95);
      txt(s, n, { x: x + 0.3, y: 1.72, w: cw - 0.6, h: 0.75, fontSize: 40, bold: true, color: RED });
      txt(s, l, { x: x + 0.3, y: 2.55, w: cw - 0.6, h: 0.8, fontSize: 14, color: INK });
    });
    s.addChart(pres.charts.BAR, [{ name: "CASPs", labels: ["Germany", "France", "Netherlands", "Cyprus", "Malta"], values: [95, 36, 28, 24, 22] }], {
      ...chartBase(), x: L - 0.1, y: 3.7, w: 7.3, h: 2.85, barDir: "bar", barGapWidthPct: 45,
      chartColors: ["0E2841"], showTitle: true, title: "Authorised CASPs by home country, 30 Sep 2026",
      dataLabelPosition: "outEnd", valAxisHidden: true, valGridLine: { style: "none" }, catAxisOrientation: "maxMin", valAxisMaxVal: 115, valAxisMinVal: 0,
    });
    source(s, "Approx. counts; ESMA interim MiCA register via public trackers (~360 CASPs in total).", L, 6.58, 7.3);
    const hx = 8.45, hw = R - hx;
    card(s, hx, 3.8, hw, 2.75, { fill: TINT, line: TINT });
    txt(s, "Where the big players licensed", { x: hx + 0.3, y: 3.98, w: hw - 0.6, h: 0.35, fontSize: 14, bold: true, color: GREY });
    [["Binance", "France"], ["Kraken & Coinbase", "Ireland"], ["OKX", "Malta"]].forEach(([a, b], i) => {
      txt(s, [{ text: a + "  →  ", options: { color: INK } }, { text: b, options: { bold: true, color: NAVY } }], { x: hx + 0.3, y: 4.5 + i * 0.6, w: hw - 0.5, h: 0.45, fontSize: 16 });
    });
  }

  // ===== 9. AGAINST: stablecoins & DeFi =====
  {
    const s = addContent("The case for and against", "Stablecoins and DeFi");
    tag(s, "AGAINST", RED);
    const pts = [["USDT", "Tether did not apply → no USDT pairs on licensed EEA exchanges from 1 July 2026"], ["ARTs", "No asset-referenced tokens licensed so far"], ["Interest ban", "May weaken euro stablecoins"], ["DeFi", "The most innovative segment sits outside MiCA: not banned, not regulated"]];
    const cw = 3.85, ch = 2.35;
    pts.forEach(([h, b], i) => {
      const x = L + (i % 2) * (cw + 0.3), y = 1.6 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      txt(s, h, { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.45, fontSize: 22, bold: true, color: RED });
      txt(s, b, { x: x + 0.3, y: y + 0.8, w: cw - 0.6, h: 1.4, fontSize: 16, color: INK });
    });
    const dx = 9.4, dw = R - dx;
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Stablecoins", labels: ["USD and other", "EUR"], values: [99.7, 0.3] }], {
      x: dx, y: 1.5, w: dw, h: 3.0, holeSize: 68, chartColors: ["D1D1D1", "C00000"], showLegend: false, showValue: false, showPercent: false,
      showTitle: false, dataBorder: { pt: 1, color: "FFFFFF" },
    });
    txt(s, "0.3%", { x: dx, y: 2.6, w: dw, h: 0.75, fontSize: 30, bold: true, color: RED, align: "center", valign: "middle" });
    txt(s, "Euro stablecoins' share of the global stablecoin market (~$0.9 bn of ~$304 bn)", { x: dx, y: 4.65, w: dw, h: 0.9, fontSize: 14, color: INK, align: "center" });
    source(s, "Stablecoin Beat, Sep–Oct 2026", dx, 5.6, dw);
  }

  // ===== 10. Fraud, AML & KYC =====
  {
    const s = addContent("The case for and against", "Fraud, AML & KYC: the price of trust", "CONTENT_PICTURE");
    const lw = 8.0;
    iconCircle(s, I.alert, L, 1.6, 0.55, RED);
    txt(s, [{ text: "Problem: ", options: { bold: true, color: RED } }, { text: "scams, hacks, money laundering, sanctions evasion" }], { x: L + 0.75, y: 1.66, w: lw - 0.75, h: 0.45, fontSize: 17 });
    txt(s, "EU response", { x: L, y: 2.45, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const ms = [[I.user, "Mandatory KYC"], [I.send, "Travel Rule"], [I.shield, "EU AML authority (AMLA)"], [I.file, "DAC8 tax reporting"]];
    const cw = (lw - 0.6) / 4;
    ms.forEach(([img, t], i) => {
      const x = L + i * (cw + 0.2);
      card(s, x, 2.85, cw, 1.75, { fill: TINT, line: TINT });
      iconCircle(s, img, x + 0.25, 3.05, 0.58);
      txt(s, t, { x: x + 0.25, y: 3.78, w: cw - 0.4, h: 0.7, fontSize: 15, bold: true, color: NAVY });
    });
    const hw = (lw - 0.3) / 2;
    [["Pro-innovation view", "A clean market attracts institutional capital", NAVY], ["Anti-innovation view", "Compliance costs favour big players", RED]].forEach(([h, b, col], i) => {
      const x = L + i * (hw + 0.3);
      card(s, x, 4.9, hw, 1.6, { line: col, lineW: 1.25 });
      txt(s, h, { x: x + 0.3, y: 5.08, w: hw - 0.6, h: 0.4, fontSize: 18, bold: true, color: col });
      txt(s, b, { x: x + 0.3, y: 5.55, w: hw - 0.6, h: 0.8, fontSize: 16, color: INK });
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
    s.addTable(rows, { x: L, y: 1.6, w: W, colW: [2.9, 4.34, 4.34], rowH: 0.62, fontSize: 17, color: "1A1A1A", valign: "middle", border: { type: "solid", pt: 1, color: "D1D1D1" }, margin: 0.12 });
    s.addText([
      { text: "Bottom line:  ", options: { bold: true } },
      { text: "Europe → compliant but smaller issuers;  US → global dollar champions" },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.2, w: W, h: 0.95, fill: { color: NAVY }, color: WHITE, fontSize: 18, valign: "middle", margin: 14 });
  }

  // ===== 12. MiCA 2 =====
  {
    const s = addContent("How the EU is adjusting", "The EU is adjusting: \u201CMiCA 2\u201D", "CONTENT_PICTURE");
    const lw = 8.0;
    // mini timeline
    const ty = 2.1;
    s.addShape(pres.shapes.LINE, { x: L + 0.17, y: ty, w: lw - 0.34, h: 0, line: { color: LIGHT, width: 3 } });
    [[L, "May 2026", "Commission opens MiCA review consultation", NAVY, "left"], [L + lw - 3.6, "30 June 2027", "Commission report due — possibly a new proposal", RED, "right"]].forEach(([x, d, t, col, al], i) => {
      const dotX = i === 0 ? L : L + lw - 0.34;
      s.addShape(pres.shapes.OVAL, { x: dotX, y: ty - 0.17, w: 0.34, h: 0.34, fill: { color: col }, line: { color: WHITE, width: 2 } });
      txt(s, d, { x, y: 1.5, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: col, align: al });
      txt(s, t, { x, y: 2.4, w: 3.6, h: 0.8, fontSize: 15, color: INK, align: al });
    });
    txt(s, "Under review", { x: L, y: 3.45, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    const chips = ["DeFi", "Staking", "Lending", "Interest ban", "Reserve rules"];
    let cx = L;
    chips.forEach((c) => {
      const w = 0.32 + c.length * 0.13;
      s.addText(c, { isTextBox: true, shape: pres.shapes.RECTANGLE, x: cx, y: 3.9, w, h: 0.5, fill: { color: TINT }, line: { color: LIGHT, width: 1 }, fontSize: 15, color: NAVY, bold: true, align: "center", valign: "middle", margin: 0 });
      cx += w + 0.15;
    });
    card(s, L, 4.9, lw, 1.45, { line: NAVY, lineW: 1.25 });
    txt(s, "Parallel proposal", { x: L + 0.3, y: 5.08, w: lw - 0.6, h: 0.4, fontSize: 18, bold: true, color: NAVY });
    txt(s, "ESMA to supervise the largest crypto firms directly", { x: L + 0.3, y: 5.55, w: lw - 0.6, h: 0.5, fontSize: 16, color: INK });
  }

  // ===== 13. Verdict =====
  {
    const s = addContent("Verdict", "\u201CMiCA doesn\u2019t kill innovation. It filters it.\u201D");
    const cw = (W - 0.3) / 2;
    [["Winners", ["Scale", "Institutions", "Late-stage capital"], NAVY], ["Losers", ["Startups", "Early-stage projects", "Offshore players"], RED]].forEach(([h, items, col], i) => {
      const x = L + i * (cw + 0.3);
      s.addText(h, { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y: 1.6, w: cw, h: 0.6, fill: { color: col }, color: WHITE, bold: true, fontSize: 20, valign: "middle", margin: 14 });
      card(s, x, 2.2, cw, 1.95);
      s.addText(items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), { isTextBox: true, x: x + 0.35, y: 2.4, w: cw - 0.7, h: 1.6, fontSize: 18, color: INK, margin: 0, valign: "top", paraSpaceAfter: 8 });
    });
    iconCircle(s, I.alert, L, 4.5, 0.55, RED);
    txt(s, [{ text: "Biggest risk ahead: ", options: { bold: true, color: RED } }, { text: "how MiCA 2 treats DeFi" }], { x: L + 0.75, y: 4.58, w: 10, h: 0.45, fontSize: 18 });
    s.addText("Europe chose trust over speed", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 5.35, w: W, h: 0.95, fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 24, valign: "middle", margin: 14 });
  }

  // ===== 14. Business implications =====
  {
    const s = addContent("Business implications", "What this means for business", "CONTENT_PICTURE");
    const rows = [[I.server, "Crypto service", "6–12 month licence process; choose your hub country"], [I.coin, "Stablecoin", "Bank or e-money licence, full reserves, no interest"], [I.net, "DeFi project", "Follow the MiCA 2 process now"], [I.globe, "EU + US", "Two separate compliance tracks"]];
    const lw = 8.0, rh = 1.1, gap = 0.18;
    rows.forEach(([img, h, b], i) => {
      const y = 1.6 + i * (rh + gap);
      card(s, L, y, lw, rh, { fill: i % 2 ? WHITE : TINT, line: i % 2 ? LIGHT : TINT });
      iconCircle(s, img, L + 0.25, y + (rh - 0.62) / 2, 0.62);
      txt(s, h, { x: L + 1.15, y: y + 0.12, w: 2.2, h: rh - 0.24, fontSize: 18, bold: true, color: NAVY, valign: "middle" });
      txt(s, b, { x: L + 3.4, y: y + 0.12, w: lw - 3.6, h: rh - 0.24, fontSize: 16, color: INK, valign: "middle" });
    });
  }

  // ===== 15. Conclusion =====
  {
    const s = addContent("Conclusion", "Conclusion & discussion", "CONTENT_PICTURE");
    const lw = 8.0;
    s.addText([
      { text: "World\u2019s first comprehensive crypto regime", options: { bullet: true, breakLine: true } },
      { text: "2026 = enforcement year", options: { bullet: true } },
    ], { isTextBox: true, x: L, y: 1.6, w: lw, h: 1.0, fontSize: 18, color: INK, margin: 0, valign: "top", paraSpaceAfter: 8 });
    txt(s, "Our answer", { x: L, y: 2.85, w: 4, h: 0.35, fontSize: 14, bold: true, color: GREY });
    txt(s, "\u201CYes — but only for those who can afford compliance\u201D", { x: L, y: 3.25, w: lw, h: 1.0, fontSize: 26, bold: true, color: RED, italic: true });
    s.addText([
      { text: "Question for discussion", options: { bold: true, fontSize: 14, color: LIGHT, breakLine: true } },
      { text: "Would you choose a MiCA-licensed exchange over a cheaper offshore one?", options: { fontSize: 20, bold: true, color: WHITE } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: L, y: 4.55, w: lw, h: 1.75, fill: { color: NAVY }, valign: "middle", margin: 18, paraSpaceAfter: 6 });
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
