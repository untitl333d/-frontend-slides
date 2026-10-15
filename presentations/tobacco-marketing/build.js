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
const QUIZ = ["A", "B", "C", "D"].map((k, i) => ({ name: "pack" + k, label: "Insert photo: pack " + k, x: L + i * (QUIZ_W + QUIZ_GAP), y: 1.45, w: QUIZ_W, h: 3.55 }));
const SPLIT = { name: "picture", label: "Insert picture", x: L, y: 1.45, w: 5.6, h: 5.2 };
const PAIR_W = (W - 0.4) / 2;
const PAIR = ["then", "now"].map((k, i) => ({ name: "pic" + k, label: "Insert picture: " + k, x: L + i * (PAIR_W + 0.4), y: 1.45, w: PAIR_W, h: 3.9 }));

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
pres.defineSlideMaster({ title: "SPLIT_PICTURE", background: { color: WHITE }, objects: [...headerObjects, picPh(SPLIT)], slideNumber });
pres.defineSlideMaster({ title: "PACK_QUIZ", background: { color: WHITE }, objects: [...headerObjects, ...QUIZ.map(picPh)], slideNumber });
pres.defineSlideMaster({ title: "THEN_NOW", background: { color: WHITE }, objects: [...headerObjects, ...PAIR.map(picPh)], slideNumber });
pres.defineSlideMaster({ title: "TITLE", background: { color: NAVY }, objects: [] });
pres.defineSlideMaster({ title: "STATEMENT", background: { color: NAVY }, objects: [{ rect: { x: 0.55, y: 0, w: 0.3, h: 1.08, fill: { color: BLUE }, line: { color: BLUE, width: 0 } } }] });

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

// Dark slide with one big line: the "pause and look at me" moments of the talk
function statement(kicker, big, small, notes) {
  const s = pres.addSlide({ masterName: "STATEMENT" });
  txt(s, kicker, { x: L, y: 1.55, w: W, h: 0.4, fontSize: 16, bold: true, color: "F2B8B8", charSpacing: 1 });
  txt(s, big, { x: L, y: 2.1, w: 10.8, h: 2.6, fontSize: 44, bold: true, color: WHITE, valign: "middle" });
  if (small) txt(s, small, { x: L, y: 4.95, w: 10.4, h: 1.0, fontSize: 20, color: LIGHT });
  s.addNotes(notes);
  return s;
}

const chartBase = () => ({
  catAxisLabelColor: "404040", valAxisLabelColor: "7F7F7F",
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
  dataLabelFontFace: "+mn-lt", dataLabelFontSize: 13, dataLabelColor: "1A1A1A", dataLabelFontBold: true,
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: "0E2841",
  valGridLine: { color: "E5E5E5", size: 0.75 }, catGridLine: { style: "none" },
  catAxisLineShow: true, catAxisLineColor: "D1D1D1", valAxisLineShow: false,
  showLegend: false, showValue: true,
});

(async () => {
  const I = {
    candy: await icon(Fi.FiDroplet), usb: await icon(Fi.FiSmartphone), phone: await icon(Fi.FiInstagram), coin: await icon(Fi.FiDollarSign),
    wind: await icon(Fi.FiWind), click: await icon(Fi.FiMousePointer), heart: await icon(Fi.FiHeart),
  };
  const iconCircle = (s, img, x, y, d = 0.62, fill = NAVY) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };
  // text column to the right of a SPLIT picture
  const TX = 7.1, TW = R - TX;

  // ===== 1. Title =====
  {
    const s = pres.addSlide({ masterName: "TITLE" });
    txt(s, "International Marketing  ·  Individual presentation", { x: 0.8, y: 1.15, w: 8.4, h: 0.4, fontSize: 14, color: LIGHT, bold: true, charSpacing: 1 });
    txt(s, "Tobacco Marketing on International Markets", { x: 0.8, y: 1.7, w: 8.2, h: 2.3, fontSize: 40, bold: true, color: WHITE });
    txt(s, "How do you sell something nobody lets you advertise?", { x: 0.8, y: 4.15, w: 8.0, h: 0.9, fontSize: 22, color: LIGHT });
    txt(s, "Corvinus University of Budapest  ·  Fall 2026", { x: 0.8, y: 6.35, w: 7.4, h: 0.4, fontSize: 14, color: LIGHT });
    s.addShape(pres.shapes.OVAL, { x: 9.95, y: 2.0, w: 2.0, h: 2.0, fill: { color: "1B3A5C" }, line: { color: WHITE, width: 1.5 }, objectName: "Photo Amir Kenenbaev" });
    txt(s, "Amir Kenenbaev", { x: 9.65, y: 4.2, w: 2.6, h: 0.4, fontSize: 15, color: WHITE, bold: true, align: "center" });
    s.addNotes("(~15 s) Hi everyone, I'm Amir. Today: tobacco marketing. Not the business side, the marketing itself: how you sell something that nobody lets you advertise. Let's start with a quick test.\n\nPhoto circle: right-click the circle > Format Shape > Fill > Picture fill to insert a portrait.");
  }

  // ===== 2. Hands up =====
  statement("Quick test", "Raise your hand if you saw a cigarette ad this year.", "Now keep it up if you saw a cigarette pack.",
    "(~40 s) Raise your hand if you have seen a cigarette ad this year. [Pause. Probably almost nobody.] Now keep your hand up, or raise it, if you have seen a cigarette pack this year. [Everyone.] That's the whole story of my presentation. The ads are gone. The marketing is not. It just moved. Let me show you where it came from.");

  // numbered "lesson" box: the conclusion each slide adds to the argument
  const lesson = (s, n, text, x, y, w, h, dark = false) => s.addText([
    { text: (n ? "LESSON " + n : "LESSON") + "   ", options: { bold: true, color: dark ? "F2B8B8" : RED, fontSize: 13, charSpacing: 1 } },
    { text, options: { bold: true, color: dark ? WHITE : NAVY, fontSize: 17 } },
  ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y, w, h, fill: { color: dark ? "1B3A5C" : TINT }, line: { color: dark ? "1B3A5C" : TINT, width: 0 }, valign: "middle", margin: 14 });

  // ===== 3. The Marlboro Man =====
  {
    const s = addContent("Flashback", "The most famous ad in history", "SPLIT_PICTURE");
    previewSlots(s, [SPLIT]);
    txt(s, "1954", { x: TX, y: 1.5, w: TW, h: 0.8, fontSize: 44, bold: true, color: RED });
    txt(s, "Marlboro was a “ladies” cigarette, sold as “Mild as May”.", { x: TX, y: 2.35, w: TW, h: 0.8, fontSize: 19, color: BODY });
    txt(s, "Then an ad agency gave it a cowboy.", { x: TX, y: 3.25, w: TW, h: 0.9, fontSize: 24, bold: true, color: NAVY });
    txt(s, "Same cigarette, new story, and a symbol the whole world understood.", { x: TX, y: 4.2, w: TW, h: 0.8, fontSize: 18, color: BODY });
    lesson(s, 1, "Sell a meaning, not a product. A strong symbol travels across cultures.", TX, 5.25, TW, 1.35);
    s.addNotes("(~50 s) Meet the Marlboro Man. In the early 1950s Marlboro was a cigarette for women, sold as 'Mild as May'. In 1954 the Leo Burnett agency didn't change the cigarette at all. They changed the story: a cowboy, freedom, wide open spaces. And this story worked everywhere: you don't need to speak English to understand a cowboy. Marlboro became one of the best-selling brands in the world, with almost the same image in every country. So lesson one: people don't buy a product, they buy a meaning. And a simple, strong symbol is what makes a brand global.\n\nPicture suggestion: a classic Marlboro Man print ad.");
  }

  // ===== 4. Joe Camel =====
  {
    const s = addContent("Flashback", "Then the cowboy got a cartoon friend", "SPLIT_PICTURE");
    previewSlots(s, [SPLIT]);
    txt(s, "Joe Camel", { x: TX, y: 1.5, w: TW, h: 0.8, fontSize: 40, bold: true, color: RED });
    txt(s, "A cool cartoon camel with sunglasses, in cigarette ads from 1988.", { x: TX, y: 2.35, w: TW, h: 0.8, fontSize: 19, color: BODY });
    s.addText([
      { text: "A 1991 study: ", options: { color: WHITE } },
      { text: "6-year-olds knew Joe Camel as well as Mickey Mouse.", options: { color: WHITE, bold: true } },
    ], { isTextBox: true, shape: pres.shapes.RECTANGLE, x: TX, y: 3.35, w: TW, h: 1.5, fill: { color: NAVY }, fontSize: 20, valign: "middle", margin: 18 });
    lesson(s, 2, "The creative shows the real target. A cartoon speaks to children, whatever the brief says.", TX, 5.25, TW, 1.35);
    s.addNotes("(~50 s) In 1988 the cowboy got a cartoon friend: Joe Camel, a cool camel with sunglasses. The company said the ads were for adult smokers. But in 1991 a study showed six-year-olds a set of logos, and they recognised Joe Camel as well as Mickey Mouse. After lawsuits and public pressure, the camel disappeared in 1997. Lesson two: if you want to know who an ad is really for, don't read the strategy document. Look at the ad. A cartoon speaks to children. Remember this camel, he comes back later.\n\nPicture suggestion: a Joe Camel ad.");
  }

  // ===== 5. The ban =====
  {
    const s = statement("Then the world said", "No more ads.", "TV, billboards, magazines, sponsorship: banned in most of the world.",
      "(~40 s) Then the world said: enough. Today most countries ban tobacco ads on TV, on billboards, in magazines, and the sponsorship of sports and festivals. So imagine you're the marketing manager. Your favourite tools are gone. Do you stop marketing? Of course not. When you ban a channel, the marketing doesn't stop. It moves to the next channel that is still open. And the one channel still in every customer's hand is: the pack.");
    txt(s, "But banning a channel doesn’t stop marketing. It moves it.", { x: L, y: 5.85, w: W, h: 0.5, fontSize: 22, bold: true, color: "F2B8B8" });
  }

  // ===== 6. Guess the country =====
  {
    const s = addContent("Your turn", "The pack became the ad. Guess the country!", "PACK_QUIZ");
    previewSlots(s, QUIZ);
    QUIZ.forEach((p, i) => {
      s.addText("ABCD"[i], { isTextBox: true, shape: pres.shapes.OVAL, x: p.x, y: 5.3, w: 0.55, h: 0.55, fill: { color: RED }, color: WHITE, bold: true, fontSize: 20, align: "center", valign: "middle", margin: 0 });
      txt(s, "Which country?", { x: p.x + 0.7, y: 5.4, w: p.w - 0.7, h: 0.36, fontSize: 16, color: GREY });
    });
    s.addNotes(
      "(~45 s) Here are four packs from four countries. Which country is each one from, and what's different? Shout it out!\n\n" +
      "PHOTOS TO INSERT (ideally the same brand, e.g. Marlboro, so only the packaging rules differ):\n" +
      "A = USA: full brand design, small text warning on the side.\n" +
      "B = Kyrgyzstan: half brand, half picture warning (Kyrgyz/Russian text).\n" +
      "C = Hungary: plain pack, brand name only, big picture warning.\n" +
      "D = Australia: plain drab brown pack, the brand name in small plain letters.\n" +
      "Click the picture icon in each box to insert a photo; it fills and crops automatically.");
  }

  // ===== 7. Reveal =====
  {
    const s = addContent("The answer", "Same cigarette, four different stories", "PACK_QUIZ");
    previewSlots(s, QUIZ);
    [
      ["A  ·  USA", "Full branding."],
      ["B  ·  Kyrgyzstan", "Half brand, half warning."],
      ["C  ·  Hungary", "No logo, just the name."],
      ["D  ·  Australia", "The “ugliest colour”."],
    ].forEach(([h, b], i) => {
      const p = QUIZ[i];
      txt(s, h, { x: p.x, y: 5.15, w: p.w, h: 0.4, fontSize: 17, bold: true, color: RED });
      txt(s, b, { x: p.x, y: 5.55, w: p.w, h: 0.4, fontSize: 15, color: BODY });
    });
    lesson(s, 3, "One brand, four packs: here the law, not culture, forces adaptation.", L, 6.05, W, 0.75);
    s.addNotes("(~60 s) Here are the answers. A is the USA: the logo, the colours, the whole front is still marketing. B is Kyrgyzstan: half brand, half picture warning. C is Hungary, the one you see in every Nemzeti Dohánybolt: no logo, no colours, just the name. And D is Australia, the first country with plain packs, in a colour smokers rated the least attractive, so people call it 'the ugliest colour in the world'. Now think about what this means for a global brand. In class we talk about standardisation versus adaptation, usually because of culture. Here the brand would love to use one global pack. It can't, because every country's law decides how much of the pack it may use. That's lesson three: for a regulated product, the legal environment, not culture, forces adaptation.");
  }

  // ===== 8. Flavour =====
  {
    const s = addContent("Plan B", "When you can’t change the ad, change the taste", "SPLIT_PICTURE");
    previewSlots(s, [SPLIT]);
    [
      [I.wind, "Menthol", "Cools the throat, so the first cigarette doesn’t hurt."],
      [I.click, "The click", "A capsule in the filter: click it and the flavour changes."],
      [I.heart, "Local taste", "In Kyrgyzstan: nasvay. In the Middle East: apple shisha."],
    ].forEach(([img, h, b], i) => {
      const y = 1.5 + i * 1.25;
      iconCircle(s, img, TX, y + 0.05, 0.65, i === 1 ? RED : NAVY);
      txt(s, h, { x: TX + 0.9, y, w: TW - 0.9, h: 0.4, fontSize: 21, bold: true, color: i === 1 ? RED : NAVY });
      txt(s, b, { x: TX + 0.9, y: y + 0.45, w: TW - 0.9, h: 0.7, fontSize: 16, color: BODY });
    });
    lesson(s, 4, "Flavour adapts the product to a culture, and it also decides who the first-time user is.", TX, 5.25, TW, 1.35);
    s.addNotes("(~55 s) Plan B: if you can't change the ad, change the product. Menthol cools the throat, so the very first cigarette doesn't hurt. Then came the click: a capsule inside the filter that you squeeze to change the flavour. It's almost a toy. And every region gets its own taste: in Kyrgyzstan people use nasvay, in the Middle East it's apple shisha. That's classic product adaptation. But notice the other side: a regular smoker doesn't need a cigarette that doesn't hurt or that clicks. A beginner does. So lesson four: flavour is not only cultural adaptation, it's also a targeting tool. It decides who your first-time customer is. And that's exactly why most countries now ban flavours.\n\nPicture suggestion: a capsule ('click') cigarette pack.");
  }

  // ===== 9. The vape chart =====
  {
    const s = addContent("Plot twist", "Then came the vape");
    s.addChart(pres.charts.BAR, [{ name: "Vapers", labels: ["2011", "2014", "2019", "2024"], values: [1.5, 13.4, 27.5, 7.8] }], {
      ...chartBase(), x: L - 0.1, y: 1.4, w: 7.4, h: 4.9, barDir: "col", barGapWidthPct: 45,
      chartColors: ["0E2841", "0E2841", "C00000", "0E2841"], showTitle: true, title: "US high-school students who vape, %",
      dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisLabelFormatCode: "0", valAxisMaxVal: 32, valAxisMinVal: 0,
    });
    source(s, "CDC / FDA National Youth Tobacco Survey. The 2019 survey moved to tablets, which may raise estimates.", L, 6.4, 7.3);
    const sx = 8.85, sw = R - sx;
    txt(s, "Teenagers had almost stopped smoking.", { x: sx, y: 1.6, w: sw, h: 0.9, fontSize: 19, color: BODY });
    txt(s, "Then 1 in 4 high school students was vaping.", { x: sx, y: 2.55, w: sw, h: 1.3, fontSize: 26, bold: true, color: RED });
    lesson(s, 5, "A new category escapes the old rules. That gap is where marketing grows fastest.", sx, 4.25, sw, 1.65);
    s.addNotes("(~55 s) And now the plot twist. By the 2010s teenagers had almost stopped smoking cigarettes. The story could have ended there. Then came the vape. In 2011 almost no American high-school student vaped. By 2019, one in four did. Why so fast? Because all the rules we just saw, the ad bans, the plain packs, the flavour bans, were written for cigarettes. A vape was a new category, so for a few years the old rules simply didn't apply. When regulators caught up and cracked down on sweet flavours, the curve went down again. Lesson five: a new category escapes the old rules, and that gap is where marketing grows fastest.");
  }

  // ===== 10. Why kids =====
  {
    const s = addContent("Plot twist", "Why vapes speak to kids: the 4Ps all point one way");
    const cw = (W - 0.9) / 4;
    [
      [I.candy, "Product", "Bubblegum, mango, cotton candy flavours"],
      [I.usb, "Package", "Colourful, looks like a USB stick, easy to hide"],
      [I.phone, "Promotion", "Influencers and TikTok, not TV"],
      [I.coin, "Price & place", "Cheap disposables at every kiosk"],
    ].forEach(([img, h, b], i) => {
      const x = L + i * (cw + 0.3);
      card(s, x, 1.55, cw, 3.0, { fill: TINT, line: TINT });
      iconCircle(s, img, x + 0.35, 1.8, 0.75, i === 0 ? RED : NAVY);
      txt(s, h, { x: x + 0.35, y: 2.75, w: cw - 0.7, h: 0.5, fontSize: 22, bold: true, color: i === 0 ? RED : NAVY });
      txt(s, b, { x: x + 0.35, y: 3.3, w: cw - 0.7, h: 1.1, fontSize: 17, color: BODY });
    });
    lesson(s, 6, "When every P points at the same customer, that customer is the real target.", L, 4.8, W, 0.85, true);
    txt(s, "“A candy-flavoured trap.”  — the head of the World Health Organization", { x: L, y: 5.9, w: W, h: 0.45, fontSize: 18, bold: true, italic: true, color: RED });
    s.addNotes("(~55 s) So how did vapes reach kids so fast? Look at the marketing mix. Product: bubblegum, mango, cotton candy. Not flavours for a 50-year-old smoker. Package: colourful, looks like a USB stick, easy to hide from parents and teachers. Promotion: not TV, which is banned, but influencers and TikTok, where the old rules were never written. Price and place: cheap disposables at every kiosk, affordable with pocket money. The companies say vapes are for adults who want to quit. But lesson six: when every P points at the same customer, that customer is your real target. The head of the WHO called it 'a candy-flavoured trap'. And JUUL, the biggest US vape brand, paid hundreds of millions of dollars over marketing to young people.");
  }

  // ===== 11. Then and now =====
  {
    const s = addContent("Look familiar?", "Same playbook, new channel", "THEN_NOW");
    previewSlots(s, PAIR);
    [["Then", "A cartoon camel in a magazine"], ["Now", "A sweet-flavoured vape on TikTok"]].forEach(([h, b], i) => {
      const p = PAIR[i];
      txt(s, h, { x: p.x, y: 5.5, w: p.w, h: 0.5, fontSize: 24, bold: true, color: i ? RED : NAVY });
      txt(s, b, { x: p.x, y: 6.05, w: p.w, h: 0.45, fontSize: 18, color: BODY });
    });
    s.addNotes("(~35 s) Remember Joe Camel? Here he is on the left: a cartoon in a magazine. On the right: a sweet-flavoured vape on TikTok. Different decade, different product, different channel. But it's the same playbook: make it fun, make it sweet, and put it where young people look. Lesson two from the beginning still holds: the creative shows you the real target.\n\nPicture suggestions: left = Joe Camel ad; right = screenshot of a vape post on social media (or colourful disposable vapes).");
  }

  // ===== 12. What it means for international marketers =====
  {
    const s = addContent("Conclusions", "What this means for an international marketer");
    const rows = [
      ["You can’t standardise a regulated product.", "Each country’s law decides how much of the 4Ps you may use, from the USA to Kyrgyzstan.", "Lessons 1, 3"],
      ["Close one P, and the others take over.", "Promotion was banned, so the pack, the flavour, the price and the shop did the selling.", "Lessons 3, 4"],
      ["Targeting shows in the mix, not in the mission.", "Cartoons, candy flavours and TikTok reveal the real customer, whatever the company says.", "Lessons 2, 4, 6"],
    ];
    rows.forEach(([h, b, ref], i) => {
      const y = 1.5 + i * 1.72;
      card(s, L, y, W, 1.5, { fill: i % 2 ? WHITE : TINT, line: i % 2 ? LIGHT : TINT });
      txt(s, String(i + 1), { x: L + 0.35, y: y + 0.2, w: 0.8, h: 1.1, fontSize: 48, bold: true, color: RED, valign: "middle" });
      txt(s, h, { x: L + 1.35, y: y + 0.2, w: W - 3.6, h: 0.5, fontSize: 22, bold: true, color: NAVY });
      txt(s, b, { x: L + 1.35, y: y + 0.75, w: W - 3.6, h: 0.6, fontSize: 17, color: BODY });
      txt(s, ref, { x: R - 2.1, y: y + 0.2, w: 1.8, h: 1.1, fontSize: 14, color: GREY, italic: true, align: "right", valign: "middle" });
    });
    s.addNotes("(~60 s) So what do these lessons mean for us as international marketers? Three conclusions. First: you cannot standardise the marketing of a regulated product. The Marlboro Man could be the same in every country; the pack cannot. The law decides, country by country, how much of the marketing mix you are allowed to use. Second: when you close one P, the other three take over. Promotion was banned, and the pack, the flavour, the price and the shop started doing the selling. Third, and for me the most important: targeting shows in the marketing mix, not in the mission statement. A company can say 'we only market to adults', but cartoons, candy flavours and TikTok tell you who the real customer is.");
  }

  // ===== 13. My view =====
  {
    const s = addContent("My view", "The tools are neutral. The target is not.", "SPLIT_PICTURE");
    previewSlots(s, [SPLIT]);
    txt(s, "The same tools can work the other way.", { x: TX, y: 1.5, w: TW, h: 0.9, fontSize: 24, bold: true, color: NAVY });
    txt(s, "In the US, the “truth” campaign used edgy ads and youth culture against smoking.", { x: TX, y: 2.5, w: TW, h: 1.0, fontSize: 19, color: BODY });
    txt(s, "Studies credit it with part of the drop in teen smoking.", { x: TX, y: 3.6, w: TW, h: 0.8, fontSize: 19, color: BODY });
    s.addText("Marketing isn’t the problem. Who you aim it at is.", { isTextBox: true, shape: pres.shapes.RECTANGLE, x: TX, y: 4.75, w: TW, h: 1.6, fill: { color: RED }, color: WHITE, bold: true, fontSize: 24, valign: "middle", margin: 18 });
    s.addNotes("(~45 s) So is marketing the villain here? I don't think so. The same tools work the other way too. In the US, the 'truth' campaign used exactly the tobacco playbook: edgy ads, youth culture, rebellion. But it told teenagers that the tobacco industry was manipulating them. Studies credit it with part of the drop in teen smoking. So my view: the tools are neutral. Storytelling, packaging, flavour, influencers: they're just tools. Marketing isn't the problem. Who you aim it at is.\n\nPicture suggestion: a 'truth' campaign ad (truthinitiative.org).");
  }

  // ===== 14. Takeaway + question =====
  {
    const s = pres.addSlide({ masterName: "STATEMENT" });
    txt(s, "To sum up", { x: L, y: 1.25, w: W, h: 0.4, fontSize: 16, bold: true, color: "F2B8B8", charSpacing: 1 });
    txt(s, "Tobacco marketing didn’t die. It moved.", { x: L, y: 1.75, w: W, h: 0.9, fontSize: 40, bold: true, color: WHITE });
    const steps = ["The ad", "The pack", "The flavour", "The phone"];
    const cw = 2.35, gap = (W - 4 * cw) / 3;
    steps.forEach((t, i) => {
      const x = L + i * (cw + gap);
      s.addText(t, { isTextBox: true, shape: pres.shapes.RECTANGLE, x, y: 3.0, w: cw, h: 0.75, fill: { color: i === 3 ? RED : "1B3A5C" }, line: { color: i === 3 ? RED : "1B3A5C", width: 0 }, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      if (i < 3) s.addShape(pres.shapes.LINE, { x: x + cw + 0.08, y: 3.375, w: gap - 0.16, h: 0, line: { color: LIGHT, width: 2, endArrowType: "triangle" } });
    });
    txt(s, "Question for you", { x: L, y: 4.45, w: W, h: 0.4, fontSize: 16, bold: true, color: LIGHT });
    txt(s, "Where is the line between marketing a legal product and marketing to children?", { x: L, y: 4.9, w: 10.8, h: 1.3, fontSize: 30, bold: true, color: WHITE });
    s.addNotes("(~35 s) To sum up: tobacco marketing didn't die. It moved: from the ad, to the pack, to the flavour, and now to the phone. Every time one door closed, marketers found the next one, and the next door will probably be a product we haven't seen yet. So I'll leave you with a question I find hard to answer: where is the line between marketing a legal product and marketing to children? Thank you!");
  }

  // ===== 15. Sources =====
  {
    const s = addContent("Sources", "Bibliography");
    const refs = [
      "Campaign for Tobacco-Free Kids (2025). Tobacco Control Laws: Kyrgyzstan, packaging and labeling.",
      "Canadian Cancer Society (2025). Cigarette Package Health Warnings: International Status Report, 9th ed.",
      "CDC (2015, 2019, 2024). Tobacco product use among middle and high school students. MMWR.",
      "Farrelly, M. C. et al. (2005). Evidence of a dose–response relationship between “truth” antismoking ads and youth smoking prevalence. American Journal of Public Health 95(3).",
      "Fischer, P. M. et al. (1991). Brand logo recognition by children aged 3 to 6 years: Mickey Mouse and Old Joe the Camel. JAMA 266(22).",
      "Keegan, W. J., & Green, M. C. (2020). Global Marketing, 10th ed. Pearson.",
      "New York Attorney General (2023). JUUL settlement with six states and DC.",
      "WHO (2024). Hooking the next generation: how the tobacco industry captures young customers.",
      "WHO Framework Convention on Tobacco Control (2003), Article 13: advertising, promotion and sponsorship.",
      "Wikipedia. Marlboro Man; Joe Camel (history of the campaigns).",
    ];
    s.addText(refs.map((r, j) => ({ text: r, options: { breakLine: j < refs.length - 1 } })), {
      isTextBox: true, x: L, y: 1.45, w: W, h: 5.3, fontSize: 14, color: INK, margin: 0, valign: "top", paraSpaceAfter: 10,
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
