// Builds the Case 33 presentation (International Trade Practices, Corvinus 2026/27).
// Run: NODE_PATH=<dir with pptxgenjs>/node_modules node build-deck.js <apply_theme.js path> [out.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[2]);
const OUT = process.argv[3] || path.join(__dirname, "Case33_Neural_Computers_Norway_Mongolia.pptx");

const THEME = {
  name: "Case 33 Teal",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A1A1A", lt1: "FFFFFF", dk2: "2E4B47", lt2: "EAF3F1",
    accent1: "2A827D", accent2: "A8DDD7", accent3: "C5D8D4",
    accent4: "6F6F6F", accent5: "C0392B", accent6: "F0C870",
    hlink: "2A827D", folHlink: "6F6F6F",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Case 33: Neural computers from Norway to Mongolia";
pres.subject = "International Trade Practices, Corvinus University";
const C = pres.SchemeColor;
const SH = pres.shapes;

const MX = 0.5, CW = 12.333, TOP = 1.4, BOTTOM = 6.72;

// ---------- layouts ----------
const SECTIONS = ["Overview", "Risks and route", "Payment", "Rights and duties", "Assessment"];

function frameObjects(active) {
  const objs = [
    { text: { text: "", options: { x: 0.42, y: 0, w: 0.56, h: 0.82, shape: SH.ROUND_2_SAME_RECTANGLE, flipV: true, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 } } } },
    { placeholder: { options: { name: "title", type: "title", x: 1.15, y: 0.1, w: 11.68, h: 0.62, fontSize: 30, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 1.15, y: 0.72, w: 11.68, h: 0.42, fontSize: 16, bold: true, color: C.accent4, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "source", type: "body", x: 8.7, y: 7.0, w: 4.13, h: 0.4, fontSize: 10, color: C.accent4, align: "right", valign: "top", margin: 0 }, text: "" } },
  ];
  if (active === "Appendix") {
    objs.push({ line: { x: 1.2, y: 6.98, w: 10.9, h: 0, line: { color: C.text2, width: 1 } } });
    objs.push({ rect: { x: 5.9, y: 6.93, w: 1.5, h: 0.07, fill: { color: C.text2 }, line: { color: C.text2, width: 0 } } });
    objs.push({ text: { text: "Appendix", options: { x: 5.9, y: 7.03, w: 1.5, h: 0.3, fontSize: 12, color: C.text1, align: "center", valign: "middle", margin: 0 } } });
    return objs;
  }
  objs.push({ line: { x: 1.2, y: 6.98, w: 7.3, h: 0, line: { color: C.text2, width: 1 } } });
  SECTIONS.forEach((s, i) => {
    const x = 1.2 + i * 1.46;
    if (s === active) objs.push({ rect: { x, y: 6.93, w: 1.46, h: 0.07, fill: { color: C.text2 }, line: { color: C.text2, width: 0 } } });
    objs.push({ text: { text: s, options: { x, y: 7.03, w: 1.46, h: 0.3, fontSize: 12, bold: s === active, color: C.text1, align: "center", valign: "middle", margin: 0 } } });
  });
  return objs;
}

[...SECTIONS, "Appendix"].forEach((s) => {
  pres.defineSlideMaster({
    title: "CONTENT_" + s.toUpperCase().replace(/ /g, "_"),
    background: { color: C.background1 },
    objects: frameObjects(s),
    slideNumber: { x: 0.42, y: 0.2, w: 0.56, h: 0.45, fontSize: 18, bold: true, color: C.background1, align: "center" },
  });
});
const L = (s) => "CONTENT_" + s.toUpperCase().replace(/ /g, "_");

pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1.0, y: 1.75, w: 11.33, h: 1.9, fontSize: 48, bold: true, color: C.background1, align: "center", valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 1.0, y: 3.85, w: 11.33, h: 0.5, fontSize: 22, bold: true, color: C.accent2, align: "center", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "meta", type: "body", x: 1.0, y: 5.0, w: 11.33, h: 1.5, fontSize: 16, color: C.background1, align: "center", valign: "top", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
function content(section, title, subtitle, source, notes) {
  const s = pres.addSlide({ masterName: L(section), sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  s.addText(subtitle, { placeholder: "subtitle", bullet: false });
  if (source) s.addText(source, { placeholder: "source", bullet: false });
  if (notes) s.addNotes(notes);
  return s;
}
function pill(s, text, x, y, w, fill) {
  s.addText(text, { x, y, w, h: 0.42, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.07, fill: { color: fill || C.accent3 }, line: { color: fill || C.accent3, width: 0 }, fontSize: 16, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Header: " + text });
}
function card(s, x, y, w, h, fill, name) {
  s.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name || "Card" });
}
// items: string | [boldLabel, rest]
function runs(items, o = {}) {
  const out = [];
  items.forEach((it, i) => {
    const last = i === items.length - 1;
    const base = { bullet: o.bullet === false ? false : (o.numbered ? { type: "number", indent: 18 } : { indent: 14 }), paraSpaceAfter: o.gap ?? 6 };
    if (Array.isArray(it)) {
      out.push({ text: it[0], options: { ...base, bold: true } });
      out.push({ text: it[1], options: { breakLine: !last } });
    } else out.push({ text: it, options: { ...base, breakLine: !last } });
  });
  return out;
}
function list(s, items, x, y, w, h, o = {}) {
  s.addText(runs(items, o), { x, y, w, h, fontSize: o.size || 14, color: o.color || C.text1, valign: "top", margin: [0, 4, 0, 0], isTextBox: true, objectName: o.name || "List" });
}
function text(s, t, x, y, w, h, o = {}) {
  s.addText(t, { x, y, w, h, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, ...o });
}
function stat(s, value, label, x, y, w, o = {}) {
  text(s, value, x, y, w, 0.55, { fontSize: o.size || 28, bold: true, color: o.color || C.text2, align: o.align || "center", valign: "middle", objectName: "Stat " + value });
  text(s, label, x, y + 0.55, w, o.lh || 0.6, { fontSize: 14, align: o.align || "center", objectName: "Stat label " + value });
}
function badge(s, n, x, y, fill, color) {
  s.addText(String(n), { x, y, w: 0.4, h: 0.4, shape: SH.OVAL, fill: { color: fill || C.text2 }, line: { color: fill || C.text2, width: 0 }, fontSize: 14, bold: true, color: color || C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Step " + n });
}
const BORDER = { type: "solid", pt: 0.75, color: C.accent3 };
const hdr = (t, o = {}) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, valign: "middle", ...o } });
const cell = (t, o = {}) => ({ text: t, options: { valign: "middle", ...o } });
function table(s, rows, x, y, colW, o = {}) {
  s.addTable(rows, { x, y, w: colW.reduce((a, b) => a + b, 0), colW, fontSize: o.size || 14, color: C.text1, border: BORDER, margin: o.margin || [0.04, 0.08, 0.04, 0.08], rowH: o.rowH, valign: "middle", objectName: o.name || "Table" });
}
const SELLER = { fill: { color: C.text2 }, color: C.background1, bold: true, align: "center" };
const BUYER = { fill: { color: C.accent2 }, color: C.text1, bold: true, align: "center" };
const OTHER = { fill: { color: C.accent3 }, color: C.text1, bold: true, align: "center" };

// ================= 1. Title =================
pres.addSection({ title: "Title" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Title" });
  s.addText("Neural Computers\nfrom Norway to Mongolia", { placeholder: "title" });
  s.addText("Case study No. 33  |  CPT Incoterms® 2020  |  Promissory note with clean collection", { placeholder: "subtitle", bullet: false });
  s.addText([
    { text: "Student A (team leader)  ·  Student B", options: { bold: true, breakLine: true, bullet: false } },
    { text: "International Trade Practices  ·  Seminar group G__  ·  Autumn 2026", options: { breakLine: true, bullet: false } },
    { text: "Corvinus University  ·  Lecturer: Dr. Katalin Csekő", options: { bullet: false } },
  ], { placeholder: "meta" });
  s.addNotes("Good morning. We are presenting case 33: the sale of neural computers from Norway to Mongolia.");
}

// ================= 2. Case overview =================
pres.addSection({ title: "Overview" });
{
  const s = content("Overview", "Case Overview",
    "A licensed dual-use sale on 90 days of unsecured credit: workable, but the seller carries most of the risk", null,
    "Student A (team leader), about 20 seconds. We are two, so we split the three aspects between us. I cover the deal and the physical flow: the goods, the risks, the intermediary, the delivery term and the route. My colleague covers the money and the legal flow: the documents, the payment, both parties' rights and duties, and our critical view. Our answer in one sentence: risk passes in Norway, cost passes in Ulaanbaatar, and payment comes ninety days later with no bank behind it.");
  const rows = [
    ["Situation", "NordNeuro Systems AS (Norway) sells 12 neural computers, a dual-use item, to Tengri Analytics LLC (Mongolia) for EUR 576,000", C.text2, C.background1],
    ["Question", "How do CPT and a promissory note with clean collection divide risk, cost and duties between seller and buyer, and how does each side protect itself?", C.accent1, C.background1],
    ["Answer", "Risk passes in Norway, cost at Ulaanbaatar airport, and payment follows 90 days later with no bank undertaking. The seller needs credit cover; the buyer needs cargo insurance", C.accent2, C.text1],
  ];
  rows.forEach((r, i) => {
    const y = TOP + i * 0.95;
    text(s, r[0], MX, y, 1.6, 0.8, { fontSize: 18, bold: true, color: i === 2 ? C.accent1 : r[2], align: "right", valign: "middle", objectName: "Label " + r[0] });
    s.addText(r[1], { x: 2.3, y, w: 10.53, h: 0.8, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.12, fill: { color: r[2] }, line: { color: r[2], width: 0 }, fontSize: 16, color: r[3], align: "center", valign: "middle", margin: [14, 14, 0, 0], isTextBox: true, objectName: r[0] });
  });
  const y0 = 4.35;
  s.addText("Who presents what", { x: MX, y: y0, w: CW, h: 0.36, shape: SH.RECTANGLE, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, fontSize: 15, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Presenters header" });
  table(s, [[
    cell([{ text: "Student A (team leader): ", options: { bold: true } }, { text: "the transaction and the goods, risks, the intermediary, CPT, route and carriers" }], { align: "center" }),
    cell([{ text: "Student B: ", options: { bold: true } }, { text: "documents, the payment method, seller's and buyer's rights and duties, breaches, our assessment" }], { align: "center" }),
  ]], MX, y0 + 0.36, [CW / 2, CW / 2], { rowH: [0.72], name: "Presenters" });
  stat(s, "EUR 576,000", "contract value", 0.9, 5.7, 3.6);
  stat(s, "90 days", "credit given to the buyer", 4.87, 5.7, 3.6);
  stat(s, "4A004.b", "dual-use control entry", 8.83, 5.7, 3.6);
}

// ================= 3. Transaction at a glance =================
{
  const s = content("Overview", "The Transaction at a Glance",
    "A Norwegian exporter sells directly to a Mongolian end user and waits 90 days for payment",
    "Companies, persons and prices are fictitious; banks are real (DNB, n.d.)",
    "Student A, about 35 seconds. A Norwegian technology company sells twelve neural computing systems to a Mongolian data company for 576,000 euros. The buyer is also the final user, and the use is civil research. The seller pays the carriage to Ulaanbaatar airport under CPT. The buyer does not pay cash: it signs a promissory note due in ninety days, and the banks only collect it. Keep that in mind, because it means the goods arrive almost three months before the money.");
  const cols = [
    { x: MX, w: 4.6, head: "Seller (Exporter)", lines: [
      { text: "NordNeuro Systems AS", options: { bold: true, fontSize: 17, breakLine: true } },
      { text: "Instituttveien 18, 2007 Kjeller, Norway", options: { breakLine: true } },
      { text: "Private limited company (AS)  ·  Org. no. NO 987 654 321 MVA", options: { breakLine: true } },
      { text: "Team: ", options: { bold: true } }, { text: "Ingrid Solberg, Export Manager; Lars Haugen, Export Compliance Officer", options: { breakLine: true } },
      { text: "Bank: ", options: { bold: true } }, { text: "DNB Bank ASA, Dronning Eufemias gate 30, 0191 Oslo  ·  BIC DNBANOKK" },
    ] },
    { x: 5.3, w: 2.73, head: "Commercial Agent", lines: [
      { text: "Blue Sky Trade Agency LLC", options: { bold: true, fontSize: 17, breakLine: true } },
      { text: "Ulaanbaatar, Mongolia", options: { breakLine: true } },
      { text: "Acts in the name and on behalf of the seller", options: { breakLine: true } },
      { text: "Commission: ", options: { bold: true } }, { text: "5 %, paid when the note is paid" },
    ] },
    { x: 8.23, w: 4.6, head: "Buyer (Importer) and Final User", lines: [
      { text: "Tengri Analytics LLC", options: { bold: true, fontSize: 17, breakLine: true } },
      { text: "Peace Avenue 17, Sukhbaatar District, Ulaanbaatar 14210, Mongolia", options: { breakLine: true } },
      { text: "Limited liability company (LLC)  ·  Reg. no. 6543210", options: { breakLine: true } },
      { text: "Team: ", options: { bold: true } }, { text: "Ganbold Bat-Erdene, CEO; Dorj Oyunaa, Procurement Manager", options: { breakLine: true } },
      { text: "Bank: ", options: { bold: true } }, { text: "Khan Bank JSC, Chinggis Avenue 6, Ulaanbaatar 17010  ·  BIC AGMOMNUB", options: { breakLine: true } },
      { text: "End use: ", options: { bold: true } }, { text: "civil AI research (speech recognition, mineral-exploration data); no resale" },
    ] },
  ];
  cols.forEach((c) => {
    pill(s, c.head, c.x, TOP, c.w);
    card(s, c.x, TOP + 0.52, c.w, 2.95, C.background2, "Card " + c.head);
    s.addText(c.lines.map((l) => ({ ...l, options: { paraSpaceAfter: 5, ...l.options } })), { x: c.x + 0.15, y: TOP + 0.64, w: c.w - 0.3, h: 2.75, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "Details " + c.head });
  });
  const tiles = [
    ["12 units", "NN-Core 64 neural computing systems"],
    ["EUR 576,000", "EUR 48,000 per unit"],
    ["CPT", "Chinggis Khaan International Airport (UBN), Incoterms® 2020"],
    ["Note, 90 days", "Promissory note, clean collection under URC 522"],
  ];
  tiles.forEach((t, i) => stat(s, t[0], t[1], MX + i * 3.13, 5.15, 2.93, { size: 26, color: C.accent1 }));
}

// ================= 4. The goods =================
{
  const s = content("Overview", "The Goods: Neural Computers",
    "Hardware that learns like neurons is a dual-use item, so the sale needs a Norwegian export licence",
    "(European Commission, 2025; Government of Norway, n.d.; PwC, n.d.)",
    "Student A, about 45 seconds. A neural computer is not an ordinary server. Its hardware adjusts its own connections as it learns, like neurons do. That is useful for speech recognition, but equally for sonar or target recognition, so it is a dual-use item under entry 4A004.b. Norway is outside the EU, so it uses its own control list, which mirrors the EU one, and its own customs tariff instead of TARIC. The seller cannot ship without an individual licence from DEKSA, the Norwegian export control directorate, and DEKSA wants a signed end-user statement. Mongolia itself is not sanctioned. Import is simple: a declaration and ten percent VAT.\n\nBackup: the EU definition is 'computational devices designed or modified to mimic the behaviour of a neuron or a collection of neurons'. The 2025 update of Annex I (Delegated Regulation (EU) 2025/2003) kept the entry.");
  // col 1
  pill(s, "Nature of the Good", MX, TOP, 3.9);
  text(s, [
    { text: "A computer whose hardware imitates neurons. It changes the " },
    { text: "weights and number of connections", options: { bold: true } },
    { text: " between its computing elements based on earlier data." },
  ], MX, TOP + 0.6, 3.9, 1.35, { fontSize: 15, objectName: "Definition" });
  card(s, MX, 3.45, 3.9, 1.0, C.background2, "Civil use card");
  text(s, [{ text: "Civil use", options: { bold: true, breakLine: true } }, { text: "Speech and image recognition, low-power AI" }], MX + 0.15, 3.55, 3.6, 0.8, { objectName: "Civil use" });
  card(s, MX, 4.6, 3.9, 1.0, C.background2, "Military use card");
  text(s, [{ text: "Military use", options: { bold: true, breakLine: true } }, { text: "Sonar, radar and target recognition" }], MX + 0.15, 4.7, 3.6, 0.8, { objectName: "Military use" });
  s.addText("Same hardware, two uses: a dual-use item", { x: MX, y: 5.8, w: 3.9, h: 0.7, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: [8, 8, 0, 0], isTextBox: true, objectName: "Dual-use callout" });
  // col 2
  pill(s, "Classification", 4.65, TOP, 3.3);
  stat(s, "4A004.b", "Dual-use entry \"Neural computers\": EU Annex I, mirrored in Norway's List II", 4.65, 1.95, 3.3, { color: C.accent1, lh: 0.75 });
  stat(s, "8471.50", "HS code: processing units of data-processing machines", 4.65, 3.5, 3.3, { color: C.accent1, lh: 0.75 });
  stat(s, "No TARIC", "Norway is outside the EU customs union: the Norwegian Customs Tariff applies", 4.65, 5.05, 3.3, { color: C.accent1, lh: 0.75 });
  // col 3
  pill(s, "Export and Import Rules", 8.2, TOP, 4.63);
  s.addText("Export from Norway", { x: 8.2, y: 1.97, w: 4.63, h: 0.36, shape: SH.RECTANGLE, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, fontSize: 14, bold: true, color: C.background1, align: "left", valign: "middle", margin: [8, 8, 0, 0], isTextBox: true, objectName: "Export header" });
  list(s, ["Individual export licence from DEKSA before shipment", "End-user statement signed by the buyer", "Export declaration to Norwegian Customs", "No re-export without Norwegian consent"], 8.2, 2.42, 4.63, 1.75, { name: "Export rules", gap: 4 });
  s.addText("Import into Mongolia", { x: 8.2, y: 4.3, w: 4.63, h: 0.36, shape: SH.RECTANGLE, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, fontSize: 14, bold: true, color: C.background1, align: "left", valign: "middle", margin: [8, 8, 0, 0], isTextBox: true, objectName: "Import header" });
  list(s, ["No sanctions or embargo against Mongolia", "Import declaration at Ulaanbaatar airport customs", "IT equipment listed as zero-rated (general duty 5 %)", "10 % import VAT, about EUR 58,000"], 8.2, 4.75, 4.63, 1.75, { name: "Import rules", gap: 4 });
}

// ================= 5. Risks =================
pres.addSection({ title: "Risks and route" });
{
  const s = content("Risks and route", "Transaction-Related Risks",
    "Non-payment and diversion are the critical risks; transport loss is severe but insurable",
    "(OECD, 2025; Montreal Convention, 1999)",
    "Student A, about 45 seconds. We see six groups of risk. The product risk is the licence and diversion. The country risk is geography: Mongolia's only neighbours are Russia and China, and it signed a trade agreement with the Eurasian Economic Union in 2025, so re-export to Russia is the scenario the Norwegian authority will worry about. The counterparty risk is the biggest commercial one, and my colleague comes back to it. The transport risk has one surprising number: under the Montreal Convention the airline's liability is limited by weight, here less than twenty thousand euros for a shipment worth 576,000. So insurance is not optional.\n\nBackup: S&P BB- and Moody's B1 (both October 2025). Carrier limit: 26 SDR per kg x 480 kg = about 12,480 SDR.");
  pill(s, "Probability and Impact", MX, TOP, 4.4);
  const lo = { fill: { color: C.accent2 } }, me = { fill: { color: C.accent6 } }, hi = { fill: { color: C.accent5 }, color: C.background1 };
  const n = (t, o) => cell(t, { align: "center", fontSize: 20, bold: true, ...o });
  const ax = (t) => cell(t, { align: "center", bold: true, fontSize: 13, border: { type: "none" } });
  table(s, [
    [ax("High"), n("4", me), n("1   3", hi), n("", hi)],
    [ax("Medium"), n("5", lo), n("2", me), n("", hi)],
    [ax("Low"), n("", lo), n("", lo), n("6", me)],
    [ax(""), ax("Low"), ax("Medium"), ax("High")],
  ], MX, 1.97, [0.95, 1.15, 1.15, 1.15], { rowH: [0.85, 0.85, 0.85, 0.36], name: "Risk matrix" });
  text(s, "Impact (rows) by probability (columns)", MX, 4.93, 4.4, 0.3, { fontSize: 12, color: C.accent4, align: "center", objectName: "Matrix caption" });
  card(s, MX, 5.4, 4.4, 1.2, C.background2, "Liability callout card");
  text(s, [
    { text: "Airline liability is limited by weight: " },
    { text: "under EUR 20,000", options: { bold: true, color: C.accent5 } },
    { text: " for a shipment worth " }, { text: "EUR 576,000", options: { bold: true } },
  ], MX + 0.15, 5.48, 4.1, 1.04, { fontSize: 15, valign: "middle", objectName: "Liability callout" });
  pill(s, "Six Risks and Our Mitigation", 5.2, TOP, 7.63);
  const R = (num, label, rest, mit) => [cell(String(num), { align: "center", bold: true }), cell([{ text: label + " ", options: { bold: true } }, { text: rest }]), cell(mit)];
  table(s, [
    [hdr("#", { align: "center" }), hdr("Risk"), hdr("Mitigation")],
    R(1, "Product:", "licence refused or late; goods diverted to military use", "Licence as a contract condition; end-user statement"),
    R(2, "Country:", "neighbours are Russia and China; OECD country risk 6 of 7", "Buyer screening; audit right; export credit insurance"),
    R(3, "Counterparty:", "buyer receives the goods, then does not pay the note", "Note in hand before shipment; Eksfin guarantee"),
    R(4, "Transport:", "loss or damage in the air or at the Istanbul transfer", "Cargo insurance for 110 % of the value"),
    R(5, "Legal:", "formally invalid note; licence copy missing at customs", "Note drafted with the bank; document checklist"),
    R(6, "Currency:", "buyer earns tugrik and owes euros; seller's costs in kroner", "Invoice in euros; forward cover on both sides"),
  ], 5.2, 1.97, [0.45, 3.98, 3.2], { rowH: [0.4, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7], name: "Risk table" });
}

// ================= 6. Intermediary =================
{
  const s = content("Risks and route", "The Intermediary: A Commercial Agent",
    "An agent keeps the contract and the note directly with the end user, which a licensed sale requires",
    "(ICC, Guide to Export-Import Basics; lecture materials)",
    "Student A, about 45 seconds. The syllabus lets us choose between an agent, a distributor and an export management company. We chose a commercial agent in Ulaanbaatar. The reason is the nature of the goods. With a dual-use item the Norwegian authority wants to see exactly who the end user is. An agent only introduces the parties, so the contract and the promissory note stay directly between NordNeuro and Tengri. A distributor would buy and resell, and the end user would be one step away from us. The price of this choice is that the seller keeps the credit risk, so we pay the agent only when the note is paid.");
  card(s, MX, TOP, CW, 0.82, C.background2, "Agent banner card");
  text(s, [
    { text: "Blue Sky Trade Agency LLC", options: { bold: true } },
    { text: ", Ulaanbaatar, acts in the name and on behalf of NordNeuro. It finds and prepares the deal but does not buy the goods. Commission: " },
    { text: "5 %, paid only when the note is paid", options: { bold: true } },
  ], MX + 0.2, TOP, CW - 0.4, 0.82, { fontSize: 15, valign: "middle", align: "center", objectName: "Agent banner" });
  pill(s, "Why an Agent and Not a Distributor", MX, 2.42, 7.0);
  const ch = { fill: { color: C.accent2 }, bold: true, align: "center" };
  const c2 = (t) => cell(t, { align: "center" });
  table(s, [
    [hdr(""), hdr("Commercial agent (our choice)", { align: "center", fontSize: 13, fill: { color: C.accent1 } }), hdr("Distributor", { align: "center", fontSize: 13 }), hdr("Export management company", { align: "center", fontSize: 13 })],
    [cell("Contracts with the end user", { bold: true }), cell("Exporter", ch), c2("Distributor"), c2("Exporter, via the company")],
    [cell("Bears the credit risk", { bold: true }), cell("Exporter", ch), c2("Distributor"), c2("Exporter")],
    [cell("Authority sees the end user", { bold: true }), cell("Yes, directly", ch), c2("One step removed"), c2("Yes")],
    [cell("Controls the price", { bold: true }), cell("Exporter", ch), c2("Distributor"), c2("Exporter")],
  ], MX, 2.98, [2.15, 1.6, 1.3, 1.95], { rowH: [0.6, 0.56, 0.46, 0.56, 0.46], name: "Intermediary comparison" });
  s.addText("With a dual-use item, the licensing authority must see exactly who the end user is", { x: MX, y: 5.98, w: 7.0, h: 0.64, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: [10, 10, 0, 0], isTextBox: true, objectName: "Agent callout" });
  pill(s, "Benefits", 7.8, 2.42, 5.03);
  list(s, ["Local language, contacts and market knowledge", "Contract and note come directly from the end user", "Helps with import clearance and chasing the note", "Costs less than a distributor's margin"], 7.8, 2.95, 5.03, 1.5, { name: "Benefits", gap: 4 });
  pill(s, "Risks and Our Answers", 7.8, 4.5, 5.03);
  list(s, [["Exporter keeps the credit risk: ", "commission only on cash received"], ["Agent exceeds its authority: ", "written agreement, no power to sign or collect money"], ["Corruption exposure: ", "anti-bribery clause and due diligence"]], 7.8, 5.03, 5.03, 1.65, { name: "Agent risks", gap: 4 });
}

// ================= 7. CPT =================
{
  const s = content("Risks and route", "CPT Incoterms® 2020: Carriage Paid To",
    "Risk passes in Kjeller but the seller pays to Ulaanbaatar, so the buyer bears a risk it did not arrange",
    "(ICC, 2019)",
    "Student A, about 45 seconds. CPT means Carriage Paid To. It has two critical points, and they are more than five thousand kilometres apart. Risk passes in Kjeller, the moment the seller hands the cartons to the first carrier. Cost passes at Ulaanbaatar airport, because the seller pays the freight that far. So the buyer carries the risk of a journey that the seller organised and paid for. And CPT obliges neither party to insure. That gap is the main weakness of this term, and our recommendation is that the buyer takes out cargo insurance from door to door.\n\nBackup: CPT suits any mode of transport, including air. If the parties want the seller to insure, the matching term is CIP, where Incoterms 2020 requires all-risks cover.");
  const place = (x, w, fill, label, addr, when, lc) => {
    card(s, x, TOP, w, 1.3, fill, "Place card " + label);
    text(s, [
      { text: label, options: { bold: true, fontSize: 13, color: lc, breakLine: true } },
      { text: addr, options: { bold: true, fontSize: 16, color: C.background1, breakLine: true } },
      { text: when, options: { fontSize: 14, color: C.background1 } },
    ], x + 0.2, TOP + 0.08, w - 0.4, 1.14, { valign: "middle", objectName: label });
  };
  place(MX, 6.07, C.text2, "PLACE OF DELIVERY: RISK PASSES TO THE BUYER", "NordNeuro dispatch dock, Instituttveien 18, 2007 Kjeller, Norway", "When the goods are handed to the first carrier", C.accent2);
  place(6.77, 6.06, C.accent1, "NAMED PLACE OF DESTINATION: SELLER'S COST ENDS", "Chinggis Khaan International Airport (UBN), Cargo Terminal, Töv Province, Mongolia", "On arrival, before import clearance", C.background1);
  const st = (t) => cell(t, { bold: true, align: "center", fontSize: 13, fill: { color: C.background2 } });
  const lab = (t) => cell(t, { bold: true, align: "center" });
  table(s, [
    [st(""), st("Seller's dock, Kjeller"), st("Truck to Oslo Airport"), st("Export clearance, Oslo"), st("Flight via Istanbul"), st("Import clearance, Ulaanbaatar"), st("Truck to buyer's lab")],
    [lab("Risk"), cell("Seller", SELLER), cell("Buyer", BUYER), cell("Buyer", BUYER), cell("Buyer", BUYER), cell("Buyer", BUYER), cell("Buyer", BUYER)],
    [lab("Cost"), cell("Seller", SELLER), cell("Seller", SELLER), cell("Seller", SELLER), cell("Seller", SELLER), cell("Buyer", BUYER), cell("Buyer", BUYER)],
  ], MX, 2.9, [1.03, 1.883, 1.883, 1.883, 1.883, 1.883, 1.885], { rowH: [0.5, 0.42, 0.42], name: "Risk and cost by stage" });
  pill(s, "Seller", MX, 4.5, 3.95);
  list(s, ["Export licence and export clearance", "Contracts and pays carriage to Ulaanbaatar", "Hands the goods to the carrier in Kjeller", "Sends air waybill and notice of dispatch"], MX, 5.02, 3.95, 1.65, { name: "Seller under CPT", gap: 4 });
  pill(s, "Buyer", 4.65, 4.5, 3.95);
  list(s, ["Bears the risk of the whole flight", "Import clearance, duty and VAT", "Unloading at destination, unless the carriage contract includes it", "On-carriage from airport to city, 52 km"], 4.65, 5.02, 3.95, 1.65, { name: "Buyer under CPT", gap: 4 });
  card(s, 8.8, 4.5, 4.03, 2.12, C.text2, "Insurance gap card");
  text(s, [
    { text: "The insurance gap", options: { bold: true, fontSize: 16, color: C.accent2, breakLine: true } },
    { text: "Neither party has to insure. The buyer carries the risk of a journey the seller arranged.", options: { color: C.background1, breakLine: true } },
    { text: "Our advice: the buyer insures door to door", options: { bold: true, color: C.background1 } },
  ], 9.0, 4.6, 3.63, 1.92, { valign: "middle", paraSpaceAfter: 6, objectName: "Insurance gap" });
}

// ================= 8. Route =================
{
  const s = content("Risks and route", "The Full Route and the Carriers",
    "Air freight via Istanbul takes days, not weeks, and needs no surface transit through Russia or China",
    "(Government of Norway, n.d.; Turkish Airlines, n.d.)",
    "Student A, about 45 seconds. The goods go by air. They are light, valuable and sensitive, and Mongolia has no sea port. A forwarder collects them in Kjeller, clears export at Oslo airport, and Turkish Airlines flies them through Istanbul to Ulaanbaatar. There the buyer's broker clears import and a local truck brings them to the lab. We rejected sea and rail through China because it takes about six weeks with many handlings. We rejected rail through Russia for a legal reason: Norway prohibits the transit of dual-use goods through Russia. That is the end of my part. My colleague continues with the documents and the payment.");
  pill(s, "Recommended Route by Air", MX, TOP, 8.2);
  const who = (t) => cell(t, t === "Seller" ? SELLER : BUYER);
  const leg = (i, a, b, w) => [cell(String(i), { align: "center", bold: true }), cell(a, { bold: true }), cell(b), who(w)];
  table(s, [
    [hdr("Leg", { align: "center" }), hdr("From, to"), hdr("Mode and carrier"), hdr("Arranged by", { align: "center" })],
    leg(1, "Kjeller to Oslo Airport (OSL), about 30 km", "Truck; freight forwarder DSV Air & Sea AS", "Seller"),
    leg(2, "Oslo Airport", "Export clearance by Norwegian Customs; security screening", "Seller"),
    leg(3, "Oslo to Istanbul (IST)", "Air; Turkish Airlines (Turkish Cargo)", "Seller"),
    leg(4, "Istanbul", "Transfer airside, no customs entry", "Seller"),
    leg(5, "Istanbul to Ulaanbaatar (UBN)", "Air; Turkish Airlines (Turkish Cargo)", "Seller"),
    leg(6, "Ulaanbaatar Airport", "Import clearance by the buyer's customs broker", "Buyer"),
    leg(7, "Airport to Peace Avenue 17, Ulaanbaatar, 52 km", "Truck; Tuushin LLC", "Buyer"),
  ], MX, 1.97, [0.6, 3.15, 3.2, 1.25], { rowH: [0.4, 0.58, 0.58, 0.5, 0.5, 0.5, 0.58, 0.58], name: "Route table" });
  text(s, "Transit time door to airport: about 3 to 4 days (team estimate)", MX, 6.3, 8.2, 0.32, { fontSize: 13, color: C.accent4, objectName: "Transit note" });
  pill(s, "Rejected Alternatives", 9.0, TOP, 3.83);
  card(s, 9.0, 1.97, 3.83, 0.95, C.background2, "Alternative 1 card");
  text(s, [{ text: "Sea to Tianjin, rail through China: ", options: { bold: true } }, { text: "about six weeks and many more handlings" }], 9.15, 1.97, 3.53, 0.95, { valign: "middle", objectName: "Alternative sea and rail" });
  card(s, 9.0, 3.05, 3.83, 0.95, C.background2, "Alternative 2 card");
  text(s, [{ text: "Rail through Russia: ", options: { bold: true, color: C.accent5 } }, { text: "Norway prohibits the transit of dual-use goods through Russian territory" }], 9.15, 3.05, 3.53, 0.95, { valign: "middle", objectName: "Alternative Russia" });
  pill(s, "Also Involved", 9.0, 4.2, 3.83);
  list(s, [["DEKSA: ", "export licence"], ["DNB and Khan Bank: ", "collection"], ["Eksfin: ", "seller's credit insurance"], ["Mongol Daatgal: ", "buyer's cargo insurance"], ["Blue Sky Trade Agency: ", "agent"]], 9.0, 4.72, 3.83, 1.9, { name: "Other parties", gap: 3 });
}

// ================= 9. Documents =================
pres.addSection({ title: "Payment" });
{
  const s = content("Payment", "The Documents and Where They Go",
    "Every commercial document goes straight to the buyer; only the promissory note passes through the banks",
    "(ICC, 1995; Montreal Convention, 1999)",
    "Student B, about 35 seconds. Look at where the documents go. The invoice, the packing list and the air waybill go straight to the buyer. Only one paper travels through the banks: the promissory note. And an air waybill is not a document of title like a bill of lading. The airline hands the goods to the consignee named on it. So the seller has nothing in its hand that the buyer needs in order to get the goods. This is what makes the payment method risky.");
  pill(s, "Who Issues What, and Where It Goes", MX, TOP, 8.2);
  const d = (a, b, c, o = {}) => [cell(a, { bold: true, ...o }), cell(b, o), cell(c, o)];
  const direct = { fill: { color: C.background2 } };
  const note = { fill: { color: C.accent2 }, bold: true };
  table(s, [
    [hdr("Document"), hdr("Issued by"), hdr("Goes to")],
    d("Sales contract, pro forma invoice", "Seller", "Buyer"),
    d("End-user statement", "Buyer", "Seller, then DEKSA"),
    d("Export licence", "DEKSA", "Seller; copy with the shipment"),
    d("Commercial invoice, packing list", "Seller", "Directly to the buyer", direct),
    d("Air waybill (not negotiable)", "Carrier / forwarder", "Directly to the buyer, with the goods", direct),
    d("Certificate of origin", "Chamber of commerce", "Directly to the buyer", direct),
    d("Export declaration", "Seller's forwarder", "Norwegian Customs"),
    d("Cargo insurance certificate", "Buyer's insurer", "Buyer"),
    d("Import declaration", "Buyer's broker", "Mongolian Customs"),
    d("Promissory note", "Buyer", "Seller, then DNB, then Khan Bank", note),
  ], MX, 1.97, [2.9, 2.1, 3.2], { rowH: Array(11).fill(0.415), name: "Document table" });
  card(s, 9.0, TOP, 3.83, 1.55, C.background2, "Stat card 1");
  stat(s, "1", "document passes through the banks: the promissory note", 9.15, TOP + 0.1, 3.53, { size: 36, color: C.accent1, lh: 0.75 });
  card(s, 9.0, 3.1, 3.83, 1.55, C.background2, "Stat card 0");
  stat(s, "0", "documents of title: an air waybill is not negotiable", 9.15, 3.2, 3.53, { size: 36, color: C.accent5, lh: 0.75 });
  card(s, 9.0, 4.8, 3.83, 1.73, C.text2, "Leverage card");
  text(s, [
    { text: "The airline releases the goods to the named consignee.", options: { color: C.background1, breakLine: true } },
    { text: "The seller holds no document the buyer needs", options: { bold: true, color: C.accent2 } },
  ], 9.2, 4.9, 3.43, 1.53, { fontSize: 15, valign: "middle", paraSpaceAfter: 6, objectName: "Leverage callout" });
}

// ================= 10. Payment =================
{
  const s = content("Payment", "The Payment Method: Promissory Note with Clean Collection",
    "The banks only present the buyer's note at maturity; no bank guarantees the payment",
    "(ICC, 1995; Geneva Convention, 1930)",
    "Student B, about 50 seconds. A promissory note is the buyer's own written promise to pay 576,000 euros on a fixed day. It is unconditional and independent of the sales contract. Clean collection means the seller's bank, DNB, sends only this note to Khan Bank in Ulaanbaatar, and Khan Bank presents it for payment at maturity. Under the ICC rules, URC 522, the banks are only messengers. They check nothing and guarantee nothing. So in substance this is an open-account sale with a ninety-day credit, plus a piece of paper that makes it easier to sue or to sell the debt. It is cheap and fast. It is very good for the buyer and weak for the seller.\n\nBackup: a note needs the words 'promissory note' in its text, an unconditional promise of a fixed sum, the maturity, the place of payment, the payee, the date and place of issue and the maker's signature.");
  const def = (x, w, head, body) => {
    card(s, x, TOP, w, 1.3, C.background2, "Definition card " + head);
    text(s, [{ text: head, options: { bold: true, fontSize: 16, color: C.accent1, breakLine: true } }, { text: body }], x + 0.2, TOP + 0.08, w - 0.4, 1.14, { valign: "middle", objectName: "Definition " + head });
  };
  def(MX, 6.07, "Promissory note", "The maker's written, unconditional promise to pay a fixed sum to the payee on a fixed day. Independent of the sales contract; transferable by endorsement.");
  def(6.77, 6.06, "Clean collection (URC 522)", "Banks collect a financial document that travels without commercial documents. They act as agents and give no payment undertaking.");
  const steps = [
    ["Tengri (maker)", "Issues the note for EUR 576,000, due 22 Feb 2027, payable at Khan Bank"],
    ["NordNeuro (principal)", "Ships the goods; sends the documents directly to Tengri"],
    ["NordNeuro", "Hands the note and a collection instruction to DNB"],
    ["DNB (remitting bank)", "Forwards both to Khan Bank in Ulaanbaatar"],
    ["Khan Bank (presenting bank)", "Presents the note to Tengri at maturity"],
    ["Tengri", "Pays; Khan Bank remits to DNB, which credits NordNeuro"],
  ];
  const sw = 1.9, gap = (CW - 6 * sw) / 5;
  steps.forEach((st, i) => {
    const x = MX + i * (sw + gap), y = 2.95;
    card(s, x, y + 0.2, sw, 2.15, i % 2 ? C.background2 : C.accent3, "Step card " + (i + 1));
    badge(s, i + 1, x + 0.12, y, C.text2);
    text(s, [{ text: st[0], options: { bold: true, breakLine: true } }, { text: st[1] }], x + 0.14, y + 0.6, sw - 0.28, 1.7, { fontSize: 14, paraSpaceAfter: 4, objectName: "Step text " + (i + 1) });
  });
  card(s, MX, 5.5, 7.5, 1.1, C.background2, "Dishonour card");
  text(s, [{ text: "If the note is not paid: ", options: { bold: true } }, { text: "Khan Bank advises non-payment, and protests the note only if the collection instruction says so. No bank pays in the buyer's place." }], MX + 0.2, 5.5, 7.1, 1.1, { valign: "middle", objectName: "Dishonour" });
  s.addText([{ text: "Seller's risk: ", options: { bold: false } }, { text: "HIGH", options: { bold: true } }], { x: 8.2, y: 5.5, w: 2.2, h: 1.1, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.accent5 }, line: { color: C.accent5, width: 0 }, fontSize: 16, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Seller risk" });
  s.addText([{ text: "Buyer's risk: ", options: { bold: false } }, { text: "LOW", options: { bold: true } }], { x: 10.63, y: 5.5, w: 2.2, h: 1.1, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, fontSize: 16, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Buyer risk" });
}

// ================= 11 / 12. Seller and buyer =================
pres.addSection({ title: "Rights and duties" });
function party(title, subtitle, source, notes, duties, rights, tools, foot) {
  const s = content("Rights and duties", title, subtitle, source, notes);
  const w = 3.97, xs = [MX, 4.68, 8.86];
  pill(s, "Duties", xs[0], TOP, w);
  list(s, duties, xs[0], 1.97, w, 3.2, { name: "Duties", gap: 7, size: 15 });
  pill(s, "Rights", xs[1], TOP, w);
  list(s, rights, xs[1], 1.97, w, 3.2, { name: "Rights", gap: 7, size: 15 });
  pill(s, "Risk-Mitigation Tools", xs[2], TOP, w, C.accent2);
  list(s, tools, xs[2], 1.97, w, 3.2, { name: "Tools", numbered: true, gap: 7, size: 15 });
  card(s, MX, 5.35, CW, 1.25, C.background2, "Instruction card");
  text(s, foot, MX + 0.2, 5.35, CW - 0.4, 1.25, { fontSize: 15, valign: "middle", objectName: "Instruction" });
}
party("The Seller: Rights, Duties and Protection",
  "NordNeuro must license, ship and wait, so it secures the note before the goods leave",
  "(CISG, 1980; ICC, 2019; Eksfin, n.d.)",
  "Student B, about 50 seconds. The seller's duties under CPT are to get the export licence, to book and pay the flight, to hand over conforming goods on time, and to send the air waybill and a notice of dispatch. Its main right is to be paid at maturity. Because the note is unsecured, the seller needs its own protection. First, it does not ship until it holds the signed note. Second, it buys a supplier credit guarantee from Eksfin. Third, it can sell the note to a bank. And if the buyer goes bankrupt while the goods are still flying, the shipper can order the airline not to deliver. That right ends when the buyer collects the goods.",
  ["Deliver conforming goods and the invoice (CISG 30, 35)", "Obtain the export licence and clear export", "Contract and pay carriage to Ulaanbaatar airport", "Hand the goods to the carrier in the agreed period", "Send the air waybill and notice of dispatch", "Pack and mark for air transport"],
  ["Receive a valid note before shipment", "Payment at maturity, interest if late (CISG 62, 78)", "Stop the goods in transit if the buyer clearly will not pay (CISG 71)", "Avoid the contract and claim damages (CISG 64, 74 to 76)", "Sell or discount the note"],
  ["No shipment before the signed note is in hand", "Supplier credit guarantee from Eksfin", "Discount or forfait the note", "Retention of title in the contract", "Protest and court action on the note", "Licence as condition precedent; no-re-export clause"],
  [{ text: "Instruction to the carrier if the buyer becomes insolvent in transit: ", options: { bold: true } }, { text: "\"As shipper under air waybill 235-1234 5675 we exercise our right of disposition. Do not deliver to the consignee. Hold the consignment at our disposal and await return instructions.\"", options: { italic: true } }]);
party("The Buyer: Rights, Duties and Protection",
  "Tengri owns the transport risk from Kjeller and must pay the note even if the goods are defective",
  "(CISG, 1980; ICC, 2019; Montreal Convention, 1999)",
  "Student B, about 50 seconds. The buyer's first duty is to give the note and to pay it. The second is honesty about the end use, because the licence depends on it. Then it must clear import and collect the goods. Its biggest risk is not payment but transport, since under CPT it owns the risk from Kjeller. So it insures for 110 percent of the value. There is a second, less obvious risk. A promissory note is unconditional. If the computers arrive defective, the buyer must still pay the note and then claim under the contract. That is why we advise a factory acceptance test before the note is signed.",
  ["Issue a valid note and pay it at maturity (CISG 53, 54)", "Give a truthful end-user statement", "Take over the goods at Ulaanbaatar (CISG 60)", "Clear import, pay VAT and local charges", "Bear the transport risk from Kjeller", "Examine the goods, report defects quickly (CISG 38, 39)"],
  ["Conforming goods, on time (CISG 35)", "Air waybill and notice of dispatch in time to insure", "Repair, replacement or price reduction (CISG 46, 50)", "Extra time, then avoidance (CISG 47, 49)", "Damages, including a cover purchase (CISG 74 to 76)"],
  ["All-risks cargo insurance, 110 % of value: EUR 633,600", "Factory acceptance test before signing the note", "Fixed delivery date, penalty for delay", "Warranty and the right to claim after paying", "Euro forward contract against the tugrik"],
  [{ text: "Instructions: ", options: { bold: true } }, { text: "to the insurer (Mongol Daatgal), cover from the seller's door to the lab; to the carrier, written notice of any damage within 14 days of receipt. " }, { text: "The note is unconditional: ", options: { bold: true } }, { text: "defects are claimed under the contract, after payment." }]);

// ================= 13. Breaches =================
{
  const s = content("Rights and duties", "Possible Breaches and the Remedies",
    "A late licence, an unpaid note and a re-export are the three breaches to plan for",
    "(CISG, 1980; ICC, 1995)",
    "Student B, about 40 seconds. Here is what can go wrong and who answers for it. On the seller's side the typical breach is delay, because the licence takes longer than planned. On the buyer's side there are two serious ones. The first is non-payment of the note, answered by protest, court action and the Eksfin guarantee. The second is specific to dual-use goods: a false end-user statement or a re-export. That is not only a breach of contract, it is a breach of export control law, and the seller is exposed too. This is why the compliance clauses matter as much as the payment clauses.");
  const by = (t) => cell(t, t === "Seller" ? SELLER : t === "Buyer" ? BUYER : OTHER);
  const b = (a, w, r) => [cell(a, { bold: true }), by(w), cell(r)];
  table(s, [
    [hdr("Breach"), hdr("By", { align: "center" }), hdr("Remedy or prevention")],
    b("No export licence: late or no delivery", "Seller", "Extra time, then avoidance (CISG 47, 49); final date in the contract"),
    b("Defective or wrong goods", "Seller", "Notice in time; repair or replacement, price reduction, damages"),
    b("No notice of dispatch: buyer cannot insure", "Seller", "Damages; buyer insures under an open policy from the contract date"),
    b("Note not issued, or formally invalid", "Buyer", "Seller withholds shipment (CISG 71)"),
    b("False end-user statement, or re-export to Russia", "Buyer", "Contract ends, penalty, report to DEKSA"),
    b("Goods not collected at Ulaanbaatar", "Buyer", "Storage at the buyer's cost; seller may resell (CISG 85, 88)"),
    b("Note not paid, paid late or paid in part", "Buyer", "Protest, action on the note, interest, Eksfin claim"),
    b("Loss, damage or delay in transit", "Carrier", "Buyer claims on its insurance; the insurer recovers from the carrier"),
  ], MX, TOP, [4.6, 1.2, 6.533], { rowH: [0.4, ...Array(8).fill(0.5)], name: "Breach table" });
  s.addText("A false end-user statement breaks export control law as well as the contract, and the seller is exposed too", { x: MX, y: 5.98, w: CW, h: 0.62, shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: [10, 10, 0, 0], isTextBox: true, objectName: "Compliance callout" });
}

// ================= 14. Assessment =================
pres.addSection({ title: "Assessment" });
{
  const s = content("Assessment", "Our Critical Assessment",
    "The given terms work, but we would secure the note and insure the cargo by contract",
    "(ICC, 2019; European Commission, 2025)",
    "Student B, about 40 seconds. Our conclusion is that the given terms work, but they are unbalanced. The seller gives ninety days of credit to a new buyer in a high-risk country, with no bank behind the note. We would ask Khan Bank to add its aval, which turns company risk into bank risk. We would also change CPT to CIP so that the cargo is insured by contract and not by good advice. And we leave one open question: the control entry for neural computers has been on the list for decades, while most AI today runs on ordinary accelerator chips. Whether the list still controls the right hardware would be worth researching. Thank you. We welcome your questions and feedback.");
  const recs = [
    ["The payment term is too weak", "Given: an unsecured note, 90 days, a first deal, country risk 6 of 7.", "We propose: ", "a bank aval from Khan Bank on the note, or a 20 to 30 % advance."],
    ["CPT leaves the cargo uninsured", "Given: neither party has a duty to insure.", "We propose: ", "CIP, where the seller must buy all-risks cover."],
    ["Compliance is a commercial term", "Given: a licensed dual-use item and a diversion risk.", "We propose: ", "licence as condition precedent, a no-re-export clause and an audit right."],
    ["The agent should share the risk", "Given: commission without responsibility for payment.", "We propose: ", "a del credere commission, so the agent answers for the buyer's payment."],
  ];
  const cw = 4.05, chh = 2.5;
  recs.forEach((r, i) => {
    const x = MX + (i % 2) * (cw + 0.2), y = TOP + Math.floor(i / 2) * (chh + 0.2);
    card(s, x, y, cw, chh, C.background2, "Recommendation card " + (i + 1));
    badge(s, i + 1, x + 0.15, y + 0.15, C.accent1);
    text(s, r[0], x + 0.68, y + 0.15, cw - 0.83, 0.4, { fontSize: 16, bold: true, valign: "middle", objectName: "Recommendation title " + (i + 1) });
    text(s, [{ text: r[1], options: { breakLine: true } }, { text: r[2], options: { bold: true, color: C.accent1 } }, { text: r[3] }], x + 0.18, y + 0.75, cw - 0.36, chh - 0.9, { fontSize: 15, paraSpaceAfter: 8, objectName: "Recommendation " + (i + 1) });
  });
  card(s, 9.0, TOP, 3.83, 2.9, C.text2, "Research question card");
  text(s, [
    { text: "Open question for research", options: { bold: true, fontSize: 16, color: C.accent2, breakLine: true } },
    { text: "Entry 4A004.b has been on the control list for decades. Does it still describe the hardware that matters, now that AI runs mostly on conventional accelerator chips?", options: { color: C.background1 } },
  ], 9.2, TOP + 0.1, 3.43, 2.7, { fontSize: 15, valign: "middle", paraSpaceAfter: 8, objectName: "Research question" });
  card(s, 9.0, 4.5, 3.83, 2.1, C.accent2, "Conclusion card");
  text(s, [
    { text: "Conclusion", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "The terms as given favour the buyer. With an aval and CIP, both sides are protected.", options: { breakLine: true } },
    { text: "Questions and feedback welcome", options: { bold: true } },
  ], 9.2, 4.6, 3.43, 1.9, { fontSize: 15, valign: "middle", paraSpaceAfter: 8, objectName: "Conclusion" });
}

// ================= 15. References =================
pres.addSection({ title: "Appendix" });
{
  const s = content("Appendix", "References", "Sources cited on the slides", null, "Shown for reference, not presented.");
  const ref = (items) => items.map((t, i) => ({ text: t, options: { breakLine: i < items.length - 1, paraSpaceAfter: 7, bullet: false } }));
  text(s, ref([
    "CISG (1980). United Nations Convention on Contracts for the International Sale of Goods. Vienna. uncitral.un.org/en/texts/salegoods/conventions/sale_of_goods/cisg",
    "DNB (n.d.). International payments. dnb.no/en/business/help-and-guidance/international-payments",
    "EFTA (n.d.). Trade relations: Mongolia. efta.int/trade-relations/free-trade-network/mongolia",
    "Eksfin (n.d.). Supplier credit guarantee. eksfin.no/en/products/supplier-credit",
    "European Commission (2024). Frequently asked questions on the \"no re-export to Russia\" clause. finance.ec.europa.eu",
    "European Commission (2025). Commission Delegated Regulation (EU) 2025/2003 replacing Annex I to Regulation (EU) 2021/821. eur-lex.europa.eu/eli/reg_del/2025/2003/oj/eng",
    "Geneva Convention (1930). Convention Providing a Uniform Law for Bills of Exchange and Promissory Notes.",
    "Government of Norway (n.d.). What is Norway's export control system? regjeringen.no/en/topics/foreign-affairs/export-control",
  ]), MX, TOP, 6.0, 5.2, { fontSize: 12, objectName: "References left" });
  text(s, ref([
    "Government of Norway (n.d.). New sanctions against Russia implemented in Norwegian law. regjeringen.no",
    "ICC (1995). Uniform Rules for Collections (URC 522). ICC Publication No. 522. Paris: International Chamber of Commerce.",
    "ICC (2019). Incoterms® 2020. ICC Publication No. 723E. Paris: International Chamber of Commerce.",
    "ICC. Guide to Export-Import Basics. ICC Publication No. 685. Paris: International Chamber of Commerce.",
    "Montreal Convention (1999). Convention for the Unification of Certain Rules for International Carriage by Air.",
    "OECD (2025). Country risk classification of the participants to the Arrangement on Officially Supported Export Credits.",
    "PwC (n.d.). Worldwide Tax Summaries: Mongolia, other taxes. taxsummaries.pwc.com/mongolia/corporate/other-taxes",
    "Turkish Airlines (n.d.). Flights to Mongolia. turkishairlines.com",
    "Lecture slides and seminar materials, International Trade Practices, Corvinus University, 2026.",
  ]), 6.83, TOP, 6.0, 5.2, { fontSize: 12, objectName: "References right" });
}

// ================= 16. GenAI =================
{
  const s = content("Appendix", "Use of Generative AI", "Required disclosure: tool, purpose and prompts", null, "Shown as required by the syllabus. Add any later prompts before uploading.");
  pill(s, "Tool and Purpose", MX, TOP, 5.3);
  list(s, [
    ["Tool: ", "Claude (Anthropic), used through Claude Code"],
    ["Used for: ", "planning the structure, researching and drafting the slide text, building the slide file"],
    ["Not used for: ", "drawings or illustrations. The deck contains none made by AI"],
    ["Checked by the team: ", "facts against the cited sources"],
  ], MX, 1.97, 5.3, 3.2, { name: "GenAI use", gap: 8 });
  pill(s, "Prompts", 6.1, TOP, 6.73);
  list(s, [
    "\"i need to do presentation, my case is 33, do the plan first, then we execute\" (syllabus and mark sheet attached)",
    "\"we have only 2 people in the group, so we will split work of 3, sent me comperhensive slide content, so after you can make the slides\"",
    "\"use these presentations as reference and create slides for our presentation\" (reference deck attached)",
  ], 6.1, 1.97, 6.73, 3.2, { name: "Prompts", numbered: true, gap: 8 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
