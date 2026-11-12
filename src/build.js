/**
 * BUILD MODE — Strategic Brand & Business Strategy Presentation
 * Generates build_mode_strategy.pptx (12 slides, 16:9) with PptxGenJS.
 *
 *   npm install && npm run build
 *
 * Everything on the slides is a native, editable PowerPoint object: text,
 * shapes and lines. No slide is flattened into an image.
 */
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply-theme");

const OUT = path.join(__dirname, "..", "build_mode_strategy.pptx");

// ---------------------------------------------------------------------------
// Brand theme. Change a colour or the font here and the whole deck follows.
// ---------------------------------------------------------------------------
const FONT = "Inter"; // bundled in assets/fonts; swap for "Aptos" or "Arial" if preferred
const THEME = {
  name: "BUILD MODE",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "101827", // deep navy — primary dark background, text on light
    lt1: "FFFFFF", // white
    dk2: "18263A", // dark blue — cards on dark slides
    lt2: "F2F4F7", // soft grey — cards on light slides
    accent1: "477BFF", // electric blue — the single accent
    accent2: "91A0B6", // muted blue-grey — secondary text on dark
    accent3: "D5DBE5", // hairlines on light
    accent4: "56657C", // secondary text on light
    accent5: "2B3D58", // hairlines and outlines on dark
    accent6: "DCE6FF", // pale blue tint
    hlink: "477BFF",
    folHlink: "91A0B6",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in, 16:9
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "BUILD MODE — Strategic Blueprint";
pres.subject = "Strategic brand and business strategy";
pres.author = "BUILD MODE";
pres.company = "BUILD MODE";

const S = pres.SchemeColor;
const C = {
  navy: S.text1,
  white: S.background1,
  blue2: S.text2,
  grey: S.background2,
  accent: S.accent1,
  mute: S.accent2,
  hair: S.accent3,
  sub: S.accent4,
  hairDark: S.accent5,
  tint: S.accent6,
};

// Grid
const W = 13.333;
const M = 0.6; // outer margin
const CW = W - 2 * M; // content width
const Y0 = 2.4; // top of the content area on every content slide
const FOOT_Y = 6.98;

// ---------------------------------------------------------------------------
// Layouts (slide frames): title/eyebrow placeholders, footer, slide number
// ---------------------------------------------------------------------------
function frame(name, dark, titleOpts) {
  const fg = dark ? C.white : C.navy;
  pres.defineSlideMaster({
    title: name,
    background: { color: dark ? C.navy : C.white },
    objects: [
      {
        text: {
          text: "BUILD MODE   /   STRATEGIC BLUEPRINT",
          options: { x: M, y: FOOT_Y, w: 5, h: 0.25, margin: 0, fontSize: 8.5, bold: true, charSpacing: 2.5, color: C.mute },
        },
      },
      {
        placeholder: {
          options: { name: "eyebrow", type: "body", x: M, y: 0.55, w: 8, h: 0.28, margin: 0, fontSize: 11, bold: true, charSpacing: 3, color: C.accent, valign: "top" },
          text: "",
        },
      },
      {
        placeholder: {
          options: Object.assign(
            { name: "title", type: "title", x: M, y: 0.95, w: 11.4, h: 1.15, margin: 0, fontSize: 30, bold: true, color: fg, align: "left", valign: "top" },
            titleOpts || {}
          ),
          text: "",
        },
      },
    ],
    slideNumber: { x: W - M - 0.8, y: FOOT_Y, w: 0.8, h: 0.25, margin: 0, fontSize: 8.5, bold: true, align: "right", color: C.mute },
  });
}

pres.defineSlideMaster({
  title: "BM Cover",
  background: { color: C.navy },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: M, y: 1.75, w: 9.6, h: 3.1, margin: 0, fontSize: 60, bold: true, color: C.white, align: "left", valign: "top" },
        text: "",
      },
    },
    {
      placeholder: {
        options: { name: "subtitle", type: "body", x: M, y: 5.05, w: 8.6, h: 0.5, margin: 0, fontSize: 20, color: C.mute, valign: "top" },
        text: "",
      },
    },
  ],
});
frame("BM Light", false);
frame("BM Dark", true);
frame("BM Statement", false, { y: 1.7, w: 7.3, h: 2.9, fontSize: 54 });

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
let section = "";
function newSlide(layout, eyebrow, title, notes) {
  const s = pres.addSlide({ masterName: layout, sectionTitle: section });
  if (eyebrow) s.addText(eyebrow, { placeholder: "eyebrow", charSpacing: 3 });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
/** Plain text box with zero inset so text aligns to the grid. */
function text(s, str, o) {
  s.addText(str, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 14 }, o));
}
function lede(s, str, dark) {
  text(s, str, { x: M, y: 1.57, w: 10.5, h: 0.4, fontSize: 16, color: dark ? C.mute : C.sub });
}
function box(s, o, name) {
  const opt = { x: o.x, y: o.y, w: o.w, h: o.h, objectName: name };
  if (o.fill) opt.fill = { color: o.fill };
  opt.line = o.line ? { color: o.line, width: o.lw || 0.75, dashType: o.dash || "solid" } : { type: "none" };
  s.addShape(pres.shapes.RECTANGLE, opt);
}
function hline(s, x, y, w, color, extra) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: Object.assign({ color, width: 0.75 }, extra || {}) });
}
function vline(s, x, y, h, color) {
  s.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color, width: 0.75 } });
}
/** Small uppercase label. */
function label(s, str, o) {
  text(s, str, Object.assign({ h: 0.25, fontSize: 10.5, bold: true, charSpacing: 2.5, color: C.accent }, o));
}
/** Outlined tag + sentence, used wherever content is unvalidated. */
function flag(s, tag, sentence, y, dark) {
  const tw = 0.3 + tag.length * 0.105;
  s.addText(tag, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16,
    x: M, y, w: tw, h: 0.32, margin: 0, align: "center", valign: "middle",
    fontSize: 9, bold: true, charSpacing: 2, color: C.accent,
    line: { color: C.accent, width: 1 }, objectName: "Flag",
  });
  text(s, sentence, { x: M + tw + 0.2, y, w: CW - tw - 0.2, h: 0.32, valign: "middle", fontSize: 13, color: dark ? C.mute : C.sub });
}
/** Row of five B-U-I-L-D tiles: the deck's recurring motif. */
function buildTiles(s, x, y, size, gap, vertical, o) {
  "BUILD".split("").forEach((ch, i) => {
    s.addText(ch, {
      shape: pres.shapes.RECTANGLE,
      x: vertical ? x : x + i * (size + gap), y: vertical ? y + i * (size + gap) : y, w: size, h: size,
      margin: 0, align: "center", valign: "middle", fontSize: o.fontSize, bold: true, color: C.white,
      fill: { color: i === 0 ? C.accent : C.blue2 },
      line: i === 0 ? { type: "none" } : { color: C.hairDark, width: 0.75 },
      objectName: `BUILD tile ${ch}`,
    });
  });
}
const cols = (n, gap) => {
  const w = (CW - gap * (n - 1)) / n;
  return { w, x: (i) => M + i * (w + gap) };
};
const num = (i) => String(i + 1).padStart(2, "0");

// ===========================================================================
// 01 — COVER
// ===========================================================================
section = "Vision & Audience";
pres.addSection({ title: section });
{
  const s = pres.addSlide({ masterName: "BM Cover", sectionTitle: section });
  box(s, { x: M, y: 0.68, w: 0.16, h: 0.16, fill: C.accent }, "Brand mark");
  text(s, "BUILD MODE", { x: M + 0.3, y: 0.62, w: 4, h: 0.28, valign: "middle", fontSize: 13, bold: true, charSpacing: 4, color: C.white });
  s.addText(
    [
      { text: "From Idea to", options: { breakLine: true } },
      { text: "AI-Powered", options: { color: C.accent, breakLine: true } },
      { text: "Online Business." },
    ],
    { placeholder: "title" }
  );
  s.addText("A strategic blueprint for building a global AI education brand.", { placeholder: "subtitle" });
  buildTiles(s, W - M - 0.92, 0.62, 0.92, 0.12, true, { fontSize: 30 });
  hline(s, M, 6.4, 10.4, C.hairDark);
  label(s, "PRACTICAL AI EDUCATION FOR ONLINE BUSINESS", { x: M, y: 6.6, w: 6, color: C.mute });
  label(s, "STRATEGIC PLANNING DECK", { x: 7, y: 6.6, w: 4, align: "right", color: C.mute });
  s.addNotes("Strategic planning deck, not a sales pitch. Hypotheses are flagged throughout; no market data has been collected yet.");
}

// ===========================================================================
// 02 — THE BUSINESS VISION
// ===========================================================================
{
  const s = newSlide("BM Light", "THE BUSINESS VISION", "What Are We Building?",
    "Central message: a digital education business that earns revenue without personal client meetings or one-to-one coaching.");
  lede(s, "Five design choices define the business model.");

  const pw = 4.7, ph = 4.15;
  box(s, { x: M, y: Y0, w: pw, h: ph, fill: C.navy }, "Central idea panel");
  label(s, "THE CENTRAL IDEA", { x: M + 0.4, y: Y0 + 0.4, w: pw - 0.8 });
  text(s, "Build a digital education business that generates revenue without requiring personal client meetings or one-to-one coaching.",
    { x: M + 0.4, y: Y0 + 0.85, w: pw - 0.8, h: 2.4, fontSize: 20, bold: true, color: C.white });
  hline(s, M + 0.4, Y0 + ph - 0.75, pw - 0.8, C.hairDark);
  label(s, "NO CLIENT MEETINGS   /   NO 1:1 COACHING", { x: M + 0.4, y: Y0 + ph - 0.55, w: pw - 0.8, color: C.mute, fontSize: 9.5, charSpacing: 1.5 });

  const rows = [
    ["Faceless brand", "A global, faceless AI education brand."],
    ["AI-built businesses", "Teaching people to build and grow online businesses using AI."],
    ["Self-paced learning", "Self-paced courses and practical business systems."],
    ["USD pricing", "USD-priced digital products."],
    ["Full automation", "Automated checkout, delivery and onboarding."],
  ];
  const rx = M + pw + 0.55, rw = W - M - rx, rh = ph / rows.length;
  rows.forEach(([head, body], i) => {
    const y = Y0 + i * rh;
    hline(s, rx, y, rw, C.hair);
    text(s, num(i), { x: rx, y: y + 0.24, w: 0.6, h: 0.3, fontSize: 12, bold: true, color: C.accent });
    text(s, head, { x: rx + 0.7, y: y + 0.2, w: 2.75, h: 0.4, fontSize: 16, bold: true, color: C.navy });
    text(s, body, { x: rx + 3.5, y: y + 0.22, w: rw - 3.5, h: 0.55, fontSize: 14, color: C.sub });
  });
  hline(s, rx, Y0 + ph, rw, C.hair);
}

// ===========================================================================
// 03 — TARGET AUDIENCE
// ===========================================================================
{
  const s = newSlide("BM Light", "TARGET AUDIENCE", "People Who Want to Build an Independent Income Stream.",
    "Eight candidate segments. The ambition lines are working assumptions to be tested in customer research.");
  const seg = [
    ["Solo entrepreneurs", "Run a lean business without hiring a team."],
    ["Independent freelancers", "Turn client skills into scalable digital income."],
    ["Housewives seeking income from home", "Earn flexibly around family life."],
    ["Part-timers", "Grow a side income into something more stable."],
    ["Aspiring digital entrepreneurs", "Launch a first online business the right way."],
    ["Content creators", "Monetize an audience beyond sponsorships."],
    ["Micro-business owners", "Work faster and sell more online with AI."],
    ["Career switchers", "Build a new income path beyond employment."],
  ];
  const g = cols(4, 0.22), ch = 1.9;
  seg.forEach(([name, amb], i) => {
    const x = g.x(i % 4), y = Y0 + Math.floor(i / 4) * (ch + 0.22);
    box(s, { x, y, w: g.w, h: ch, fill: C.grey }, `Segment ${num(i)}`);
    text(s, num(i), { x: x + 0.25, y: y + 0.22, w: 1, h: 0.25, fontSize: 12, bold: true, color: C.accent });
    text(s, name, { x: x + 0.25, y: y + 0.52, w: g.w - 0.5, h: 0.58, fontSize: 16, bold: true, color: C.navy });
    text(s, amb, { x: x + 0.25, y: y + 1.14, w: g.w - 0.5, h: 0.66, fontSize: 14, color: C.sub });
  });
  text(s, "Ambitions shown are working assumptions, to be tested in customer research.",
    { x: M, y: 6.6, w: CW, h: 0.25, fontSize: 10.5, color: C.sub });
}

// ===========================================================================
// 04 — AUDIENCE PRIORITIZATION
// ===========================================================================
{
  const s = newSlide("BM Light", "AUDIENCE PRIORITIZATION", "Where We Focus First: Three Audience Tiers",
    "These tiers are audience hypotheses, not validated market rankings. The order must be tested before it drives product or marketing decisions.");
  lede(s, "A proposed order of focus, to be tested rather than assumed.");

  const tiers = [
    { tag: "TIER 1   /   PRIMARY", name: "One-person business builders", chips: ["Aspiring digital entrepreneurs", "Solo entrepreneurs", "Freelancers", "Part-timers"] },
    { tag: "TIER 2   /   SECONDARY", name: "People seeking independent income", chips: ["Housewives", "Working adults", "Career switchers"] },
    { tag: "TIER 3   /   EXPANSION", name: "Existing online business operators", chips: ["Content creators", "Micro-business owners", "Digital product sellers"] },
  ];
  const bh = 1.12, bg = 0.16, cx0 = 5.25, cg = 0.16;
  const cw = (W - M - 0.3 - cx0 - 3 * cg) / 4;
  tiers.forEach((t, i) => {
    const y = Y0 + i * (bh + bg);
    const st = [
      { fill: C.navy, tag: C.accent, name: C.white, chipFill: C.accent, chipText: C.white },
      { fill: C.grey, tag: C.accent, name: C.navy, chipFill: C.white, chipText: C.navy },
      { line: C.mute, dash: "dash", tag: C.sub, name: C.navy, chipLine: C.hair, chipText: C.sub },
    ][i];
    box(s, { x: M, y, w: CW, h: bh, fill: st.fill, line: st.line, dash: st.dash, lw: 1 }, `Tier ${i + 1} band`);
    label(s, t.tag, { x: M + 0.35, y: y + 0.24, w: 4.2, color: st.tag });
    text(s, t.name, { x: M + 0.35, y: y + 0.54, w: 4.2, h: 0.38, fontSize: 17, bold: true, color: st.name });
    t.chips.forEach((c, j) => {
      s.addText(c, {
        shape: pres.shapes.RECTANGLE, x: cx0 + j * (cw + cg), y: y + 0.24, w: cw, h: bh - 0.48,
        margin: [0, 6, 0, 6], align: "center", valign: "middle", fontSize: 13, bold: i === 0, color: st.chipText,
        fill: st.chipFill ? { color: st.chipFill } : undefined,
        line: st.chipLine ? { color: st.chipLine, width: 0.75 } : { type: "none" },
        objectName: `Tier ${i + 1} segment`,
      });
    });
  });
  flag(s, "HYPOTHESIS", "These are audience hypotheses, not validated market rankings.", 6.33);
}

// ===========================================================================
// 05 — CORE CUSTOMER PROBLEMS
// ===========================================================================
section = "Customer";
pres.addSection({ title: section });
{
  const s = newSlide("BM Dark", "CORE CUSTOMER PROBLEMS", "They Want to Build an Online Business, But They Don't Know How to Execute.",
    "Eight pain points grouped into four themes. These are assumed problems drawn from the brief; complaint analysis is a next step.");
  const groups = [
    ["DIRECTION", ["Too many business ideas, no clear direction.", "Unsure what people are willing to buy."]],
    ["AI CLARITY", ["Don't know which AI tools to use.", "Know how to generate AI content but not how to build a business."]],
    ["OFFER & SALES", ["Struggle to create products and offers.", "Don't know how to market and sell."]],
    ["SYSTEMS", ["Waste time on repetitive manual work.", "Have no repeatable business system."]],
  ];
  const g = cols(4, 0.22), ch = 1.72;
  groups.forEach(([name, pains], i) => {
    const x = g.x(i);
    label(s, name, { x, y: Y0, w: g.w });
    pains.forEach((p, j) => {
      const y = Y0 + 0.42 + j * (ch + 0.2);
      box(s, { x, y, w: g.w, h: ch, fill: C.blue2, line: C.hairDark }, `Pain ${num(i * 2 + j)}`);
      text(s, num(i * 2 + j), { x: x + 0.25, y: y + 0.22, w: 1, h: 0.25, fontSize: 12, bold: true, color: C.mute });
      text(s, p, { x: x + 0.25, y: y + 0.58, w: g.w - 0.5, h: 1.0, fontSize: 15, color: C.white });
    });
  });
}

// ===========================================================================
// 06 — CUSTOMER JOURNEY
// ===========================================================================
{
  const s = newSlide("BM Light", "CUSTOMER JOURNEY", "From Uncertainty to an AI-Powered Business.",
    "Seven stages, each framed as the question the customer is asking at that moment.");
  lede(s, "Seven questions every new online business builder has to answer, in order.");
  const stages = [
    ["DISCOVER", "What business should I start?"],
    ["VALIDATE", "Is there real market demand?"],
    ["BUILD", "What should I create and sell?"],
    ["MARKET", "How do I attract customers?"],
    ["SELL", "How do I convert interest into revenue?"],
    ["AUTOMATE", "How do I reduce repetitive work?"],
    ["GROW", "How do I improve and scale?"],
  ];
  const cw = CW / 7, d = 0.52, ly = 3.3;
  label(s, "UNCERTAINTY", { x: M, y: 2.55, w: 3, color: C.sub });
  label(s, "AI-POWERED BUSINESS", { x: W - M - 4, y: 2.55, w: 4, align: "right" });
  hline(s, M, ly, CW, C.mute, { width: 1.25, endArrowType: "triangle" });
  stages.forEach(([name, q], i) => {
    const x = M + i * cw, last = i === stages.length - 1;
    s.addText(num(i), {
      shape: pres.shapes.OVAL, x, y: ly - d / 2, w: d, h: d, margin: 0, align: "center", valign: "middle",
      fontSize: 11, bold: true, color: C.white, fill: { color: last ? C.accent : C.navy }, line: { color: C.white, width: 3 },
      objectName: `Stage ${num(i)} node`,
    });
    text(s, name, { x, y: ly + 0.5, w: cw - 0.2, h: 0.35, fontSize: 16, bold: true, color: last ? C.accent : C.navy });
    text(s, q, { x, y: ly + 0.92, w: cw - 0.3, h: 1.2, fontSize: 14, color: C.sub });
  });
  box(s, { x: M, y: 5.7, w: CW, h: 0.8, fill: C.grey }, "Journey takeaway");
  text(s, [
    { text: "Design principle  ", options: { bold: true, color: C.navy } },
    { text: "Each stage answers one question before the next begins. The curriculum follows the same order.", options: { color: C.sub } },
  ], { x: M + 0.3, y: 5.7, w: CW - 0.6, h: 0.8, valign: "middle", fontSize: 14 });
}

// ===========================================================================
// 07 — BRAND POSITIONING
// ===========================================================================
section = "Brand & Method";
pres.addSection({ title: section });
{
  const s = newSlide("BM Statement", "BRAND POSITIONING", "Practical AI Education for Online Business.",
    "AI is the enabler. Business building is the outcome. This is the single idea every product and message should reinforce.");
  box(s, { x: M, y: 1.2, w: 0.16, h: 0.16, fill: C.accent }, "Brand mark");
  text(s, "BUILD MODE", { x: M + 0.3, y: 1.14, w: 4, h: 0.28, valign: "middle", fontSize: 13, bold: true, charSpacing: 4, color: C.navy });
  text(s, "Learn how to use AI to research, build, launch, monetize and automate an online business.",
    { x: M, y: 4.8, w: 6.6, h: 0.75, fontSize: 18, color: C.sub });
  hline(s, M, 5.8, 7.0, C.hair);
  label(s, "RESEARCH   /   BUILD   /   LAUNCH   /   MONETIZE   /   AUTOMATE", { x: M, y: 6.0, w: 7.2 });

  const px = 8.35, pw = W - M - px;
  box(s, { x: px, y: 0.6, w: pw, h: 5.95, fill: C.navy }, "Principle panel");
  label(s, "THE PRINCIPLE", { x: px + 0.45, y: 1.05, w: pw - 0.9 });
  text(s, [
    { text: "AI is the enabler.", options: { color: C.mute, breakLine: true } },
    { text: "Business building is the outcome.", options: { color: C.white } },
  ], { x: px + 0.45, y: 2.6, w: pw - 0.85, h: 3.5, valign: "bottom", fontSize: 32, bold: true, paraSpaceAfter: 10 });
}

// ===========================================================================
// 08 — SIGNATURE METHODOLOGY
// ===========================================================================
{
  const s = newSlide("BM Dark", "SIGNATURE METHODOLOGY", "The BUILD Method.",
    "The signature five-stage methodology. Every product in the ecosystem maps back to one or more of these stages.");
  lede(s, "A five-stage path from first idea to an automated, growing business.", true);
  const steps = [
    ["B", "BUSINESS DISCOVERY", "Research opportunities, customer problems, competitors and validate ideas."],
    ["U", "UNDERSTAND & USE AI", "Learn AI tools, prompting, research techniques and effective AI collaboration."],
    ["I", "IMPLEMENT YOUR BUSINESS", "Create products, develop offers, build content and establish workflows."],
    ["L", "LAUNCH & MONETIZE", "Develop pricing, marketing, sales funnels and launch strategies."],
    ["D", "DEVELOP & AUTOMATE", "Improve performance, automate tasks, track results and grow."],
  ];
  const g = cols(5, 0.2), ch = 4.2;
  steps.forEach(([letter, name, desc], i) => {
    const x = g.x(i);
    box(s, { x, y: Y0, w: g.w, h: ch, fill: C.blue2, line: C.hairDark }, `BUILD stage ${letter}`);
    text(s, letter, { x: x + 0.28, y: Y0 + 0.15, w: 1.6, h: 1.6, valign: "middle", fontSize: 108, bold: true, color: C.white });
    text(s, num(i), { x: x + g.w - 0.9, y: Y0 + 0.3, w: 0.62, h: 0.25, align: "right", fontSize: 11, bold: true, color: C.mute });
    hline(s, x + 0.3, Y0 + 1.85, g.w - 0.6, C.hairDark);
    text(s, name, { x: x + 0.3, y: Y0 + 2.03, w: g.w - 0.55, h: 0.5, fontSize: 12, bold: true, charSpacing: 1, color: C.accent });
    text(s, desc, { x: x + 0.3, y: Y0 + 2.6, w: g.w - 0.55, h: 1.5, fontSize: 14, color: C.grey });
  });
}

// ===========================================================================
// 09 — DIFFERENTIATION
// ===========================================================================
{
  const s = newSlide("BM Light", "DIFFERENTIATION", "What Makes BUILD MODE Different?",
    "A proposed differentiation hypothesis. Competitor research has not been done yet, so the left column is an assumption about the category, not a finding.");
  lede(s, "Shifting the focus from AI tools to business outcomes.");
  const pairs = [
    ["Teaches individual tools.", "Teaches business outcomes."],
    ["Focuses on prompts.", "Focuses on repeatable workflows."],
    ["Random AI tutorials.", "Structured methodology."],
    ["Tool-specific knowledge.", "Tool-independent principles."],
    ["Information consumption.", "Practical implementation."],
  ];
  const pw = 5.55, gapW = CW - 2 * pw, rx = M + pw + gapW, hh = 0.7, rh = 0.57;
  const ph = hh + pairs.length * rh + 0.1;
  box(s, { x: M, y: Y0, w: pw, h: ph, fill: C.grey }, "Typical AI education");
  box(s, { x: rx, y: Y0, w: pw, h: ph, fill: C.navy }, "BUILD MODE direction");
  label(s, "TYPICAL AI EDUCATION", { x: M + 0.4, y: Y0, w: pw - 0.8, h: hh, valign: "middle", color: C.sub });
  label(s, "BUILD MODE DIRECTION", { x: rx + 0.4, y: Y0, w: pw - 0.8, h: hh, valign: "middle" });
  pairs.forEach(([a, b], i) => {
    const y = Y0 + hh + i * rh;
    hline(s, M + 0.4, y, pw - 0.8, C.hair);
    hline(s, rx + 0.4, y, pw - 0.8, C.hairDark);
    text(s, a, { x: M + 0.4, y, w: pw - 0.8, h: rh, valign: "middle", fontSize: 15, color: C.sub });
    text(s, b, { x: rx + 0.4, y, w: pw - 0.8, h: rh, valign: "middle", fontSize: 15, bold: true, color: C.white });
    hline(s, M + pw + 0.22, y + rh / 2, gapW - 0.44, C.accent, { width: 1.25, endArrowType: "triangle" });
  });
  flag(s, "HYPOTHESIS", "Proposed differentiation hypothesis. Requires market validation against real competitor products.", 6.33);
}

// ===========================================================================
// 10 — PRODUCT ECOSYSTEM
// ===========================================================================
{
  const s = newSlide("BM Light", "PRODUCT ECOSYSTEM", "One Core Methodology. Multiple Digital Products.",
    "The flagship course is the business. Supporting products extend it; trackers and templates are not the positioning.");
  lede(s, "The method is the product. Everything else supports it.");
  const fh = 2.05;
  box(s, { x: M, y: Y0, w: CW, h: fh, fill: C.navy }, "Flagship course");
  label(s, "FLAGSHIP COURSE", { x: M + 0.45, y: Y0 + 0.35, w: 5 });
  text(s, "The BUILD Method", { x: M + 0.45, y: Y0 + 0.68, w: 6.6, h: 0.65, valign: "middle", fontSize: 36, bold: true, color: C.white });
  text(s, "Complete AI-Powered Online Business System.", { x: M + 0.45, y: Y0 + 1.4, w: 6.6, h: 0.35, fontSize: 17, color: C.mute });
  const ts = 0.8, tg = 0.1;
  buildTiles(s, W - M - 0.45 - (5 * ts + 4 * tg), Y0 + (fh - ts) / 2, ts, tg, false, { fontSize: 26 });

  const prods = [
    ["Mini courses", "Focused lessons on one stage of the method."],
    ["Business templates", "Ready-to-adapt plans, offers and launch docs."],
    ["Workflow trackers", "Simple tools to track progress by stage."],
    ["Prompt libraries", "Curated prompts that support each workflow."],
  ];
  const g = cols(4, 0.22), busY = Y0 + fh + 0.42, cy = busY + 0.26, ch = 6.5 - cy;
  vline(s, W / 2, Y0 + fh, 0.42, C.mute);
  hline(s, g.x(0) + g.w / 2, busY, g.x(3) - g.x(0), C.mute);
  label(s, "SUPPORTING PRODUCTS", { x: W / 2 + 0.15, y: Y0 + fh + 0.1, w: 3, color: C.sub, fontSize: 9.5 });
  prods.forEach(([name, desc], i) => {
    const x = g.x(i);
    vline(s, x + g.w / 2, busY, 0.26, C.mute);
    box(s, { x, y: cy, w: g.w, h: ch, fill: C.grey }, `Supporting product ${num(i)}`);
    text(s, name, { x: x + 0.25, y: cy + 0.2, w: g.w - 0.5, h: 0.32, fontSize: 16, bold: true, color: C.navy });
    text(s, desc, { x: x + 0.25, y: cy + 0.58, w: g.w - 0.5, h: 0.7, fontSize: 14, color: C.sub });
  });
}

// ===========================================================================
// 11 — MARKET VALIDATION FRAMEWORK
// ===========================================================================
section = "Validation & Next Steps";
pres.addSection({ title: section });
{
  const s = newSlide("BM Light", "MARKET VALIDATION FRAMEWORK", "Before Building, We Must Prove Five Things.",
    "Five validation gates. None has been passed yet: this slide defines the questions, not the answers.");
  lede(s, "Five validation gates stand between this strategy and course production.");
  const gates = [
    ["DEMAND", "Are people actively searching for solutions?"],
    ["PAIN", "Is the problem important enough to pay to solve?"],
    ["COMPETITION", "What are existing products missing?"],
    ["PURCHASE INTENT", "Are customers actually willing to spend?"],
    ["PRODUCT FIT", "Can we deliver through self-paced products?"],
  ];
  const gap = 0.22, g = cols(5, gap), ch = 3.05;
  gates.forEach(([name, q], i) => {
    const x = g.x(i);
    box(s, { x, y: Y0, w: g.w, h: ch, fill: C.white, line: C.hair, lw: 1 }, `Gate ${i + 1}`);
    text(s, String(i + 1), { x: x + 0.28, y: Y0 + 0.22, w: 1, h: 0.8, valign: "middle", fontSize: 48, bold: true, color: C.navy });
    label(s, "GATE", { x: x + g.w - 1.1, y: Y0 + 0.3, w: 0.82, align: "right", color: C.mute, fontSize: 9.5 });
    hline(s, x + 0.28, Y0 + 1.2, g.w - 0.56, C.hair);
    label(s, name, { x: x + 0.28, y: Y0 + 1.4, w: g.w - 0.5, fontSize: 11, charSpacing: 0.5 });
    text(s, q, { x: x + 0.28, y: Y0 + 1.78, w: g.w - 0.5, h: 1.2, fontSize: 14, color: C.sub });
    if (i < gates.length - 1) {
      s.addShape(pres.shapes.CHEVRON, {
        x: x + g.w + (gap - 0.13) / 2, y: Y0 + ch / 2 - 0.1, w: 0.13, h: 0.2, fill: { color: C.accent }, line: { type: "none" }, objectName: "Gate connector",
      });
    }
  });
  flag(s, "STATUS: OPEN", "All five gates are unproven today. Course production begins only once each one is passed.", 6.0);
}

// ===========================================================================
// 12 — NEXT STRATEGIC MOVES
// ===========================================================================
{
  const s = newSlide("BM Dark", "NEXT STRATEGIC MOVES", "Research First. Build Second. Scale After Validation.",
    "Closing: Don't just learn AI. Build something with it. Research, validate, build, launch, grow.");
  lede(s, "Six moves, in order, before any full course is produced.", true);
  const steps = [
    "Identify the strongest customer segment.",
    "Research competitors, courses and product pricing.",
    "Analyse customer complaints and unmet needs.",
    "Validate the BUILD Method against real customer problems.",
    "Select one flagship product and its first customer segment.",
    "Test demand before investing in full course production.",
  ];
  const lw = 6.95, rh = 0.6;
  steps.forEach((t, i) => {
    const y = Y0 + i * rh;
    hline(s, M, y, lw, C.hairDark);
    text(s, num(i), { x: M, y, w: 0.6, h: rh, valign: "middle", fontSize: 12, bold: true, color: C.accent });
    text(s, t, { x: M + 0.7, y, w: lw - 0.7, h: rh, valign: "middle", fontSize: 15, color: C.white });
  });
  hline(s, M, Y0 + steps.length * rh, lw, C.hairDark);

  const px = 7.95, pw = W - M - px, ph = steps.length * rh;
  box(s, { x: px, y: Y0, w: pw, h: ph, fill: C.accent }, "Closing statement");
  text(s, "BUILD MODE", { x: px + 0.4, y: Y0 + 0.38, w: pw - 0.8, h: 0.28, fontSize: 12, bold: true, charSpacing: 4, color: C.white });
  text(s, [
    { text: "Don't just learn AI.", options: { breakLine: true } },
    { text: "Build something with it." },
  ], { x: px + 0.4, y: Y0 + 0.9, w: pw - 0.75, h: ph - 1.25, valign: "bottom", fontSize: 24, bold: true, color: C.white, paraSpaceAfter: 6 });

  const seq = ["RESEARCH", "VALIDATE", "BUILD", "LAUNCH", "GROW"];
  const runs = [];
  seq.forEach((wd, i) => {
    runs.push({ text: wd, options: { color: C.white } });
    if (i < seq.length - 1) runs.push({ text: "   →   ", options: { color: C.accent } });
  });
  text(s, runs, { x: M, y: 6.3, w: CW, h: 0.35, valign: "middle", fontSize: 12, bold: true, charSpacing: 3 });
}

// ---------------------------------------------------------------------------
(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(`Wrote ${OUT} (${pres.slides.length} slides)`);
})();
