const pptxgen = require("pptxgenjs");
const fs = require("fs");

const accountId = process.argv[2] || "ACC-1001";
const data = JSON.parse(fs.readFileSync(`/home/claude/success_plan/data/payload_${accountId}.json`, "utf8"));

// ---- Palette: Midnight Executive ----
const NAVY = "1E2761";
const NAVY_DEEP = "141B49";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const SLATE = "44506B";
const GREY = "6B7280";
const LIGHT_BG = "F5F7FC";
const GREEN = "1E8A5F";
const AMBER = "B7791F";
const RED = "B4312F";

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

function statusColor(status) {
  const s = (status || "").toLowerCase();
  if (s.includes("complete")) return GREEN;
  if (s.includes("risk") || s.includes("down")) return RED;
  if (s.includes("progress") || s.includes("scheduled") || s.includes("open")) return AMBER;
  return GREY;
}
function sentimentColor(s) {
  s = (s || "").toLowerCase();
  if (s === "champion") return GREEN;
  if (s === "at risk") return RED;
  return AMBER;
}
function severityColor(s) {
  s = (s || "").toLowerCase();
  if (s === "high") return RED;
  if (s === "medium") return AMBER;
  return GREEN;
}
function healthColor(score) {
  if (score >= 75) return GREEN;
  if (score >= 50) return AMBER;
  return RED;
}
function formatMetricValue(value, unit) {
  const n = Number(value);
  if (Number.isNaN(n)) return String(value);
  if (unit && unit.startsWith("$")) return "$" + n.toLocaleString("en-US");
  if (n >= 1000) return n.toLocaleString("en-US");
  return String(n);
}
function formatUnitLabel(unit) {
  return unit && unit.startsWith("$") ? unit.slice(1).replace(/^\//, "per ") : unit;
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
const PAGE_W = 13.33, PAGE_H = 7.5;

function addFooter(slide, label) {
  slide.addText(`${data.account.account_name}  |  Customer Success Plan`, {
    x: 0.5, y: PAGE_H - 0.4, w: 7, h: 0.3, fontFace: FONT_BODY, fontSize: 9, color: GREY, isTextBox: true, margin: 0,
  });
  slide.addText(label || "", {
    x: PAGE_W - 3.5, y: PAGE_H - 0.4, w: 3, h: 0.3, align: "right", fontFace: FONT_BODY, fontSize: 9, color: GREY, isTextBox: true, margin: 0,
  });
}

function sectionTitle(slide, kicker, title) {
  slide.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.35, w: 8, h: 0.3, fontFace: FONT_BODY, fontSize: 12, bold: true, color: SLATE, charSpacing: 2, isTextBox: true, margin: 0,
  });
  slide.addText(title, {
    x: 0.6, y: 0.62, w: 11.5, h: 0.7, fontFace: FONT_HEAD, fontSize: 30, bold: true, color: NAVY, isTextBox: true, margin: 0,
  });
}

// =========================================================
// SLIDE 1 — Title
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY_DEEP };
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: PAGE_W, h: PAGE_H, fill: { color: NAVY_DEEP } });
  slide.addShape(pres.ShapeType.ellipse, { x: 9.7, y: -2.2, w: 6, h: 6, fill: { color: NAVY, transparency: 40 }, line: { type: "none" } });
  slide.addShape(pres.ShapeType.ellipse, { x: 11.2, y: 3.6, w: 4, h: 4, fill: { color: ICE, transparency: 88 }, line: { type: "none" } });

  slide.addText("CUSTOMER SUCCESS PLAN", {
    x: 0.9, y: 2.15, w: 9, h: 0.4, fontFace: FONT_BODY, fontSize: 14, bold: true, color: ICE, charSpacing: 3, isTextBox: true, margin: 0,
  });
  slide.addText(data.account.account_name, {
    x: 0.9, y: 2.55, w: 10.5, h: 1.1, fontFace: FONT_HEAD, fontSize: 44, bold: true, color: WHITE, isTextBox: true, margin: 0,
  });
  slide.addText(`${data.account.industry}  \u2022  ${data.account.segment} Segment  \u2022  CSM: ${data.account.csm_owner}`, {
    x: 0.9, y: 3.65, w: 10, h: 0.4, fontFace: FONT_BODY, fontSize: 15, color: ICE, isTextBox: true, margin: 0,
  });

  const stats = [
    ["ARR", data.account.arr_fmt],
    ["Health Score", `${data.account.health_score} / 100`],
    ["Renewal Date", data.account.renewal_date_fmt],
  ];
  let sx = 0.9;
  stats.forEach(([label, val]) => {
    slide.addShape(pres.ShapeType.rect, { x: sx, y: 4.7, w: 3.3, h: 1.3, fill: { color: WHITE, transparency: 92 }, line: { color: ICE, width: 0.75 } });
    slide.addText(val, { x: sx + 0.2, y: 4.85, w: 2.9, h: 0.6, fontFace: FONT_HEAD, fontSize: 24, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    slide.addText(label.toUpperCase(), { x: sx + 0.2, y: 5.45, w: 2.9, h: 0.4, fontFace: FONT_BODY, fontSize: 10.5, color: ICE, charSpacing: 1, isTextBox: true, margin: 0 });
    sx += 3.55;
  });

  slide.addText(`Prepared ${data.generated}`, {
    x: 0.9, y: 6.85, w: 6, h: 0.35, fontFace: FONT_BODY, fontSize: 10.5, italic: true, color: "8C97C4", isTextBox: true, margin: 0,
  });
}

// =========================================================
// SLIDE 2 — Executive Summary
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Overview", "Executive Summary");

  // Health gauge (semi-donut via pie workaround -> use simple ring text callout instead)
  const hc = healthColor(data.account.health_score);
  slide.addShape(pres.ShapeType.ellipse, { x: 0.7, y: 1.7, w: 2.2, h: 2.2, fill: { color: LIGHT_BG }, line: { color: hc, width: 6 } });
  slide.addText(String(data.account.health_score), { x: 0.7, y: 2.15, w: 2.2, h: 0.8, align: "center", fontFace: FONT_HEAD, fontSize: 36, bold: true, color: hc, isTextBox: true, margin: 0 });
  slide.addText("HEALTH SCORE", { x: 0.7, y: 2.85, w: 2.2, h: 0.35, align: "center", fontFace: FONT_BODY, fontSize: 10, bold: true, color: GREY, charSpacing: 1, isTextBox: true, margin: 0 });

  const goalsText = data.goals.map(g => ({ text: g.description, options: { bullet: { code: "2022" }, color: "222222", breakLine: true, paraSpaceAfter: 8 } }));
  slide.addText("Primary Business Goals", { x: 3.3, y: 1.6, w: 4.6, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  slide.addText(goalsText, { x: 3.3, y: 2.0, w: 4.6, h: 2.2, fontFace: FONT_BODY, fontSize: 13, valign: "top", isTextBox: true, margin: 0 });

  const nextMilestone = data.milestones.find(m => m.status !== "Complete") || data.milestones[0];
  slide.addText("Next Key Milestone", { x: 8.3, y: 1.6, w: 4.3, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  slide.addShape(pres.ShapeType.rect, { x: 8.3, y: 2.0, w: 4.3, h: 1.35, fill: { color: LIGHT_BG }, line: { type: "none" } });
  slide.addText(nextMilestone.milestone, { x: 8.5, y: 2.12, w: 3.9, h: 0.7, fontFace: FONT_BODY, fontSize: 13, bold: true, color: "222222", isTextBox: true, margin: 0 });
  slide.addText(`Target: ${nextMilestone.target_date_fmt}  \u2022  ${nextMilestone.status}`, { x: 8.5, y: 2.75, w: 3.9, h: 0.4, fontFace: FONT_BODY, fontSize: 10.5, color: statusColor(nextMilestone.status), bold: true, isTextBox: true, margin: 0 });

  // bottom row: quick stat cards, one per customer-value metric
  let bx = 0.6;
  data.metrics.forEach((m) => {
    slide.addShape(pres.ShapeType.roundRect, { x: bx, y: 4.5, w: 3.9, h: 1.9, rectRadius: 0.08, fill: { color: LIGHT_BG }, line: { color: "E2E6F0", width: 1 } });
    slide.addText(m.category.toUpperCase(), { x: bx + 0.25, y: 4.68, w: 3.4, h: 0.3, fontFace: FONT_BODY, fontSize: 10.5, bold: true, color: SLATE, charSpacing: 1, isTextBox: true, margin: 0 });
    slide.addText(`${formatMetricValue(m.current_value, m.unit)} / ${formatMetricValue(m.target_value, m.unit)}`, { x: bx + 0.25, y: 5.0, w: 3.4, h: 0.6, fontFace: FONT_HEAD, fontSize: 26, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    slide.addText(`${m.metric_name} (${formatUnitLabel(m.unit)})`, { x: bx + 0.25, y: 5.62, w: 3.5, h: 0.65, fontFace: FONT_BODY, fontSize: 10.5, color: GREY, isTextBox: true, margin: 0 });
    bx += 4.15;
  });

  addFooter(slide, "1 / 8");
}

// =========================================================
// SLIDE 3 — Customer Profile
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 1", "Customer Profile");

  slide.addText("Key Stakeholders", { x: 0.6, y: 1.55, w: 6, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  let sy = 2.0;
  data.stakeholders.forEach((s) => {
    slide.addShape(pres.ShapeType.ellipse, { x: 0.6, y: sy, w: 0.5, h: 0.5, fill: { color: ICE }, line: { type: "none" } });
    slide.addText(s.name.split(" ").map(p => p[0]).join(""), { x: 0.6, y: sy, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 12, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    slide.addText(`${s.name}  \u2014  ${s.title}`, { x: 1.25, y: sy + 0.0, w: 5.5, h: 0.3, fontFace: FONT_BODY, fontSize: 12.5, bold: true, color: "222222", isTextBox: true, margin: 0 });
    slide.addText(s.role, { x: 1.25, y: sy + 0.27, w: 5.5, h: 0.28, fontFace: FONT_BODY, fontSize: 10.5, color: GREY, isTextBox: true, margin: 0 });
    slide.addShape(pres.ShapeType.roundRect, { x: 5.55, y: sy + 0.06, w: 1.2, h: 0.36, rectRadius: 0.06, fill: { color: sentimentColor(s.sentiment) }, line: { type: "none" } });
    slide.addText(s.sentiment, { x: 5.55, y: sy + 0.06, w: 1.2, h: 0.36, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    sy += 0.95;
  });

  slide.addShape(pres.ShapeType.rect, { x: 7.3, y: 1.55, w: 0.02, h: 5.15, fill: { color: "E2E6F0" }, line: { type: "none" } });

  slide.addText("Key Use Cases", { x: 7.7, y: 1.55, w: 4.9, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  const ucText = data.use_cases.map(u => ({ text: `${u.description}  (${u.priority} priority)`, options: { bullet: { code: "2022" }, breakLine: true, paraSpaceAfter: 10, color: "222222" } }));
  slide.addText(ucText, { x: 7.7, y: 1.95, w: 4.9, h: 1.8, fontFace: FONT_BODY, fontSize: 12.5, valign: "top", isTextBox: true, margin: 0 });

  slide.addText("Business Goals", { x: 7.7, y: 3.9, w: 4.9, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  const goalText = data.goals.map(g => ({ text: `${g.description}  (${g.priority} priority)`, options: { bullet: { code: "2022" }, breakLine: true, paraSpaceAfter: 10, color: "222222" } }));
  slide.addText(goalText, { x: 7.7, y: 4.3, w: 4.9, h: 2.2, fontFace: FONT_BODY, fontSize: 12.5, valign: "top", isTextBox: true, margin: 0 });

  addFooter(slide, "2 / 8");
}

// =========================================================
// SLIDE 4 — Current State Assessment
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 2", "Current State Assessment");

  slide.addText("Feature Adoption", { x: 0.6, y: 1.55, w: 6.3, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });

  const chartData = [{
    name: "Adoption",
    labels: data.usage.map(u => u.feature_module),
    values: data.usage.map(u => Math.round(u.adoption_pct * 100)),
  }];
  slide.addChart(pres.ChartType.bar, chartData, {
    x: 0.6, y: 2.0, w: 6.3, h: 3.0,
    barDir: "bar",
    chartColors: [NAVY],
    showTitle: false,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: "222222",
    dataLabelFontSize: 11,
    valAxisMaxVal: 100,
    valAxisTitle: "Adoption %",
    showValAxisTitle: true,
    catAxisLabelColor: "444444",
    valAxisLabelColor: "444444",
    valGridLine: { color: "E9ECF5", size: 1 },
    catGridLine: { style: "none" },
    showLegend: false,
  });

  slide.addShape(pres.ShapeType.rect, { x: 7.3, y: 1.55, w: 0.02, h: 5.15, fill: { color: "E2E6F0" }, line: { type: "none" } });

  slide.addText("Known Challenges", { x: 7.7, y: 1.55, w: 4.9, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  let cy = 2.0;
  data.usage.filter(u => u.known_challenge && u.known_challenge !== "None reported").forEach(u => {
    slide.addShape(pres.ShapeType.roundRect, { x: 7.7, y: cy, w: 4.9, h: 1.15, rectRadius: 0.06, fill: { color: LIGHT_BG }, line: { type: "none" } });
    slide.addShape(pres.ShapeType.roundRect, { x: 7.85, y: cy + 0.13, w: 0.9, h: 0.32, rectRadius: 0.05, fill: { color: severityColor(u.severity) }, line: { type: "none" } });
    slide.addText(u.severity, { x: 7.85, y: cy + 0.13, w: 0.9, h: 0.32, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 9, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    slide.addText(u.feature_module, { x: 8.85, y: cy + 0.1, w: 3.6, h: 0.32, fontFace: FONT_BODY, fontSize: 11.5, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    slide.addText(u.known_challenge, { x: 7.85, y: cy + 0.52, w: 4.6, h: 0.55, fontFace: FONT_BODY, fontSize: 10.5, color: "333333", isTextBox: true, margin: 0 });
    cy += 1.32;
  });

  addFooter(slide, "3 / 8");
}

// =========================================================
// SLIDE 5 — Milestones & Timeline
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 3", "Milestones & Timeline");

  const n = data.milestones.length;
  const trackY = 3.3;
  const startX = 1.75, endX = 11.6;
  const step = (endX - startX) / (n - 1 || 1);

  slide.addShape(pres.ShapeType.line, { x: startX, y: trackY, w: endX - startX, h: 0, line: { color: "D7DCEC", width: 3 } });

  data.milestones.forEach((m, i) => {
    const cx = startX + step * i;
    const col = statusColor(m.status);
    slide.addShape(pres.ShapeType.ellipse, { x: cx - 0.14, y: trackY - 0.14, w: 0.28, h: 0.28, fill: { color: col }, line: { color: WHITE, width: 2 } });

    const above = i % 2 === 0;
    const boxY = above ? trackY - 1.9 : trackY + 0.35;
    slide.addShape(pres.ShapeType.line, { x: cx, y: above ? boxY + 1.4 : trackY + 0.14, w: 0, h: above ? (trackY - 0.14 - (boxY + 1.4)) : 0.2, line: { color: "D7DCEC", width: 1.5, dashType: "dash" } });

    slide.addShape(pres.ShapeType.roundRect, { x: cx - 1.15, y: boxY, w: 2.3, h: 1.35, rectRadius: 0.06, fill: { color: LIGHT_BG }, line: { color: "E2E6F0", width: 1 } });
    slide.addText(m.stage.toUpperCase(), { x: cx - 1.0, y: boxY + 0.1, w: 2.0, h: 0.25, fontFace: FONT_BODY, fontSize: 9, bold: true, color: SLATE, charSpacing: 1, isTextBox: true, margin: 0 });
    slide.addText(m.milestone, { x: cx - 1.0, y: boxY + 0.35, w: 2.0, h: 0.6, fontFace: FONT_BODY, fontSize: 10, bold: true, color: "222222", isTextBox: true, margin: 0 });
    slide.addText(m.target_date_fmt, { x: cx - 1.0, y: boxY + 0.95, w: 2.0, h: 0.25, fontFace: FONT_BODY, fontSize: 9, color: col, bold: true, isTextBox: true, margin: 0 });
  });

  // legend
  const legend = [["Complete", GREEN], ["In Progress / Scheduled", AMBER], ["Not Started", GREY], ["At Risk", RED]];
  let lx = 0.6;
  legend.forEach(([label, col]) => {
    slide.addShape(pres.ShapeType.ellipse, { x: lx, y: 6.75, w: 0.18, h: 0.18, fill: { color: col }, line: { type: "none" } });
    slide.addText(label, { x: lx + 0.25, y: 6.68, w: 2.2, h: 0.3, fontFace: FONT_BODY, fontSize: 10, color: GREY, isTextBox: true, margin: 0 });
    lx += 2.7;
  });

  addFooter(slide, "4 / 8");
}

// =========================================================
// SLIDE 6 — Success Metrics
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 4", "Success Metrics");

  const catMeta = {
    "Time Saved": "Hours or cycle time given back to the customer",
    "Cost Reduction": "Spend avoided or reduced for the customer",
    "Process Simplified": "Steps, tools, or manual work removed",
    "Revenue Growth": "Sales or revenue enabled for the customer",
  };
  const cardW = Math.min(4.0, 11.9 / data.metrics.length);
  let cx = 0.6;
  data.metrics.forEach((m) => {
    const higherIsBetter = m.higher_is_better !== 0;
    const rawPct = higherIsBetter ? (m.current_value / m.target_value) * 100 : (m.target_value / m.current_value) * 100;
    const pct = Math.max(0, Math.min(100, Math.round(rawPct)));
    slide.addShape(pres.ShapeType.roundRect, { x: cx, y: 1.65, w: cardW - 0.25, h: 3.0, rectRadius: 0.08, fill: { color: LIGHT_BG }, line: { color: "E2E6F0", width: 1 } });
    slide.addText(m.category.toUpperCase(), { x: cx + 0.3, y: 1.9, w: cardW - 0.85, h: 0.3, fontFace: FONT_BODY, fontSize: 11, bold: true, color: SLATE, charSpacing: 1, isTextBox: true, margin: 0 });
    slide.addText(catMeta[m.category] || "Outcome delivered for the customer", { x: cx + 0.3, y: 2.2, w: cardW - 0.85, h: 0.35, fontFace: FONT_BODY, fontSize: 10, italic: true, color: GREY, isTextBox: true, margin: 0 });
    slide.addText(m.metric_name, { x: cx + 0.3, y: 2.5, w: cardW - 0.85, h: 0.4, fontFace: FONT_BODY, fontSize: 12, bold: true, color: "222222", isTextBox: true, margin: 0 });
    slide.addText(formatMetricValue(m.current_value, m.unit), { x: cx + 0.3, y: 2.85, w: cardW - 0.85, h: 0.65, fontFace: FONT_HEAD, fontSize: 34, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    slide.addText(`of ${formatMetricValue(m.target_value, m.unit)} ${formatUnitLabel(m.unit)} target`, { x: cx + 0.3, y: 3.48, w: cardW - 0.85, h: 0.3, fontFace: FONT_BODY, fontSize: 10.5, color: GREY, isTextBox: true, margin: 0 });
    // progress bar
    slide.addShape(pres.ShapeType.roundRect, { x: cx + 0.3, y: 3.88, w: cardW - 0.85, h: 0.22, rectRadius: 0.03, fill: { color: "E2E6F0" }, line: { type: "none" } });
    slide.addShape(pres.ShapeType.roundRect, { x: cx + 0.3, y: 3.88, w: Math.max(0.15, (cardW - 0.85) * pct / 100), h: 0.22, rectRadius: 0.03, fill: { color: pct >= 90 ? GREEN : pct >= 60 ? AMBER : RED }, line: { type: "none" } });
    slide.addText(`${pct}% of target`, { x: cx + 0.3, y: 4.15, w: cardW - 0.85, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, color: GREY, isTextBox: true, margin: 0 });
    cx += cardW;
  });

  slide.addText("Metrics are tracked in the terms that matter to the customer: time saved, cost avoided, revenue enabled, or complexity removed from their operations.", {
    x: 0.6, y: 4.95, w: 12.1, h: 0.5, fontFace: FONT_BODY, fontSize: 10.5, italic: true, color: GREY, isTextBox: true, margin: 0,
  });

  addFooter(slide, "5 / 8");
}

// =========================================================
// SLIDE 7 — Actions & Responsibilities
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 5", "Actions & Responsibilities");

  const cols = [
    { owner: "CS Team", x: 0.6, color: NAVY },
    { owner: "Customer", x: 6.9, color: SLATE },
  ];
  cols.forEach(col => {
    slide.addShape(pres.ShapeType.roundRect, { x: col.x, y: 1.55, w: 5.8, h: 0.55, rectRadius: 0.06, fill: { color: col.color }, line: { type: "none" } });
    slide.addText(`${col.owner} Owns`, { x: col.x, y: 1.55, w: 5.8, h: 0.55, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 14, bold: true, color: WHITE, isTextBox: true, margin: 0 });

    let ay = 2.3;
    data.actions.filter(a => a.owner === col.owner).forEach(a => {
      slide.addShape(pres.ShapeType.roundRect, { x: col.x, y: ay, w: 5.8, h: 1.0, rectRadius: 0.06, fill: { color: LIGHT_BG }, line: { color: "E2E6F0", width: 1 } });
      slide.addText(a.action, { x: col.x + 0.25, y: ay + 0.1, w: 5.3, h: 0.55, fontFace: FONT_BODY, fontSize: 11.5, bold: true, color: "222222", isTextBox: true, margin: 0 });
      slide.addText(`Due ${a.due_date_fmt}`, { x: col.x + 0.25, y: ay + 0.65, w: 3, h: 0.3, fontFace: FONT_BODY, fontSize: 10, color: GREY, isTextBox: true, margin: 0 });
      slide.addShape(pres.ShapeType.roundRect, { x: col.x + 4.3, y: ay + 0.62, w: 1.3, h: 0.32, rectRadius: 0.05, fill: { color: statusColor(a.status) }, line: { type: "none" } });
      slide.addText(a.status, { x: col.x + 4.3, y: ay + 0.62, w: 1.3, h: 0.32, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 9, bold: true, color: WHITE, isTextBox: true, margin: 0 });
      ay += 1.18;
    });
  });

  addFooter(slide, "6 / 8");
}

// =========================================================
// SLIDE 8 — Risk Management
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  sectionTitle(slide, "Section 6", "Risk Management");

  const headers = ["Risk", "Likelihood", "Impact", "Mitigation"];
  const rowsTable = data.risks.map(r => [r.risk, r.likelihood, r.impact, r.mitigation]);

  const tableRows = [
    headers.map(h => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontFace: FONT_BODY, fontSize: 11 } })),
  ];
  rowsTable.forEach((r, i) => {
    tableRows.push([
      { text: r[0], options: { fontFace: FONT_BODY, fontSize: 10.5, color: "222222", fill: { color: i % 2 ? LIGHT_BG : WHITE } } },
      { text: r[1], options: { fontFace: FONT_BODY, fontSize: 10.5, color: severityColor(r[1]), bold: true, align: "center", fill: { color: i % 2 ? LIGHT_BG : WHITE } } },
      { text: r[2], options: { fontFace: FONT_BODY, fontSize: 10.5, color: severityColor(r[2]), bold: true, align: "center", fill: { color: i % 2 ? LIGHT_BG : WHITE } } },
      { text: r[3], options: { fontFace: FONT_BODY, fontSize: 10.5, color: "222222", fill: { color: i % 2 ? LIGHT_BG : WHITE } } },
    ]);
  });

  slide.addTable(tableRows, {
    x: 0.6, y: 1.65, w: 12.1,
    colW: [4.6, 1.5, 1.5, 4.5],
    border: { type: "solid", color: "E2E6F0", pt: 0.75 },
    autoPage: false,
    valign: "middle",
    rowH: 0.85,
  });

  if (data.risks.length === 0) {
    slide.addText("No open risks identified for this account at this time.", { x: 0.6, y: 2.0, w: 10, h: 0.5, fontFace: FONT_BODY, fontSize: 13, color: GREY, isTextBox: true, margin: 0 });
  }

  addFooter(slide, "7 / 8");
}

// =========================================================
// SLIDE 9 — Closing / Next Steps
// =========================================================
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY_DEEP };
  slide.addShape(pres.ShapeType.ellipse, { x: -2, y: 4.5, w: 6, h: 6, fill: { color: NAVY, transparency: 40 }, line: { type: "none" } });

  slide.addText("NEXT STEPS", { x: 0.9, y: 1.3, w: 8, h: 0.4, fontFace: FONT_BODY, fontSize: 14, bold: true, color: ICE, charSpacing: 3, isTextBox: true, margin: 0 });
  slide.addText("Keeping the plan on track", { x: 0.9, y: 1.7, w: 10, h: 0.8, fontFace: FONT_HEAD, fontSize: 32, bold: true, color: WHITE, isTextBox: true, margin: 0 });

  const openActions = data.actions.filter(a => a.status !== "Complete").slice(0, 4);
  let ny = 2.8;
  openActions.forEach((a, i) => {
    slide.addShape(pres.ShapeType.ellipse, { x: 0.9, y: ny, w: 0.5, h: 0.5, fill: { color: ICE }, line: { type: "none" } });
    slide.addText(String(i + 1), { x: 0.9, y: ny, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 14, bold: true, color: NAVY_DEEP, isTextBox: true, margin: 0 });
    slide.addText(a.action, { x: 1.6, y: ny + 0.02, w: 10.5, h: 0.35, fontFace: FONT_BODY, fontSize: 14, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    slide.addText(`${a.owner}  \u2022  Due ${a.due_date_fmt}`, { x: 1.6, y: ny + 0.35, w: 10.5, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: "9AA6D6", isTextBox: true, margin: 0 });
    ny += 0.85;
  });

  slide.addText(`Prepared by ${data.account.csm_owner}  \u2022  ${data.generated}`, {
    x: 0.9, y: 6.85, w: 8, h: 0.35, fontFace: FONT_BODY, fontSize: 10.5, italic: true, color: "8C97C4", isTextBox: true, margin: 0,
  });
}

const outPath = `/home/claude/success_plan/output/success_plan_${accountId}.pptx`;
pres.writeFile({ fileName: outPath }).then(() => console.log("saved", outPath));
