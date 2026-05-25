// api/generate-pptx.js
// Vercel serverless function — builds a student classroom PowerPoint
// Uses pptxgenjs only (no sharp/react-icons for bundle compatibility)

const pptxgen = require("pptxgenjs");

// ── Color palette ─────────────────────────────────────────────
const C = {
  primary:   "0D6E8A",
  secondary: "F5A623",
  accent:    "E8344A",
  dark:      "0A3D50",
  light:     "E8F6FA",
  white:     "FFFFFF",
  offwhite:  "F8FBFC",
  textDark:  "1A2E35",
  textMid:   "4A6670",
  textLight: "8AAAB5",
  intentions:"3B5BA5",
  learning:  "1A7A4A",
  assess:    "C06820",
  reflect:   "7B3FA0",
  hook:      "E8344A",
  practice:  "C06820",
  apply:     "7B3FA0",
};

function safe(str, max = 120) {
  if (!str) return "";
  return String(str).substring(0, max);
}

function safeArr(arr, max = 6) {
  if (!Array.isArray(arr)) return [];
  return arr.slice(0, max);
}

// ── Slide builders ────────────────────────────────────────────

function addTitleSlide(pres, slide, meta) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.OVAL, {
    x: 7.2, y: -1.2, w: 4.2, h: 4.2,
    fill: { color: C.primary, transparency: 50 },
    line: { color: C.primary, transparency: 50 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: -1, y: 3.5, w: 2.8, h: 2.8,
    fill: { color: C.secondary, transparency: 62 },
    line: { color: C.secondary, transparency: 62 }
  });

  // Gold left bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: C.secondary }, line: { color: C.secondary }
  });

  // Subject badge
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 0.42, w: 2.8, h: 0.38,
    fill: { color: C.secondary }, line: { color: C.secondary }, rectRadius: 0.08
  });
  s.addText(safe(meta.subject || "Grade 8").toUpperCase(), {
    x: 0.45, y: 0.42, w: 2.8, h: 0.38,
    fontSize: 10, bold: true, color: C.dark,
    align: "center", valign: "middle", margin: 0, charSpacing: 1
  });

  // Main title
  s.addText(safe(slide.title, 80), {
    x: 0.45, y: 1.0, w: 8.8, h: 2.0,
    fontSize: 34, bold: true, color: C.white,
    fontFace: "Trebuchet MS", charSpacing: -0.5,
    align: "left", valign: "middle"
  });

  s.addText(safe(slide.subtitle || `${meta.term} · ${meta.week}`), {
    x: 0.45, y: 3.05, w: 8.8, h: 0.4,
    fontSize: 13, color: C.secondary, bold: true, fontFace: "Calibri"
  });

  s.addText(`${safe(meta.teacher)}  ·  ${safe(meta.section)}`, {
    x: 0.45, y: 3.5, w: 8.8, h: 0.35,
    fontSize: 12, color: C.textLight, fontFace: "Calibri"
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.15, w: 10, h: 0.475,
    fill: { color: C.primary }, line: { color: C.primary }
  });
  s.addText("Department of Education  ·  MATATAG Curriculum  ·  SY 2026–2027", {
    x: 0.3, y: 5.15, w: 9.4, h: 0.475,
    fontSize: 9, color: C.white, align: "center", valign: "middle", margin: 0
  });
}

function addObjectivesSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.08,
    fill: { color: C.intentions }, line: { color: C.intentions }
  });
  s.addText("🎯  " + safe(slide.title, 60), {
    x: 0.4, y: 0, w: 9.2, h: 1.08,
    fontSize: 22, bold: true, color: C.white,
    valign: "middle", fontFace: "Trebuchet MS", margin: 0
  });

  const objectives = safeArr(slide.objectives, 4);
  const colors = [C.intentions, C.learning, C.assess, C.reflect];
  const yStart = 1.25;
  const boxH = 0.86;
  const gap  = 0.13;

  for (let i = 0; i < objectives.length; i++) {
    const y = yStart + i * (boxH + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: boxH,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.07 },
      line: { color: "E0EAED", width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.16, h: boxH,
      fill: { color: colors[i % colors.length] },
      line: { color: colors[i % colors.length] }
    });
    s.addText(safe(objectives[i], 160), {
      x: 0.72, y: y + 0.06, w: 8.7, h: boxH - 0.12,
      fontSize: 13, color: C.textDark, valign: "middle", fontFace: "Calibri"
    });
  }

  addTeacherNote(s, pres, slide.teacherNote);
}

function addHookSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.OVAL, {
    x: -1.2, y: -1.2, w: 5.5, h: 5.5,
    fill: { color: C.hook, transparency: 78 },
    line: { color: C.hook, transparency: 78 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 6.8, y: 2.2, w: 4.2, h: 4.2,
    fill: { color: C.secondary, transparency: 82 },
    line: { color: C.secondary, transparency: 82 }
  });

  s.addText("🤔  " + safe(slide.title, 50), {
    x: 0.4, y: 0.12, w: 9.2, h: 0.75,
    fontSize: 20, bold: true, color: C.secondary,
    valign: "middle", fontFace: "Trebuchet MS"
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 2.4,
    fill: { color: C.hook, transparency: 87 },
    line: { color: C.hook, width: 1.5 }, rectRadius: 0.14
  });
  s.addText(safe(slide.hookQuestion, 220), {
    x: 0.6, y: 1.1, w: 8.8, h: 2.3,
    fontSize: 16, color: C.white,
    valign: "middle", fontFace: "Calibri", align: "left"
  });

  s.addText("Activity: " + safe(slide.activity, 150), {
    x: 0.4, y: 3.6, w: 9.2, h: 0.65,
    fontSize: 12, color: C.textLight, italic: true,
    fontFace: "Calibri", valign: "middle"
  });

  addTeacherNote(s, pres, slide.teacherNote, true);
}

function addKeyTermsSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: C.primary }, line: { color: C.primary }
  });
  s.addText("🔑  " + safe(slide.title, 60), {
    x: 0.4, y: 0, w: 9.2, h: 1.05,
    fontSize: 22, bold: true, color: C.white,
    valign: "middle", fontFace: "Trebuchet MS", margin: 0
  });

  const terms = safeArr(slide.terms, 4);
  const cols  = terms.length <= 2 ? 1 : 2;
  const perCol = Math.ceil(terms.length / cols);
  const colW  = cols === 1 ? 9.2 : 4.45;
  const termColors = [C.primary, C.intentions, C.learning, C.assess];

  for (let i = 0; i < terms.length; i++) {
    const col = Math.floor(i / perCol);
    const row = i % perCol;
    const x   = 0.4 + col * (colW + 0.3);
    const y   = 1.18 + row * 1.48;
    const tc  = termColors[i % termColors.length];

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: colW, h: 1.32,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.08 },
      line: { color: "E0EAED", width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: colW, h: 0.38,
      fill: { color: tc }, line: { color: tc }
    });
    s.addText(safe(terms[i].term, 40).toUpperCase(), {
      x: x + 0.14, y, w: colW - 0.2, h: 0.38,
      fontSize: 11, bold: true, color: C.white,
      valign: "middle", fontFace: "Trebuchet MS",
      margin: 0, charSpacing: 1
    });
    s.addText(safe(terms[i].definition, 180), {
      x: x + 0.14, y: y + 0.42, w: colW - 0.24, h: 0.84,
      fontSize: 12, color: C.textDark, valign: "top", fontFace: "Calibri"
    });
  }

  addTeacherNote(s, pres, slide.teacherNote);
}

function addContentSlide(pres, slide, accentColor, emoji) {
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.2, h: 5.625,
    fill: { color: accentColor }, line: { color: accentColor }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 0, w: 9.8, h: 1.0,
    fill: { color: C.light }, line: { color: C.light }
  });
  s.addText(`${emoji}  ${safe(slide.title, 70)}`, {
    x: 0.4, y: 0, w: 9.3, h: 1.0,
    fontSize: 22, bold: true, color: accentColor,
    valign: "middle", fontFace: "Trebuchet MS", margin: 0
  });

  const points = safeArr(slide.points || [], 5);
  const bH    = points.length > 3 ? 0.7 : 0.85;
  const yStart = 1.1;

  for (let i = 0; i < points.length; i++) {
    const y = yStart + i * (bH + 0.08);
    s.addShape(pres.shapes.OVAL, {
      x: 0.42, y: y + bH / 2 - 0.14, w: 0.28, h: 0.28,
      fill: { color: accentColor }, line: { color: accentColor }
    });
    s.addText(safe(points[i], 200), {
      x: 0.86, y, w: 8.8, h: bH,
      fontSize: bH > 0.75 ? 14 : 12.5, color: C.textDark,
      valign: "middle", fontFace: "Calibri"
    });
  }

  if (slide.example) {
    const exY = yStart + points.length * (bH + 0.08) + 0.06;
    const remH = 5.08 - exY;
    if (remH > 0.48) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.38, y: exY, w: 9.24, h: Math.min(remH, 0.88),
        fill: { color: accentColor, transparency: 90 },
        line: { color: accentColor, width: 1 }, rectRadius: 0.1
      });
      s.addText("💡  Example: " + safe(slide.example, 170), {
        x: 0.55, y: exY + 0.05, w: 9.0, h: Math.min(remH, 0.88) - 0.1,
        fontSize: 11.5, color: accentColor, italic: true,
        valign: "middle", fontFace: "Calibri"
      });
    }
  }

  addTeacherNote(s, pres, slide.teacherNote);
}

function addGuidedPracticeSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: C.practice }, line: { color: C.practice }
  });
  s.addText("👥  " + safe(slide.title, 60), {
    x: 0.4, y: 0, w: 9.2, h: 1.05,
    fontSize: 22, bold: true, color: C.white,
    valign: "middle", fontFace: "Trebuchet MS", margin: 0
  });

  s.addText(safe(slide.instructions, 200), {
    x: 0.4, y: 1.14, w: 9.2, h: 0.58,
    fontSize: 13, color: C.textDark, italic: true,
    valign: "middle", fontFace: "Calibri"
  });

  const steps = safeArr(slide.steps || [], 5);
  for (let i = 0; i < steps.length; i++) {
    const y = 1.86 + i * 0.7;
    s.addShape(pres.shapes.OVAL, {
      x: 0.38, y: y + 0.04, w: 0.42, h: 0.42,
      fill: { color: C.practice }, line: { color: C.practice }
    });
    s.addText(String(i + 1), {
      x: 0.38, y: y + 0.04, w: 0.42, h: 0.42,
      fontSize: 13, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(safe(steps[i], 200), {
      x: 0.98, y, w: 8.65, h: 0.6,
      fontSize: 13, color: C.textDark, valign: "middle", fontFace: "Calibri"
    });
  }

  if (slide.guideQuestions && slide.guideQuestions.length) {
    const qY  = 1.86 + steps.length * 0.7 + 0.1;
    const remH = 5.06 - qY;
    if (remH > 0.4) {
      const qs = safeArr(slide.guideQuestions, 2).map(q => "❓ " + q).join("   |   ");
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: qY, w: 10, h: Math.min(remH, 0.88),
        fill: { color: C.practice, transparency: 90 },
        line: { color: C.practice, transparency: 90 }
      });
      s.addText(safe(qs, 220), {
        x: 0.4, y: qY + 0.05, w: 9.2, h: Math.min(remH, 0.88) - 0.1,
        fontSize: 11, color: C.practice, italic: true,
        valign: "middle", fontFace: "Calibri"
      });
    }
  }

  addTeacherNote(s, pres, slide.teacherNote);
}

function addApplicationSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.OVAL, {
    x: 6.2, y: -1.2, w: 5.2, h: 5.2,
    fill: { color: C.apply, transparency: 78 },
    line: { color: C.apply, transparency: 78 }
  });

  s.addText("✏️  " + safe(slide.title, 50), {
    x: 0.4, y: 0.1, w: 9.2, h: 0.72,
    fontSize: 24, bold: true, color: C.secondary,
    valign: "middle", fontFace: "Trebuchet MS"
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 0.98, w: 3.5, h: 0.42,
    fill: { color: C.apply }, line: { color: C.apply }, rectRadius: 0.08
  });
  s.addText(safe(slide.taskTitle || "Activity", 40).toUpperCase(), {
    x: 0.4, y: 0.98, w: 3.5, h: 0.42,
    fontSize: 11, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0, charSpacing: 1
  });

  if (slide.timeGiven) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 4.1, y: 0.98, w: 1.9, h: 0.42,
      fill: { color: C.secondary }, line: { color: C.secondary }, rectRadius: 0.08
    });
    s.addText("⏱  " + safe(slide.timeGiven, 18), {
      x: 4.1, y: 0.98, w: 1.9, h: 0.42,
      fontSize: 11, bold: true, color: C.dark,
      align: "center", valign: "middle", margin: 0
    });
  }

  s.addText(safe(slide.instructions, 220), {
    x: 0.4, y: 1.52, w: 9.2, h: 0.54,
    fontSize: 13, color: C.textLight, italic: true,
    fontFace: "Calibri", valign: "middle"
  });

  const steps = safeArr(slide.steps || [], 4);
  for (let i = 0; i < steps.length; i++) {
    const y = 2.16 + i * 0.68;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y: y + 0.06, w: 0.05, h: 0.46,
      fill: { color: C.secondary }, line: { color: C.secondary }
    });
    s.addText(safe(steps[i], 200), {
      x: 0.64, y, w: 8.98, h: 0.58,
      fontSize: 13, color: C.white, valign: "middle", fontFace: "Calibri"
    });
  }

  addTeacherNote(s, pres, slide.teacherNote, true);
}

function addSummarySlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.primary };

  s.addShape(pres.shapes.OVAL, {
    x: -1.8, y: -1.2, w: 5.2, h: 5.2,
    fill: { color: C.white, transparency: 92 },
    line: { color: C.white, transparency: 92 }
  });

  s.addText("🏆  " + safe(slide.title, 60), {
    x: 0.4, y: 0.1, w: 9.2, h: 0.72,
    fontSize: 24, bold: true, color: C.white,
    valign: "middle", fontFace: "Trebuchet MS"
  });

  if (slide.bigIdea) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 0.98, w: 9.2, h: 0.72,
      fill: { color: C.secondary }, line: { color: C.secondary }, rectRadius: 0.1
    });
    s.addText("💡  " + safe(slide.bigIdea, 170), {
      x: 0.55, y: 0.98, w: 9.0, h: 0.72,
      fontSize: 14, bold: true, color: C.dark,
      valign: "middle", fontFace: "Trebuchet MS", margin: 0
    });
  }

  const takeaways = safeArr(slide.takeaways || [], 4);
  for (let i = 0; i < takeaways.length; i++) {
    const y = 1.84 + i * 0.8;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.7,
      fill: { color: C.white }, line: { color: "D0E8EF", width: 0.5 }
    });
    s.addText("✓", {
      x: 0.52, y, w: 0.42, h: 0.7,
      fontSize: 14, bold: true, color: C.primary,
      align: "center", valign: "middle"
    });
    s.addText(safe(takeaways[i], 190), {
      x: 1.04, y: y + 0.08, w: 8.3, h: 0.54,
      fontSize: 13, color: C.textDark, valign: "middle", fontFace: "Calibri"
    });
  }

  addTeacherNote(s, pres, slide.teacherNote);
}

function addExitTicketSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.OVAL, {
    x: 5.2, y: 0.1, w: 5.4, h: 5.4,
    fill: { color: C.hook, transparency: 84 },
    line: { color: C.hook, transparency: 84 }
  });

  s.addText("📋  " + safe(slide.title, 60), {
    x: 0.4, y: 0.1, w: 9.2, h: 0.72,
    fontSize: 24, bold: true, color: C.hook,
    valign: "middle", fontFace: "Trebuchet MS"
  });

  // White card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.04, w: 9.2, h: 2.7,
    fill: { color: C.white },
    line: { color: C.hook, width: 1.5 }, rectRadius: 0.14
  });
  s.addText(safe(slide.question, 300), {
    x: 0.64, y: 1.14, w: 8.74, h: 2.5,
    fontSize: 16, bold: true, color: C.textDark,
    valign: "middle", fontFace: "Calibri", align: "left"
  });

  s.addText("📝  " + safe(slide.directions, 170), {
    x: 0.4, y: 3.88, w: 9.2, h: 0.62,
    fontSize: 12, color: C.secondary, italic: true,
    valign: "middle", fontFace: "Calibri"
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.05, w: 10, h: 0.575,
    fill: { color: C.hook }, line: { color: C.hook }
  });
  s.addText("Great work today! 🎉", {
    x: 0, y: 5.05, w: 10, h: 0.575,
    fontSize: 14, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0
  });
}

// ── Teacher note bar ─────────────────────────────────────────
function addTeacherNote(s, pres, note, darkBg = false) {
  if (!note) return;
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.05, w: 10, h: 0.575,
    fill: { color: darkBg ? "F5A623" : "FFF8E7" },
    line: { color: darkBg ? "F5A623" : "FFD966", width: 0.5 }
  });
  s.addText("🧑‍🏫  " + safe(note, 130), {
    x: 0.3, y: 5.05, w: 9.4, h: 0.575,
    fontSize: 9, color: darkBg ? "0A3D50" : "7B5B00",
    valign: "middle", fontFace: "Calibri", italic: true, margin: 0
  });
}

// ── Main generator ─────────────────────────────────────────────
async function generatePPTX(lessonData) {
  const slides = lessonData.slides || [];
  const meta   = lessonData.meta   || {};

  const pres = new pptxgen();
  pres.layout  = "LAYOUT_16x9";
  pres.author  = meta.teacher || "DepEd Teacher";
  pres.title   = meta.topic   || "Lesson Presentation";
  pres.subject = `${meta.subject || "Grade 8"} – ${meta.term || ""}`;

  for (const slide of slides) {
    switch (slide.type) {
      case "title":
        addTitleSlide(pres, slide, meta); break;
      case "objectives":
        addObjectivesSlide(pres, slide); break;
      case "hook":
        addHookSlide(pres, slide); break;
      case "keyterms":
        addKeyTermsSlide(pres, slide); break;
      case "content1":
        addContentSlide(pres, slide, C.learning, "📗"); break;
      case "content2":
        addContentSlide(pres, slide, C.intentions, "🔬"); break;
      case "guidedpractice":
        addGuidedPracticeSlide(pres, slide); break;
      case "application":
        addApplicationSlide(pres, slide); break;
      case "summary":
        addSummarySlide(pres, slide); break;
      case "exitticket":
        addExitTicketSlide(pres, slide); break;
      default:
        addContentSlide(pres, slide, C.primary, "📌");
    }
  }

  // Return as base64 buffer
  const base64 = await pres.write({ outputType: "base64" });
  return Buffer.from(base64, "base64");
}

// ── Vercel handler ─────────────────────────────────────────────
module.exports = async (req, res) => {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const lessonData = req.body;
    if (!lessonData || !lessonData.slides) {
      return res.status(400).json({ error: "Missing lesson data" });
    }

    const buffer   = await generatePPTX(lessonData);
    const safeName = (lessonData.meta?.topic || "Lesson")
      .replace(/[^a-zA-Z0-9\s]/g, "")
      .replace(/\s+/g, "_")
      .substring(0, 40);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="Grade8_${safeName}.pptx"`);
    res.setHeader("Content-Length", buffer.length);
    res.status(200).send(buffer);

  } catch (err) {
    console.error("PPTX error:", err);
    res.status(500).json({ error: "PPTX generation failed: " + err.message });
  }
};
