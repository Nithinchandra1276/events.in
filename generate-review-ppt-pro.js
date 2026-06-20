const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Project Team';
pptx.company = 'Final Review';
pptx.subject = 'AI Event and Service Management Platform';
pptx.title = 'Professional Project Review Deck';
pptx.lang = 'en-US';

const THEME = {
  navy: '0A1F2F',
  cyan: '0F8AA6',
  slate: '1D3444',
  ink: '0D1720',
  text: '1F2937',
  soft: 'F4F8FB',
  white: 'FFFFFF',
  accent: 'F48B3A',
  line: 'D4E2EA',
  muted: '5A6B79',
  success: '0F9D58'
};

const reviewDate = new Date().toLocaleDateString('en-GB', {
  day: '2-digit',
  month: 'long',
  year: 'numeric'
});

function addBase(slide, slideNo, totalSlides) {
  slide.background = { color: THEME.soft };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.78,
    fill: { color: THEME.navy },
    line: { color: THEME.navy }
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 7.23,
    w: 13.33,
    h: 0.27,
    fill: { color: 'E8F0F5' },
    line: { color: 'E8F0F5' }
  });

  slide.addText('AI Event + Service Platform | Final Review 2026', {
    x: 0.45,
    y: 0.22,
    w: 8.6,
    h: 0.3,
    fontSize: 11,
    color: 'D6E9F4',
    bold: true,
    fontFace: 'Calibri'
  });

  slide.addText(`Slide ${slideNo}/${totalSlides}`, {
    x: 11.1,
    y: 7.28,
    w: 1.9,
    h: 0.16,
    fontSize: 9,
    color: '607180',
    align: 'right',
    fontFace: 'Calibri'
  });
}

function addHeadline(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.58,
    y: 0.96,
    w: 12,
    h: 0.46,
    fontSize: 25,
    bold: true,
    color: THEME.ink,
    fontFace: 'Calibri'
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6,
      y: 1.45,
      w: 12,
      h: 0.28,
      fontSize: 13,
      italic: true,
      color: THEME.cyan,
      fontFace: 'Calibri'
    });
  }
}

function addBulletPanel(slide, bullets, options = {}) {
  const x = options.x || 0.72;
  const y = options.y || 1.86;
  const w = options.w || 12;
  const h = options.h || 5.1;

  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    radius: 0.08,
    fill: { color: THEME.white },
    line: { color: THEME.line, pt: 1 }
  });

  const runs = bullets.map((text) => ({
    text,
    options: { bullet: { indent: 16 } }
  }));

  slide.addText(runs, {
    x: x + 0.26,
    y: y + 0.34,
    w: w - 0.5,
    h: h - 0.5,
    fontSize: 20,
    color: THEME.text,
    breakLine: true,
    paraSpaceAfterPt: 13,
    valign: 'top',
    margin: 3,
    fontFace: 'Calibri'
  });
}

function addSectionSlide(slide, label, title, points) {
  slide.background = { color: THEME.navy };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 7.5,
    fill: { color: THEME.navy },
    line: { color: THEME.navy }
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 0.8,
    w: 2.7,
    h: 0.46,
    radius: 0.06,
    fill: { color: THEME.accent },
    line: { color: THEME.accent }
  });

  slide.addText(label, {
    x: 0.96,
    y: 0.93,
    w: 2.3,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: THEME.white,
    fontFace: 'Calibri'
  });

  slide.addText(title, {
    x: 0.85,
    y: 1.6,
    w: 11.6,
    h: 0.95,
    fontSize: 36,
    bold: true,
    color: THEME.white,
    fontFace: 'Calibri'
  });

  const runs = points.map((text) => ({ text, options: { bullet: { indent: 15 } } }));
  slide.addText(runs, {
    x: 1.05,
    y: 3,
    w: 11,
    h: 3,
    fontSize: 18,
    color: 'D9EBF5',
    breakLine: true,
    paraSpaceAfterPt: 14,
    fontFace: 'Calibri'
  });

  slide.addText('Professional Review Deck', {
    x: 0.85,
    y: 7.13,
    w: 5,
    h: 0.2,
    fontSize: 10,
    color: '9FC2D6',
    fontFace: 'Calibri'
  });
}

function addKpiSlide(slide, items) {
  addHeadline(slide, 'Measured Progress Snapshot', 'Live module health and review readiness indicators');

  const baseY = 2.0;
  const cardW = 2.87;
  const cardGap = 0.26;

  items.forEach((item, idx) => {
    const x = 0.74 + idx * (cardW + cardGap);

    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: baseY,
      w: cardW,
      h: 2.3,
      radius: 0.08,
      fill: { color: THEME.white },
      line: { color: THEME.line, pt: 1 }
    });

    slide.addText(item.title, {
      x: x + 0.2,
      y: baseY + 0.28,
      w: cardW - 0.4,
      h: 0.35,
      fontSize: 11,
      bold: true,
      color: THEME.muted,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addText(item.value, {
      x: x + 0.2,
      y: baseY + 0.83,
      w: cardW - 0.4,
      h: 0.78,
      fontSize: 30,
      bold: true,
      color: THEME.cyan,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addText(item.note, {
      x: x + 0.16,
      y: baseY + 1.72,
      w: cardW - 0.32,
      h: 0.45,
      fontSize: 10,
      color: THEME.text,
      align: 'center',
      fontFace: 'Calibri'
    });
  });

  addBulletPanel(slide, [
    'Unified data model in AppShared computes readiness, budget variance, and task health.',
    'Analytics, planner, and options pages consume the same shared metrics for consistency.',
    'Dashboard messaging supports quick faculty review decisions with quantifiable indicators.'
  ], { x: 0.74, y: 4.55, w: 11.84, h: 2.3 });
}

function addArchitectureSlide(slide) {
  addHeadline(slide, 'System Architecture', 'Single-platform flow from user intent to generated output and governance dashboards');

  const nodes = [
    { x: 0.7, y: 2.3, w: 2.45, h: 1.1, title: 'Client Layer', text: 'index.html\noptions.html\nmodule pages' },
    { x: 3.7, y: 2.3, w: 2.45, h: 1.1, title: 'Shared Logic', text: 'app-shared.js\nreadiness model\nlocal data state' },
    { x: 6.7, y: 2.3, w: 2.45, h: 1.1, title: 'Express API', text: 'index.js\n/generate\n/download routes' },
    { x: 9.7, y: 2.3, w: 2.9, h: 1.1, title: 'AI Image Backends', text: 'OpenAI gpt-image-1\nHuggingFace fallback\nlocal sample fallback' }
  ];

  nodes.forEach((node) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: node.x,
      y: node.y,
      w: node.w,
      h: node.h,
      radius: 0.08,
      fill: { color: THEME.white },
      line: { color: THEME.cyan, pt: 1.2 }
    });

    slide.addText(node.title, {
      x: node.x + 0.15,
      y: node.y + 0.13,
      w: node.w - 0.3,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: THEME.ink,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addText(node.text, {
      x: node.x + 0.15,
      y: node.y + 0.4,
      w: node.w - 0.3,
      h: 0.62,
      fontSize: 10,
      color: THEME.text,
      align: 'center',
      valign: 'mid',
      breakLine: true,
      fontFace: 'Calibri'
    });
  });

  const arrows = [
    { x: 3.2, y: 2.8, w: 0.45 },
    { x: 6.2, y: 2.8, w: 0.45 },
    { x: 9.2, y: 2.8, w: 0.45 }
  ];

  arrows.forEach((item) => {
    slide.addShape(pptx.ShapeType.line, {
      x: item.x,
      y: item.y,
      w: item.w,
      h: 0,
      line: { color: THEME.accent, pt: 2, beginArrowType: 'none', endArrowType: 'triangle' }
    });
  });

  addBulletPanel(slide, [
    'Client module handles UX and command capture (voice + manual input).',
    'Server builds enhanced prompts and enforces robust fallback strategy for generation reliability.',
    'Governance modules (Admin + Analytics + Planner) use shared state for traceable review evidence.'
  ], { x: 0.72, y: 4.1, w: 12, h: 2.75 });
}

function addTimelineSlide(slide) {
  addHeadline(slide, 'Implementation Journey', 'From core prototype to review-ready professional platform');

  const phases = [
    { label: 'Phase 1', name: 'Core Build', detail: 'Authentication + module routing + service flow' },
    { label: 'Phase 2', name: 'AI Upgrade', detail: 'Voice assistant + stage prompt engine + image pipeline' },
    { label: 'Phase 3', name: 'Governance', detail: 'Admin, analytics dashboard, planner workspace' },
    { label: 'Phase 4', name: 'Review Prep', detail: 'Checklist, reports, exports, and presentation assets' }
  ];

  slide.addShape(pptx.ShapeType.line, {
    x: 1.0,
    y: 3.45,
    w: 11.2,
    h: 0,
    line: { color: 'A5C2D2', pt: 2 }
  });

  phases.forEach((phase, idx) => {
    const x = 1.08 + idx * 2.8;

    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 2.9,
      w: 0.62,
      h: 0.62,
      radius: 0.08,
      fill: { color: THEME.cyan },
      line: { color: THEME.cyan }
    });

    slide.addText(String(idx + 1), {
      x: x + 0.2,
      y: 3.08,
      w: 0.2,
      h: 0.2,
      fontSize: 14,
      bold: true,
      color: THEME.white,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addShape(pptx.ShapeType.roundRect, {
      x: x - 0.35,
      y: 4.0,
      w: 2.0,
      h: 2.15,
      radius: 0.08,
      fill: { color: THEME.white },
      line: { color: THEME.line, pt: 1 }
    });

    slide.addText(phase.label, {
      x: x - 0.2,
      y: 4.2,
      w: 1.65,
      h: 0.18,
      fontSize: 10,
      bold: true,
      color: THEME.accent,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addText(phase.name, {
      x: x - 0.25,
      y: 4.45,
      w: 1.75,
      h: 0.35,
      fontSize: 12,
      bold: true,
      color: THEME.ink,
      align: 'center',
      fontFace: 'Calibri'
    });

    slide.addText(phase.detail, {
      x: x - 0.27,
      y: 4.87,
      w: 1.8,
      h: 1.13,
      fontSize: 9,
      color: THEME.text,
      align: 'center',
      breakLine: true,
      fontFace: 'Calibri'
    });
  });
}

function getSnippetFromFile(relativePath, marker, maxLines) {
  try {
    const absolutePath = path.join(__dirname, relativePath);
    const source = fs.readFileSync(absolutePath, 'utf8');
    const lines = source.split(/\r?\n/);
    const start = lines.findIndex((line) => line.includes(marker));

    if (start < 0) {
      return `// Marker not found: ${marker}\n// File: ${relativePath}`;
    }

    return lines
      .slice(start, start + maxLines)
      .map((line) => (line.length > 93 ? `${line.slice(0, 90)}...` : line))
      .join('\n');
  } catch (error) {
    return `// Unable to read ${relativePath}\n// ${error.message}`;
  }
}

function addCodeSlide(slide, title, subtitle, fileLabel, codeSnippet, focusPoints) {
  addHeadline(slide, title, subtitle);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.72,
    y: 1.95,
    w: 7.8,
    h: 4.95,
    radius: 0.08,
    fill: { color: '0F172A' },
    line: { color: '334155', pt: 1 }
  });

  slide.addText(`Source: ${fileLabel}`, {
    x: 0.94,
    y: 2.1,
    w: 7.4,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: '93C5FD',
    fontFace: 'Calibri'
  });

  slide.addText(codeSnippet, {
    x: 0.95,
    y: 2.37,
    w: 7.35,
    h: 4.35,
    fontSize: 9.4,
    color: 'E5E7EB',
    fontFace: 'Consolas',
    breakLine: true,
    valign: 'top',
    margin: 2
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.78,
    y: 1.95,
    w: 3.5,
    h: 4.95,
    radius: 0.08,
    fill: { color: THEME.white },
    line: { color: THEME.line, pt: 1 }
  });

  slide.addText('Why This Matters', {
    x: 9.02,
    y: 2.2,
    w: 3.1,
    h: 0.3,
    fontSize: 13,
    bold: true,
    color: THEME.ink,
    fontFace: 'Calibri'
  });

  const runs = focusPoints.map((text) => ({ text, options: { bullet: { indent: 12 } } }));
  slide.addText(runs, {
    x: 9.02,
    y: 2.58,
    w: 3.08,
    h: 4.1,
    fontSize: 11,
    color: THEME.text,
    breakLine: true,
    paraSpaceAfterPt: 10,
    fontFace: 'Calibri'
  });
}

function addTitleSlide(slide, slideNo, totalSlides) {
  slide.background = { color: THEME.navy };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 7.5,
    fill: { color: THEME.navy },
    line: { color: THEME.navy }
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.86,
    y: 0.82,
    w: 3.4,
    h: 0.52,
    radius: 0.07,
    fill: { color: THEME.accent },
    line: { color: THEME.accent }
  });

  slide.addText('FINAL REVIEW PRESENTATION', {
    x: 1.03,
    y: 0.99,
    w: 3.05,
    h: 0.22,
    fontSize: 11,
    bold: true,
    color: THEME.white,
    fontFace: 'Calibri'
  });

  slide.addText('AI Based Event and Service Management', {
    x: 0.95,
    y: 1.75,
    w: 8.9,
    h: 1.12,
    fontSize: 42,
    bold: true,
    color: THEME.white,
    fontFace: 'Calibri'
  });

  slide.addText('Professional project deck aligned with live code, analytics workflow, and faculty demo requirements.', {
    x: 1.0,
    y: 3.05,
    w: 8.3,
    h: 0.72,
    fontSize: 16,
    color: 'C9DEEA',
    breakLine: true,
    fontFace: 'Calibri'
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.98,
    y: 4.2,
    w: 5.95,
    h: 1.55,
    radius: 0.08,
    fill: { color: '133448' },
    line: { color: '204B63', pt: 1 }
  });

  slide.addText([
    { text: 'Review Date: ', options: { bold: true, color: '9FD7EA' } },
    { text: `${reviewDate}` },
    { text: '\nProject Type: ', options: { bold: true, color: '9FD7EA' } },
    { text: 'Web-based Intelligent Management Platform' },
    { text: '\nCore Modules: ', options: { bold: true, color: '9FD7EA' } },
    { text: 'AI Event Designer, Services, Admin, Analytics, Planner' }
  ], {
    x: 1.2,
    y: 4.48,
    w: 5.6,
    h: 1.1,
    fontSize: 12,
    color: THEME.white,
    breakLine: true,
    fontFace: 'Calibri'
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.2,
    y: 1.15,
    w: 3.2,
    h: 4.9,
    radius: 0.1,
    fill: { color: '123447' },
    line: { color: '2A5670', pt: 1 }
  });

  const highlightItems = [
    'Unified portal architecture',
    'Voice-first AI configuration',
    'Fallback-safe image generation',
    'Readiness scoring and analytics',
    'Planner + budget + risk workflow',
    'Review-ready reporting outputs'
  ];

  slide.addText('Executive Highlights', {
    x: 9.48,
    y: 1.46,
    w: 2.7,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: THEME.white,
    fontFace: 'Calibri'
  });

  const runs = highlightItems.map((text) => ({ text, options: { bullet: { indent: 12 } } }));
  slide.addText(runs, {
    x: 9.45,
    y: 1.82,
    w: 2.72,
    h: 3.95,
    fontSize: 10.7,
    color: 'D3E8F3',
    paraSpaceAfterPt: 8,
    breakLine: true,
    fontFace: 'Calibri'
  });

  slide.addText(`Slide ${slideNo}/${totalSlides}`, {
    x: 11.05,
    y: 7.24,
    w: 1.95,
    h: 0.16,
    fontSize: 9,
    color: '9ABCCF',
    align: 'right',
    fontFace: 'Calibri'
  });
}

const narrativeSlides = [
  {
    title: 'Executive Agenda',
    subtitle: 'A high-impact walkthrough designed for professional final review',
    bullets: [
      'Business problem and why integrated event-service management matters today.',
      'Platform architecture and technical choices with reliability guarantees.',
      'Module deep-dive: AI Event Designer, Services, Admin, Analytics, and Planner.',
      'Evidence section from live source code and controlled fallback behavior.',
      'Readiness status, measurable outcomes, and next-phase scale roadmap.'
    ]
  },
  {
    title: 'Problem Context and Opportunity',
    subtitle: 'Current workflows are fragmented, manual, and hard to audit',
    bullets: [
      'Event setup, service applications, and admin tracking usually live in separate tools.',
      'Operational teams lose visibility due to disconnected data and delayed status updates.',
      'Manual planning raises errors, slows response time, and hurts review readiness.',
      'Institutions need one intelligent platform with both execution and governance capabilities.'
    ]
  },
  {
    title: 'Solution Positioning',
    subtitle: 'One platform combining experience, automation, and governance',
    bullets: [
      'Unified access flow from login to options, event design, and service workflows.',
      'AI-assisted stage planning through voice and prompt-driven generation engine.',
      'Built-in governance via admin controls, analytics intelligence, and planner workspace.',
      'Presentation-ready reporting support to accelerate final review confidence.'
    ]
  },
  {
    title: 'Technology Stack and Runtime',
    subtitle: 'Lightweight web stack, high feature density, practical deployment model',
    bullets: [
      'Frontend: HTML + CSS + JavaScript with modern professional UI patterns.',
      'Backend: Node.js + Express serving static modules and API endpoints.',
      'AI Integration: OpenAI image generation with HuggingFace and local fallback safety.',
      'Browser APIs: Speech recognition and synthesis for voice-first interaction.'
    ]
  },
  {
    title: 'Module 1: Identity and Access',
    subtitle: 'Structured onboarding for users and administrators',
    bullets: [
      'Landing module supports user registration, login, and admin sign-in entry.',
      'Session-aware project profile updates personalize ownership and demo continuity.',
      'Validation and guided interactions increase trust during live demonstration.'
    ]
  },
  {
    title: 'Module 2: AI Event Designer',
    subtitle: 'Voice + parameter control for stage generation workflow',
    bullets: [
      'Stage dimensions, LED patterns, and sound systems are configurable via natural commands.',
      'Complex instructions map to multi-parameter updates in one interaction cycle.',
      'Generated outputs can be saved and surfaced in admin-side event application review.'
    ]
  },
  {
    title: 'Module 3: Service Management',
    subtitle: 'Structured cards and forms for schemes, loans, and support services',
    bullets: [
      'Service options are categorized for faster user discovery and decision-making.',
      'Detail modals and forms collect information required for operational action.',
      'Application flow integrates with governance views for institutional visibility.'
    ]
  },
  {
    title: 'Module 4: Governance Layer',
    subtitle: 'Admin, analytics, and planner are built for review discipline',
    bullets: [
      'Admin dashboard provides user and application management surface.',
      'Analytics dashboard translates system activity into health and readiness indicators.',
      'Planner workspace tracks tasks, budget, milestones, risk, and exports review reports.'
    ]
  },
  {
    title: 'Security, Reliability, and Compatibility',
    subtitle: 'Designed for stable demo execution even under external API failures',
    bullets: [
      'Server-side validation on generation requests prevents malformed input usage.',
      'Multi-level fallback: OpenAI -> HuggingFace -> local sample image path.',
      'Cross-browser support with best voice performance in Chrome and Edge.',
      'Error handling paths preserve user flow and reduce review-time interruptions.'
    ]
  },
  {
    title: 'Review Strategy and Demo Flow',
    subtitle: 'Recommended sequence for faculty review session',
    bullets: [
      'Start with login and options command center to frame full platform breadth.',
      'Run AI event designer using voice command and generation endpoint demonstration.',
      'Show service application, then verify governance trace in admin + analytics.',
      'Close with planner report export and readiness summary for evaluation panel.'
    ]
  },
  {
    title: 'Outcome and Next Phase',
    subtitle: 'Current build is review-strong and ready for production hardening',
    bullets: [
      'Platform already demonstrates end-to-end journey across user, operations, and admin.',
      'AI-driven module creates a strong innovation differentiator for final review.',
      'Next upgrades: persistent database, role-based permissions, deployment automation.',
      'Roadmap can extend into multi-tenant event operations for institutional scale.'
    ]
  }
];

const codeEvidenceSlides = [
  {
    title: 'Code Evidence: Generation Endpoint',
    subtitle: 'Prompt construction and fallback logic in backend API',
    file: 'index.js',
    marker: "app.post('/generate'",
    maxLines: 30,
    notes: [
      'Receives prompt + stage parameters and validates required input.',
      'Builds enhanced prompt ensuring strict composition constraints.',
      'Returns fallback image on failure path to keep demo continuity.'
    ]
  },
  {
    title: 'Code Evidence: Shared Intelligence Model',
    subtitle: 'Readiness score and operational KPIs from shared app layer',
    file: 'app-shared.js',
    marker: 'function getOverview(totalChecklistItems)',
    maxLines: 32,
    notes: [
      'Aggregates users, applications, tasks, milestones, and risks.',
      'Computes readiness score using weighted health factors.',
      'Powers consistent analytics across index, options, admin, and planner.'
    ]
  },
  {
    title: 'Code Evidence: Planner Report Export',
    subtitle: 'Review artifact generation directly from planner workspace',
    file: 'planner.html',
    marker: 'function exportPlannerReport()',
    maxLines: 30,
    notes: [
      'Produces structured review report text for final presentation support.',
      'Captures profile, KPIs, and task-risk status in one export operation.',
      'Demonstrates practical project governance beyond UI screens.'
    ]
  }
];

const totalSlides = 1 + 1 + narrativeSlides.length + 3 + 1 + codeEvidenceSlides.length + 1;
let slideNo = 1;

const titleSlide = pptx.addSlide();
addTitleSlide(titleSlide, slideNo, totalSlides);
slideNo += 1;

const sec1 = pptx.addSlide();
addSectionSlide(sec1, 'SECTION 1', 'Business and Platform Narrative', [
  'Problem framing, solution approach, and architecture decisions.',
  'How each module contributes to measurable review readiness.'
]);
slideNo += 1;

narrativeSlides.forEach((item) => {
  const slide = pptx.addSlide();
  addBase(slide, slideNo, totalSlides);
  addHeadline(slide, item.title, item.subtitle);
  addBulletPanel(slide, item.bullets);
  slideNo += 1;
});

const sec2 = pptx.addSlide();
addSectionSlide(sec2, 'SECTION 2', 'Execution Evidence and Metrics', [
  'Quantitative snapshot, implementation timeline, and code-backed proof.',
  'Focus on reliability, governance depth, and professional delivery quality.'
]);
slideNo += 1;

const kpi = pptx.addSlide();
addBase(kpi, slideNo, totalSlides);
addKpiSlide(kpi, [
  { title: 'Core Modules', value: '8+', note: 'Auth, Options, Event, Service, Admin, Analytics, Planner, Shared' },
  { title: 'Supported Languages', value: '15+', note: 'Inclusive UX with regional language options' },
  { title: 'Generation Resilience', value: '3-Tier', note: 'OpenAI -> HuggingFace -> local fallback' },
  { title: 'Review Readiness', value: 'Live KPI', note: 'Shared readiness scoring across dashboard views' }
]);
slideNo += 1;

const architecture = pptx.addSlide();
addBase(architecture, slideNo, totalSlides);
addArchitectureSlide(architecture);
slideNo += 1;

const timeline = pptx.addSlide();
addBase(timeline, slideNo, totalSlides);
addTimelineSlide(timeline);
slideNo += 1;

codeEvidenceSlides.forEach((item) => {
  const slide = pptx.addSlide();
  addBase(slide, slideNo, totalSlides);
  const code = getSnippetFromFile(item.file, item.marker, item.maxLines);
  addCodeSlide(slide, item.title, item.subtitle, item.file, code, item.notes);
  slideNo += 1;
});

const closing = pptx.addSlide();
addBase(closing, slideNo, totalSlides);
addHeadline(closing, 'Thank You', 'Project ready for faculty review and live demonstration');
addBulletPanel(closing, [
  'The platform demonstrates a complete and integrated event-service lifecycle.',
  'AI innovation is grounded in practical controls, reliability, and evidence slides.',
  'Governance modules strengthen professional quality for final-year evaluation.',
  'Prepared for Q&A, live walkthrough, and next-phase implementation planning.'
], { x: 0.72, y: 2.02, w: 12, h: 3.5 });

closing.addShape(pptx.ShapeType.roundRect, {
  x: 0.72,
  y: 5.75,
  w: 12,
  h: 1.1,
  radius: 0.08,
  fill: { color: 'E8F4FA' },
  line: { color: 'C7DEE9', pt: 1 }
});

closing.addText('Submission Assets', {
  x: 0.95,
  y: 6.02,
  w: 2.4,
  h: 0.2,
  fontSize: 12,
  bold: true,
  color: THEME.ink,
  fontFace: 'Calibri'
});

closing.addText('Presentation_Pro_2026-03-16.pptx | Project_Review_22_Slides.pptx | Event planner report export', {
  x: 0.95,
  y: 6.26,
  w: 11.45,
  h: 0.3,
  fontSize: 11,
  color: THEME.text,
  fontFace: 'Calibri'
});

const outputPath = path.join(__dirname, 'Presentation_Pro_2026-03-16.pptx');
pptx
  .writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`Professional PPT created: ${outputPath}`);
  })
  .catch((error) => {
    console.error('Failed to create professional PPT:', error);
    process.exitCode = 1;
  });
