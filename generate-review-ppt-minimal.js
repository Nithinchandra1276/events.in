const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Project Team';
pptx.title = 'Project Review - Minimal';

const COLORS = {
  primary: '667EEA',
  dark: '1F2937',
  light: 'F8FAFC',
  white: 'FFFFFF'
};

const reviewDate = new Date().toLocaleDateString('en-GB', {
  day: '2-digit',
  month: 'short',
  year: 'numeric'
});

function drawFrame(slide, title, slideNo, totalSlides) {
  slide.background = { color: COLORS.light };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.85,
    fill: { color: COLORS.primary },
    line: { color: COLORS.primary }
  });

  slide.addText(title, {
    x: 0.45,
    y: 0.18,
    w: 12,
    h: 0.4,
    fontSize: 19,
    bold: true,
    color: COLORS.white,
    fontFace: 'Calibri'
  });

  slide.addText(`Slide ${slideNo}/${totalSlides}`, {
    x: 11.7,
    y: 7.18,
    w: 1.3,
    h: 0.2,
    fontSize: 10,
    color: '6B7280',
    align: 'right',
    fontFace: 'Calibri'
  });
}

function drawBullets(slide, points) {
  const runs = points.map((text) => ({ text, options: { bullet: { indent: 16 } } }));
  slide.addText(runs, {
    x: 0.9,
    y: 1.55,
    w: 11.8,
    h: 5.5,
    fontSize: 24,
    color: COLORS.dark,
    breakLine: true,
    paraSpaceAfterPt: 14,
    margin: 4,
    fontFace: 'Calibri'
  });
}

function getSnippetFromFile(relativePath, marker, maxLines = 18) {
  try {
    const absolutePath = path.join(__dirname, relativePath);
    const source = fs.readFileSync(absolutePath, 'utf8');
    const lines = source.split(/\r?\n/);
    const start = lines.findIndex((line) => line.includes(marker));

    if (start < 0) {
      return `// Marker not found: ${marker}\n// File: ${relativePath}`;
    }

    const selected = lines.slice(start, start + maxLines).map((line) => {
      if (line.length <= 90) return line;
      return `${line.slice(0, 87)}...`;
    });

    return selected.join('\n');
  } catch (error) {
    return `// Unable to read file: ${relativePath}\n// ${error.message}`;
  }
}

function drawCode(slide, fileLabel, code) {
  slide.addText(`Source: ${fileLabel}`, {
    x: 0.95,
    y: 1.32,
    w: 11.6,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: '374151',
    fontFace: 'Calibri'
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.85,
    y: 1.65,
    w: 11.8,
    h: 5.55,
    radius: 0.08,
    fill: { color: '0F172A' },
    line: { color: '334155', pt: 1 }
  });

  slide.addText(code, {
    x: 1.05,
    y: 1.92,
    w: 11.35,
    h: 5.05,
    fontSize: 11,
    color: 'E5E7EB',
    fontFace: 'Consolas',
    breakLine: true,
    valign: 'top',
    margin: 3
  });
}

const content = [
  ['Project Review', ['AI Event + Service Management', 'Web app project demo', `Review date: ${reviewDate}`]],
  ['Agenda', ['Problem', 'Features', 'Demo + next steps']],
  ['Problem', ['Planning is fragmented', 'Service access is complex', 'Admin tracking is manual']],
  ['Goal', ['Single platform', 'Faster user flow', 'Smart AI assistance']],
  ['Architecture', ['Frontend pages', 'Node + Express backend', 'API-driven image generation']],
  ['Tech Stack', ['HTML/CSS/JavaScript', 'Express server', 'Speech + AI APIs']],
  ['Login Module', ['User/Admin login', 'Register flow', 'Basic form validation']],
  ['Language Module', ['Translator prompt', '15+ language choices', 'Inclusive UX']],
  ['Options Module', ['Central feature hub', 'Event management path', 'Service management path']],
  ['Event Module', ['Stage size controls', 'LED and sound options', 'Design generation flow']],
  ['AI Assistant', ['Voice commands', 'Natural language handling', 'Audio + visual feedback']],
  ['Sample Commands', ['Set width/height', 'Apply LED and sound', 'Generate multiple designs']],
  ['Generation API', ['POST /generate', 'Prompt enhancement logic', 'Fallback if API fails']],
  ['Service Module', ['Loans and schemes', 'Detail modals', 'Application form flow']],
  ['Admin Module', ['Admin login', 'User management view', 'Application tracking']],
  ['UI/UX', ['Modern gradient style', 'Smooth interactions', 'Responsive layout']],
  ['Security', ['Input checks', 'Error handling', 'Env key support']],
  ['Performance', ['Lightweight stack', 'Fast local startup', 'Cross-device support']],
  ['Verification', ['Startup verified', 'Core routes checked', 'No review-blocking issue']],
  ['Demo Plan', ['Login and options', 'AI event generation', 'Service + admin overview']],
  ['Improvements', ['Database persistence', 'Role-based access', 'Production deployment']],
  ['Conclusion', ['Project is review-ready', 'AI feature is key strength', 'Ready for next phase']]
];

const codeSlides = [
  {
    title: 'Code: Backend API',
    points: ['Real code snippet from image generation endpoint'],
    file: 'index.js',
    marker: "app.post('/generate'",
    maxLines: 19
  },
  {
    title: 'Code: Event Save',
    points: ['Real code snippet from event save workflow'],
    file: 'option1.html',
    marker: 'function saveEventDetails()',
    maxLines: 22
  },
  {
    title: 'Code: Admin Event Table',
    points: ['Real code snippet from admin event rendering'],
    file: 'admin.html',
    marker: 'function renderEventApplications()',
    maxLines: 22
  }
];

const totalSlides = content.length + codeSlides.length;
let slideNo = 1;

content.forEach((item) => {
  const slide = pptx.addSlide();
  drawFrame(slide, item[0], slideNo, totalSlides);
  drawBullets(slide, item[1]);
  slideNo += 1;
});

codeSlides.forEach((item) => {
  const slide = pptx.addSlide();
  drawFrame(slide, item.title, slideNo, totalSlides);
  drawCode(slide, item.file, getSnippetFromFile(item.file, item.marker, item.maxLines));
  slideNo += 1;
});

const outputPath = path.join(__dirname, 'Project_Review_22_Slides_Minimal.pptx');
pptx
  .writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`Minimal PPT created: ${outputPath}`);
  })
  .catch((error) => {
    console.error('Failed to create minimal PPT:', error);
    process.exitCode = 1;
  });
