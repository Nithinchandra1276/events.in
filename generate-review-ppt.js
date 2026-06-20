const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Project Team';
pptx.company = 'Project Review';
pptx.subject = 'AI Based Event and Service Management';
pptx.title = 'Project Review Presentation';
pptx.lang = 'en-US';

const COLORS = {
  primary: '667EEA',
  secondary: '764BA2',
  dark: '1F2937',
  light: 'F8FAFC',
  white: 'FFFFFF'
};

const reviewDate = new Date().toLocaleDateString('en-GB', {
  day: '2-digit',
  month: 'long',
  year: 'numeric'
});

function addHeader(slide, title, subtitle = '') {
  slide.background = { color: COLORS.light };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.9,
    fill: { color: COLORS.primary },
    line: { color: COLORS.primary }
  });

  slide.addText(title, {
    x: 0.4,
    y: 0.2,
    w: 9.8,
    h: 0.45,
    fontSize: 20,
    bold: true,
    color: COLORS.white,
    fontFace: 'Calibri'
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.4,
      y: 1.05,
      w: 12.4,
      h: 0.4,
      fontSize: 13,
      color: COLORS.secondary,
      italic: true,
      fontFace: 'Calibri'
    });
  }
}

function addBullets(slide, bullets) {
  const runs = bullets.map((line) => ({
    text: line,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(runs, {
    x: 0.8,
    y: 1.8,
    w: 11.9,
    h: 5.2,
    fontSize: 20,
    color: COLORS.dark,
    breakLine: true,
    paraSpaceAfterPt: 12,
    margin: 6,
    valign: 'top',
    fontFace: 'Calibri'
  });
}

function addFooter(slide, pageNumber, totalSlides) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 7.2,
    w: 13.33,
    h: 0.3,
    fill: { color: 'E5E7EB' },
    line: { color: 'E5E7EB' }
  });

  slide.addText(`Project Review 2026 | Slide ${pageNumber}/${totalSlides}`, {
    x: 0.5,
    y: 7.23,
    w: 12.2,
    h: 0.2,
    fontSize: 10,
    color: '6B7280',
    align: 'right',
    fontFace: 'Calibri'
  });
}

function addPhotoSlide(slide, title, subtitle, caption, imagePath) {
  addHeader(slide, title, subtitle);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.9,
    y: 1.55,
    w: 11.5,
    h: 4.95,
    radius: 0.08,
    fill: { color: COLORS.white },
    line: { color: 'D1D5DB', pt: 1 }
  });

  try {
    slide.addImage({
      path: imagePath,
      x: 1.15,
      y: 1.8,
      w: 11,
      h: 4.35,
      sizing: { type: 'contain', x: 1.15, y: 1.8, w: 11, h: 4.35 }
    });
  } catch (error) {
    slide.addText('Photo could not be embedded in this environment.', {
      x: 1.3,
      y: 3.7,
      w: 10.6,
      h: 0.35,
      fontSize: 16,
      color: 'B91C1C',
      bold: true,
      align: 'center',
      fontFace: 'Calibri'
    });
  }

  slide.addText(caption, {
    x: 1.05,
    y: 6.72,
    w: 11.9,
    h: 0.35,
    fontSize: 13,
    color: '4B5563',
    italic: true,
    align: 'center',
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

function addCodeSlide(slide, title, subtitle, fileLabel, code) {
  addHeader(slide, title, subtitle);

  slide.addText(`Source: ${fileLabel}`, {
    x: 0.8,
    y: 1.45,
    w: 12,
    h: 0.3,
    fontSize: 12,
    color: '374151',
    bold: true,
    fontFace: 'Calibri'
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 1.78,
    w: 12,
    h: 5.45,
    radius: 0.08,
    fill: { color: '0F172A' },
    line: { color: '334155', pt: 1 }
  });

  slide.addText(code, {
    x: 0.9,
    y: 2.02,
    w: 11.6,
    h: 5.0,
    fontSize: 11,
    color: 'E5E7EB',
    fontFace: 'Consolas',
    valign: 'top',
    margin: 3,
    breakLine: true
  });
}

const slides = [
  {
    title: 'AI Based Event and Service Management',
    subtitle: 'Project Review Presentation',
    bullets: [
      `Review Date: ${reviewDate}`,
      'Project Type: Web Application',
      'Core Areas: Event Management + Service Management',
      'Prepared for today\'s project review meeting'
    ]
  },
  {
    title: 'Agenda',
    subtitle: 'What this review covers',
    bullets: [
      'Problem statement and project goals',
      'System architecture and technology stack',
      'Key modules and major features',
      'AI assistant capabilities and demo flow',
      'Testing, performance, security, and next steps'
    ]
  },
  {
    title: 'Problem Statement',
    subtitle: 'Why this project is needed',
    bullets: [
      'Event planning and service access are often fragmented',
      'Users need one platform for event setup and support services',
      'Manual workflows reduce speed and consistency',
      'Admins require clear control over users and applications'
    ]
  },
  {
    title: 'Project Objectives',
    subtitle: 'Primary goals',
    bullets: [
      'Build a modern, user-friendly, responsive web platform',
      'Provide event stage customization with AI assistance',
      'Include financial services and scheme application workflows',
      'Add admin visibility for operational management'
    ]
  },
  {
    title: 'High-Level Architecture',
    subtitle: 'Frontend + backend overview',
    bullets: [
      'Frontend: HTML, CSS, JavaScript pages',
      'Backend: Node.js + Express server',
      'Static content served from project root',
      'API endpoint /generate for stage image generation'
    ]
  },
  {
    title: 'Technology Stack',
    subtitle: 'Tools and frameworks used',
    bullets: [
      'Runtime: Node.js',
      'Server Framework: Express',
      'Styling: Custom CSS with gradients and animations',
      'AI/Media APIs: Gemini Imagen + HuggingFace fallback',
      'Browser APIs: Web Speech + Speech Synthesis'
    ]
  },
  {
    title: 'Module 1: Authentication',
    subtitle: 'Entry point and access control',
    bullets: [
      'Landing page supports User and Admin login flows',
      'User registration with password confirmation',
      'Password visibility toggles improve usability',
      'Form validation and guided UI interactions'
    ]
  },
  {
    title: 'Module 2: Language Support',
    subtitle: 'Inclusive user experience',
    bullets: [
      'Translator modal prompts users on first interaction',
      'Supports multiple Indian and regional languages',
      'Language selection integrated into user journey',
      'Improves accessibility for diverse audience'
    ]
  },
  {
    title: 'Module 3: Options Hub',
    subtitle: 'Feature branching page',
    bullets: [
      'Options page acts as central navigation layer',
      'Path A: Event Management',
      'Path B: Service Management',
      'Interactive card-based UI with responsive behavior'
    ]
  },
  {
    title: 'Module 4: Event Management',
    subtitle: 'AI Event Planner and Stage Designer',
    bullets: [
      'Stage configuration includes width and height controls',
      'LED pattern selection: geometric, neon, hologram, etc.',
      'Sound system options: stereo, quad, surround, line array',
      'Supports equipment combinations and gallery workflow'
    ]
  },
  {
    title: 'AI Voice Assistant',
    subtitle: 'Key differentiator of the project',
    bullets: [
      'Voice-enabled command interface for stage design',
      'Understands natural language and complex instructions',
      'Provides visual and spoken feedback',
      'Continuous listening and context-aware responses'
    ]
  },
  {
    title: 'Voice Command Examples',
    subtitle: 'Real usage patterns',
    bullets: [
      '"Set width to 30 feet"',
      '"Use neon LED with surround sound"',
      '"Create professional concert setup"',
      '"Generate 3 different designs"',
      '"Reset everything" or "Help" for guided usage'
    ]
  },
  {
    title: 'Image Generation Pipeline',
    subtitle: 'Server-side generation logic',
    bullets: [
      'POST /generate receives prompt and configuration',
      'Builds enhanced prompt focused on stage and sound assets',
      'Tries Gemini Imagen model first',
      'Falls back to HuggingFace, then local sample image'
    ]
  },
  {
    title: 'Module 5: Service Management',
    subtitle: 'Loans and schemes support',
    bullets: [
      'Financial services section with categorized cards',
      'Government scheme information and eligibility details',
      'Modal-based deep details and application forms',
      'Responsive layout for mobile and desktop users'
    ]
  },
  {
    title: 'Module 6: Admin Panel',
    subtitle: 'Operational control center',
    bullets: [
      'Admin login route from landing page',
      'User list and management actions',
      'Application tracking for service requests',
      'Clean dashboard layout with strong readability'
    ]
  },
  {
    title: 'UI/UX Highlights',
    subtitle: 'Design choices and user impact',
    bullets: [
      'Glassmorphism-inspired cards and modern gradients',
      'Smooth transitions, hover effects, and animated sections',
      'Consistent typography and visual hierarchy',
      'Intuitive forms and clear interaction flow'
    ]
  },
  {
    title: 'Business Impact Snapshot',
    subtitle: 'Practical value for users and admins',
    bullets: [
      'Reduces event planning effort through one guided interface',
      'Improves user confidence with voice and visual confirmations',
      'Speeds up service discovery with organized card-based journeys',
      'Provides admins a single place to monitor requests and users'
    ]
  },
  {
    title: 'Security and Validation',
    subtitle: 'Reliability by design',
    bullets: [
      'Input validation applied on critical server routes',
      'Error handling for API failures and missing inputs',
      'Environment variable support for external API keys',
      'Dependency vulnerabilities addressed in verification phase'
    ]
  },
  {
    title: 'Performance and Compatibility',
    subtitle: 'Runtime behavior and support matrix',
    bullets: [
      'Optimized event handling and responsive rendering',
      'Fast local server startup with Express',
      'Best voice experience on Chrome and Edge',
      'Cross-device compatibility: mobile, tablet, desktop'
    ]
  },
  {
    title: 'Testing and Verification',
    subtitle: 'Current quality status',
    bullets: [
      'Project verification report indicates stable startup',
      'Core pages and feature routes validated',
      'Fallback behavior tested for generation failures',
      'No critical blocker identified for project review demo'
    ]
  },
  {
    title: 'Demo Flow for Review',
    subtitle: 'Suggested live walkthrough sequence',
    bullets: [
      'Step 1: Login as user and show language prompt',
      'Step 2: Open Options and choose Event Management',
      'Step 3: Run voice command and generate stage output',
      'Step 4: Switch to Service Management and submit sample form',
      'Step 5: Show Admin panel overview'
    ]
  },
  {
    title: 'Challenges and Improvements',
    subtitle: 'What we learned and what can improve',
    bullets: [
      'Handling API reliability required robust fallback design',
      'Voice UX tuning improved recognition clarity',
      'Future: persistent database integration for users/forms',
      'Future: role-based access and audit logging'
    ]
  },
  {
    title: 'Conclusion and Next Steps',
    subtitle: 'Project status summary',
    bullets: [
      'Project delivers a complete, feature-rich web platform',
      'AI-enabled event planning is the strongest innovation point',
      'Current build is review-ready and demo-capable',
      'Next phase: production hardening and deployment pipeline'
    ]
  }
];

const codeSlides = [
  {
    title: 'Code Evidence: Backend API',
    subtitle: 'Actual implementation used for image generation',
    file: 'index.js',
    marker: "app.post('/generate'",
    maxLines: 20
  },
  {
    title: 'Code Evidence: Event Save Flow',
    subtitle: 'Event details are persisted and available in Admin',
    file: 'option1.html',
    marker: 'function saveEventDetails()',
    maxLines: 22
  },
  {
    title: 'Code Evidence: Admin Event View',
    subtitle: 'Admin event table rendering and management',
    file: 'admin.html',
    marker: 'function renderEventApplications()',
    maxLines: 22
  }
];

const reviewPhotoPath = path.join(__dirname, 'large-outdoor-concert-stage-truss-600nw-2695141455.webp');
const totalSlides = slides.length + codeSlides.length + 1;
let pageNumber = 1;

slides.forEach((item) => {
  const slide = pptx.addSlide();
  addHeader(slide, item.title, item.subtitle);
  addBullets(slide, item.bullets);
  addFooter(slide, pageNumber, totalSlides);
  pageNumber += 1;
});

codeSlides.forEach((item) => {
  const slide = pptx.addSlide();
  const codeSnippet = getSnippetFromFile(item.file, item.marker, item.maxLines);
  addCodeSlide(slide, item.title, item.subtitle, item.file, codeSnippet);
  addFooter(slide, pageNumber, totalSlides);
  pageNumber += 1;
});

const photoSlide = pptx.addSlide();
addPhotoSlide(
  photoSlide,
  'Project Photos',
  'Sample stage output used in demonstration',
  'Reference photo: AI-focused stage and sound setup for review demo.',
  reviewPhotoPath
);
addFooter(photoSlide, pageNumber, totalSlides);

const outputPath = path.join(__dirname, 'Project_Review_22_Slides.pptx');

pptx
  .writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`PPT created successfully: ${outputPath}`);
  })
  .catch((error) => {
    console.error('Failed to create PPT:', error);
    process.exitCode = 1;
  });
