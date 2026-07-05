const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Header,
  Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, HeadingLevel, TabStopPosition, TabStopType,
  convertInchesToTwip, PageBreak
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

const PORT = process.env.PORT || 3000;

// ── Helpers ──────────────────────────────────────────────────────────

function addSection(items, fn) {
  return items && items.length > 0 ? fn() : [];
}

function bullet(text) {
  return new Paragraph({
    spacing: { after: 40 },
    bullet: { level: 0 },
    children: [
      new TextRun({ text, size: 22, font: 'Calibri' }),
    ],
  });
}

function simplePara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after !== undefined ? opts.after : 80 },
    children: [
      new TextRun({
        text,
        size: opts.size || 22,
        font: 'Calibri',
        bold: opts.bold || false,
        color: opts.color || '000000',
      }),
    ],
  });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: '333333' },
    },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        size: 26,
        font: 'Calibri',
        bold: true,
        color: '2E4057',
      }),
    ],
  });
}

function dateRange(start, end) {
  if (!start && !end) return '';
  if (!end) return `${start} – Present`;
  if (!start) return `Until ${end}`;
  return `${start} – ${end}`;
}

// ── DOCX builder ─────────────────────────────────────────────────────

async function buildDocx(json) {
  const children = [];

  // ── Basics ───────────────────────────────────────────────────────
  const b = json.basics || {};
  if (b.name) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [
          new TextRun({
            text: b.name,
            size: 44,
            font: 'Calibri',
            bold: true,
            color: '1C2833',
          }),
        ],
      })
    );
  }

  const contactParts = [];
  if (b.label) contactParts.push(b.label);
  if (b.email) contactParts.push(b.email);
  if (b.phone) contactParts.push(b.phone);
  if (b.location) {
    const loc = b.location;
    const locStr = [loc.city, loc.region, loc.countryCode].filter(Boolean).join(', ');
    if (locStr) contactParts.push(locStr);
  }
  if (b.url) contactParts.push(b.url);

  if (contactParts.length) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: contactParts.join('  |  '),
            size: 20,
            font: 'Calibri',
            color: '566573',
          }),
        ],
      })
    );
  }

  if (b.summary) {
    children.push(
      new Paragraph({
        spacing: { before: 80, after: 120 },
        children: [
          new TextRun({
            text: b.summary,
            size: 22,
            font: 'Calibri',
            color: '2C3E50',
          }),
        ],
      })
    );
  }

  // ── Work ─────────────────────────────────────────────────────────
  addSection(json.work, () => {
    children.push(sectionHeader('Experience'));
    for (const w of json.work) {
      // company / position row
      const headerParts = [];
      headerParts.push(
        new TextRun({ text: w.position || 'Position', size: 24, font: 'Calibri', bold: true })
      );
      if (w.company) {
        headerParts.push(
          new TextRun({ text: `  –  ${w.company}`, size: 24, font: 'Calibri', color: '566573' })
        );
      }

      children.push(
        new Paragraph({
          spacing: { before: 140, after: 0 },
          children: headerParts,
        })
      );

      const dates = dateRange(w.startDate, w.endDate);
      if (dates || w.location) {
        children.push(
          new Paragraph({
            spacing: { after: 60 },
            children: [
              new TextRun({
                text: [dates, w.location].filter(Boolean).join('  |  '),
                size: 20,
                font: 'Calibri',
                italics: true,
                color: '7F8C8D',
              }),
            ],
          })
        );
      }

      if (w.summary) {
        children.push(simplePara(w.summary, { after: 60 }));
      }

      if (w.highlights && w.highlights.length) {
        for (const h of w.highlights) {
          children.push(bullet(h));
        }
      }
    }
  });

  // ── Education ────────────────────────────────────────────────────
  addSection(json.education, () => {
    children.push(sectionHeader('Education'));
    for (const e of json.education) {
      const titleParts = [];
      if (e.studyType) titleParts.push(e.studyType);
      if (e.area) titleParts.push(e.area);
      const title = titleParts.join(' in ');

      children.push(
        new Paragraph({
          spacing: { before: 140, after: 0 },
          children: [
            new TextRun({ text: title || 'Education', size: 24, font: 'Calibri', bold: true }),
            e.institution
              ? new TextRun({
                  text: `  –  ${e.institution}`,
                  size: 24,
                  font: 'Calibri',
                  color: '566573',
                })
              : undefined,
          ].filter(Boolean),
        })
      );

      const dates = dateRange(e.startDate, e.endDate);
      const dateParts = [dates];
      if (e.gpa) dateParts.push(`GPA: ${e.gpa}`);
      const dateText = dateParts.filter(Boolean).join('  |  ');
      if (dateText) {
        children.push(
          new Paragraph({
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: dateText,
                size: 20,
                font: 'Calibri',
                italics: true,
                color: '7F8C8D',
              }),
            ],
          })
        );
      }
    }
  });

  // ── Skills ───────────────────────────────────────────────────────
  addSection(json.skills, () => {
    children.push(sectionHeader('Skills'));
    for (const s of json.skills) {
      const parts = [s.name];
      if (s.level) parts.push(`(${s.level})`);
      let line = parts.join(' ');
      if (s.keywords && s.keywords.length) {
        line += ': ' + s.keywords.join(', ');
      }
      children.push(simplePara(line, { after: 40 }));
    }
  });

  // ── Projects ─────────────────────────────────────────────────────
  addSection(json.projects, () => {
    children.push(sectionHeader('Projects'));
    for (const p of json.projects) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 0 },
          children: [
            new TextRun({ text: p.name || 'Project', size: 24, font: 'Calibri', bold: true }),
            p.url
              ? new TextRun({
                  text: `  (${p.url})`,
                  size: 20,
                  font: 'Calibri',
                  color: '2980B9',
                })
              : undefined,
          ].filter(Boolean),
        })
      );

      if (p.description) {
        children.push(simplePara(p.description, { after: 60 }));
      }

      if (p.highlights && p.highlights.length) {
        for (const h of p.highlights) {
          children.push(bullet(h));
        }
      }
    }
  });

  // ── Awards / Certifications (optional) ────────────────────────────
  addSection(json.awards, () => {
    children.push(sectionHeader('Awards'));
    for (const a of json.awards) {
      children.push(
        new Paragraph({
          spacing: { before: 100, after: 40 },
          children: [
            new TextRun({ text: a.title, size: 24, font: 'Calibri', bold: true }),
            a.awarder
              ? new TextRun({ text: `  –  ${a.awarder}`, size: 22, font: 'Calibri', color: '566573' })
              : undefined,
          ].filter(Boolean),
        })
      );
      if (a.summary) children.push(simplePara(a.summary, { after: 40 }));
    }
  });

  addSection(json.certificates, () => {
    children.push(sectionHeader('Certifications'));
    for (const c of json.certificates) {
      children.push(
        new Paragraph({
          spacing: { before: 100, after: 40 },
          children: [
            new TextRun({ text: c.name, size: 24, font: 'Calibri', bold: true }),
            c.issuer
              ? new TextRun({ text: `  –  ${c.issuer}`, size: 22, font: 'Calibri', color: '566573' })
              : undefined,
          ].filter(Boolean),
        })
      );
    }
  });

  // ── Build document ────────────────────────────────────────────────
  const doc = new Document({
    creator: 'Instant DOCX Resume',
    title: `${b.name || 'Resume'}`,
    styles: {
      default: {
        document: {
          run: { size: 22, font: 'Calibri' },
          paragraph: { spacing: { after: 80 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.7),
              bottom: convertInchesToTwip(0.7),
              left: convertInchesToTwip(0.8),
              right: convertInchesToTwip(0.8),
            },
          },
        },
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
}

// ── Route ─────────────────────────────────────────────────────────────

app.post('/resume', async (req, res) => {
  try {
    const body = req.body;

    if (!body || typeof body !== 'object') {
      return res.status(400).json({ error: 'Request body must be a JSON object (JSON Resume format).' });
    }

    const buffer = await buildDocx(body);

    const name = (body.basics && body.basics.name) ? body.basics.name.replace(/\s+/g, '_') : 'resume';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${name}.docx"`);
    res.send(buffer);
  } catch (err) {
    console.error('Error generating resume:', err);
    res.status(500).json({ error: 'Failed to generate resume DOCX', detail: err.message });
  }
});

// ── Health check ──────────────────────────────────────────────────────

app.get('/', (req, res) => {
  res.json({ service: 'instant-docx-resume', status: 'ok', usage: 'POST /resume with JSON Resume body' });
});

app.get('/health', (req, res) => {
  res.json({ service: 'instant-docx-resume', status: 'ok' });
});

// ── Start ─────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`Instant DOCX Resume service running on http://localhost:${PORT}`);
});