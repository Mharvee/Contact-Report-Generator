'use strict';

require('dotenv').config();
const express  = require('express');
const multer   = require('multer');
const fetch    = require('node-fetch');
const path     = require('path');
const fs       = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  HeadingLevel
} = require('docx');

// ── Config ─────────────────────────────────────────────────────────────────
const PORT          = process.env.PORT || 3000;
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
const GEMINI_KEY    = process.env.GEMINI_API_KEY;

if (!ANTHROPIC_KEY) { console.error('ERROR: ANTHROPIC_API_KEY is not set in .env'); process.exit(1); }
if (!GEMINI_KEY)    { console.error('ERROR: GEMINI_API_KEY is not set in .env');    process.exit(1); }

// ── App setup ──────────────────────────────────────────────────────────────
const app = express();

// Accept audio/video files (for transcription) AND text files (transcripts)
const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const allowedAudioMime = [
      'audio/mpeg', 'audio/mp4', 'audio/wav', 'audio/ogg',
      'audio/webm', 'audio/flac', 'audio/x-m4a',
      'video/mp4',  'video/webm'
    ];
    const allowedAudioExt = ['.mp3','.mp4','.m4a','.wav','.ogg','.webm','.flac'];
    const allowedTextExt  = ['.txt','.md','.text'];
    const ext = path.extname(file.originalname).toLowerCase();

    if (allowedAudioMime.includes(file.mimetype) || allowedAudioExt.includes(ext) || allowedTextExt.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Unsupported file type: ' + file.originalname));
    }
  }
});

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Helpers ────────────────────────────────────────────────────────────────
function geminiMimeType(originalname, mimetype) {
  const ext = path.extname(originalname).toLowerCase();
  const map = {
    '.mp3':  'audio/mpeg',
    '.mp4':  'audio/mp4',
    '.m4a':  'audio/mp4',
    '.wav':  'audio/wav',
    '.ogg':  'audio/ogg',
    '.webm': 'audio/webm',
    '.flac': 'audio/flac',
  };
  return map[ext] || mimetype || 'audio/mpeg';
}

function isAudioFile(originalname, mimetype) {
  const audioMime = ['audio/mpeg','audio/mp4','audio/wav','audio/ogg','audio/webm','audio/flac','audio/x-m4a','video/mp4','video/webm'];
  const audioExt  = ['.mp3','.mp4','.m4a','.wav','.ogg','.webm','.flac'];
  const ext = path.extname(originalname).toLowerCase();
  return audioMime.includes(mimetype) || audioExt.includes(ext);
}

// ── DOCX builder (matching Contact_Report_Template.docx layout) ────────────
async function buildDocx(fields) {
  const today = new Date().toLocaleDateString('en-GB', {
    day: '2-digit', month: 'long', year: 'numeric'
  });

  const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: border, bottom: border, left: border, right: border };

  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  const headerBorder = {
    top:    { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' },
    bottom: { style: BorderStyle.SINGLE, size: 6, color: '1F3864', space: 2 },
    left:   { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' },
    right:  { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' }
  };

  // Helper: labelled cell (dark header + value)
  function labelCell(label, value, width) {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      borders,
      children: [
        new Paragraph({
          children: [new TextRun({ text: label, bold: true, size: 18, font: 'Calibri', color: '1F3864' })],
          spacing: { before: 60, after: 0 }
        }),
        new Paragraph({
          children: [new TextRun({ text: value || '', size: 20, font: 'Calibri' })],
          spacing: { before: 0, after: 60 }
        })
      ],
      margins: { top: 80, bottom: 80, left: 120, right: 120 }
    });
  }

  function fullWidthCell(label, value, bgColor) {
    return new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      borders,
      shading: bgColor ? { fill: bgColor, type: ShadingType.CLEAR } : undefined,
      columnSpan: 4,
      children: [
        new Paragraph({
          children: [new TextRun({ text: label, bold: true, size: 18, font: 'Calibri', color: '1F3864' })],
          spacing: { before: 60, after: 0 }
        }),
        ...(value || '').split('\n').map(line =>
          new Paragraph({
            children: [new TextRun({ text: line, size: 20, font: 'Calibri' })],
            spacing: { before: 0, after: 40 }
          })
        )
      ],
      margins: { top: 80, bottom: 80, left: 120, right: 120 }
    });
  }

  // Company names extracted from attendees or use placeholders
  const clientCompany  = extractCompany(fields.attendees, 'client') || 'Client Organisation';
  const agencyCompany  = extractCompany(fields.attendees, 'agency') || 'Cerebre Digital';

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        // ── Title ──────────────────────────────────────
        new Paragraph({
          children: [new TextRun({ text: 'CONTACT REPORT', bold: true, size: 36, font: 'Calibri', color: '1F3864' })],
          alignment: AlignmentType.CENTER,
          border: headerBorder,
          spacing: { before: 0, after: 240 }
        }),

        // ── Client row ─────────────────────────────────
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1170, 3510, 1170, 3510],
          rows: [
            new TableRow({
              children: [
                labelCell('CLIENT:', fields.client || clientCompany, 1170),
                new TableCell({ width:{size:3510,type:WidthType.DXA}, borders, children:[new Paragraph({ children:[new TextRun({text:'',size:20})], spacing:{before:80,after:80} })], margins:{top:80,bottom:80,left:120,right:120} }),
                new TableCell({ width:{size:1170,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} }),
                new TableCell({ width:{size:3510,type:WidthType.DXA}, borders, children:[new Paragraph({ children:[new TextRun({text:'',size:20})], spacing:{before:80,after:80} })], margins:{top:80,bottom:80,left:120,right:120} }),
              ]
            })
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Present / Distribution row ─────────────────
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({
              children: [
                labelCell(`PRESENT FOR ${clientCompany}`, (fields.clientAttendees || ''), 2340),
                new TableCell({ width:{size:2340,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} }),
                labelCell(`PRESENT FOR ${agencyCompany}`, (fields.agencyAttendees || ''), 2340),
                new TableCell({ width:{size:2340,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} }),
              ]
            }),
            new TableRow({
              children: [
                labelCell('DISTRIBUTION', '', 2340),
                new TableCell({ width:{size:2340,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} }),
                labelCell('DISTRIBUTION', '', 2340),
                new TableCell({ width:{size:2340,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} }),
              ]
            }),
            new TableRow({
              children: [
                labelCell('DATE OF MEETING', fields.date || '', 2340),
                labelCell('DATE OF REPORT', today, 2340),
                labelCell('TIME STARTED', fields.time || '', 2340),
                labelCell('TIME ENDED', fields.timeEnded || '', 2340),
              ]
            })
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Venue + Key Discussions ─────────────────────
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1170, 8190],
          rows: [
            new TableRow({
              children: [
                labelCell('VENUE:', fields.venue || '', 1170),
                new TableCell({ width:{size:8190,type:WidthType.DXA}, borders, children:[new Paragraph({children:[new TextRun({text:fields.venue||'',size:20})]})], margins:{top:80,bottom:80,left:120,right:120} })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({
                  width:{size:9360,type:WidthType.DXA}, borders, columnSpan:2,
                  children:[
                    new Paragraph({ children:[new TextRun({text:'Highlights: KEY DISCUSSIONS', bold:true, size:18, font:'Calibri', color:'1F3864'})], spacing:{before:80,after:0} }),
                    new Paragraph({ children:[new TextRun({text:'',size:20})], spacing:{before:0,after:80} })
                  ],
                  margins:{top:80,bottom:80,left:120,right:120}
                })
              ]
            }),
            ...(fields.discussionPoints || '').split('\n').filter(l=>l.trim()).map(line =>
              new TableRow({
                children:[
                  new TableCell({
                    width:{size:9360,type:WidthType.DXA}, borders, columnSpan:2,
                    children:[new Paragraph({ children:[new TextRun({text:line,size:20,font:'Calibri'})], spacing:{before:40,after:40} })],
                    margins:{top:60,bottom:60,left:120,right:120}
                  })
                ]
              })
            )
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Objective ──────────────────────────────────
        new Table({
          width:{ size:9360, type:WidthType.DXA },
          columnWidths:[2340,7020],
          rows:[
            new TableRow({ children:[
              labelCell('MEETING OBJECTIVE', '', 2340),
              new TableCell({
                width:{size:7020,type:WidthType.DXA}, borders,
                children:(fields.objective||'').split('\n').map(l=>new Paragraph({children:[new TextRun({text:l,size:20,font:'Calibri'})],spacing:{before:40,after:40}})),
                margins:{top:80,bottom:80,left:120,right:120}
              })
            ]})
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Decisions ──────────────────────────────────
        new Table({
          width:{ size:9360, type:WidthType.DXA },
          columnWidths:[9360],
          rows:[
            new TableRow({ children:[
              new TableCell({
                width:{size:9360,type:WidthType.DXA}, borders,
                children:[
                  new Paragraph({children:[new TextRun({text:'DECISIONS MADE', bold:true, size:18, font:'Calibri', color:'1F3864'})],spacing:{before:80,after:0}}),
                  ...(fields.decisions||'').split('\n').map(l=>new Paragraph({children:[new TextRun({text:l,size:20,font:'Calibri'})],spacing:{before:40,after:40}}))
                ],
                margins:{top:80,bottom:80,left:120,right:120}
              })
            ]})
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Action Items ────────────────────────────────
        new Table({
          width:{ size:9360, type:WidthType.DXA },
          columnWidths:[9360],
          rows:[
            new TableRow({ children:[
              new TableCell({
                width:{size:9360,type:WidthType.DXA}, borders,
                children:[
                  new Paragraph({children:[new TextRun({text:'ACTION ITEMS', bold:true, size:18, font:'Calibri', color:'1F3864'})],spacing:{before:80,after:0}}),
                  ...(fields.actionItems||'').split('\n').map(l=>new Paragraph({children:[new TextRun({text:l,size:20,font:'Calibri'})],spacing:{before:40,after:40}}))
                ],
                margins:{top:80,bottom:80,left:120,right:120}
              })
            ]})
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:120, after:0 } }),

        // ── Next Steps + Follow-up ──────────────────────
        new Table({
          width:{ size:9360, type:WidthType.DXA },
          columnWidths:[9360],
          rows:[
            new TableRow({ children:[
              new TableCell({
                width:{size:9360,type:WidthType.DXA}, borders,
                children:[
                  new Paragraph({children:[new TextRun({text:'NEXT STEPS', bold:true, size:18, font:'Calibri', color:'1F3864'})],spacing:{before:80,after:0}}),
                  ...(fields.nextSteps||'').split('\n').map(l=>new Paragraph({children:[new TextRun({text:l,size:20,font:'Calibri'})],spacing:{before:40,after:40}})),
                  new Paragraph({children:[new TextRun({text:`FOLLOW-UP DATE: ${fields.followUpDate||'To be confirmed'}`, bold:true, size:18, font:'Calibri', color:'1F3864'})],spacing:{before:120,after:80}})
                ],
                margins:{top:80,bottom:80,left:120,right:120}
              })
            ]})
          ]
        }),

        new Paragraph({ children:[], spacing:{ before:240, after:0 } }),

        // ── Disclaimer footer ──────────────────────────
        new Paragraph({
          children:[
            new TextRun({
              text:`This Contact Report has been produced as a representation of the agreed actions and conversations that took place on the date stated above. If there are statements that are not accurate, these should be brought to the attention of ${agencyCompany} within 48 hours of the date of this report.`,
              bold:true, size:16, font:'Calibri'
            })
          ],
          border:{ top:{ style:BorderStyle.SINGLE, size:4, color:'1F3864' } },
          spacing:{ before:240, after:120 }
        }),

        new Paragraph({
          children:[new TextRun({text:`On behalf of ${agencyCompany}`, bold:true, size:16, font:'Calibri'})],
          spacing:{ before:0, after:0 }
        }),
      ]
    }]
  });

  return Packer.toBuffer(doc);
}

// Simple heuristic: split attendees into two groups (client vs agency)
function extractCompany(attendees, role) {
  if (!attendees) return null;
  // Try to extract company from lines like "• Name (Company)" or "Name — Company"
  const lines = attendees.split('\n').map(l => l.trim()).filter(Boolean);
  // Return unique company names
  const companies = [];
  lines.forEach(line => {
    const m = line.match(/\(([^)]+)\)/) || line.match(/[—–-]\s*(.+)$/);
    if (m) companies.push(m[1].trim());
  });
  if (companies.length === 0) return null;
  if (role === 'client') return companies[0];
  if (role === 'agency' && companies.length > 1) return companies[1];
  return null;
}

// Split attendees into client-side and agency-side (first half / second half heuristic)
function splitAttendees(attendeeText) {
  if (!attendeeText) return { client: '', agency: '' };
  const lines = attendeeText.split('\n').filter(l => l.trim());
  const mid = Math.ceil(lines.length / 2);
  return {
    client: lines.slice(0, mid).join('\n'),
    agency: lines.slice(mid).join('\n')
  };
}

// ── POST /api/upload ───────────────────────────────────────────────────────
// Unified endpoint: accepts audio/video files OR text transcripts
// - Audio → transcribes via Gemini, returns { transcript, inputType:'audio' }
// - Text  → reads as transcript directly, returns { transcript, inputType:'transcript' }
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file provided.' });

  const ext = path.extname(req.file.originalname).toLowerCase();

  // ── Branch: transcript text file ──
  if (['.txt','.md','.text'].includes(ext) || req.file.mimetype === 'text/plain') {
    const transcript = req.file.buffer.toString('utf8');
    if (!transcript.trim()) return res.status(400).json({ error: 'The uploaded transcript file is empty.' });
    return res.json({ transcript, inputType: 'transcript' });
  }

  // ── Branch: audio/video file → Gemini transcription ──
  if (!isAudioFile(req.file.originalname, req.file.mimetype)) {
    return res.status(400).json({ error: 'Unsupported file type. Upload an audio/video file or a .txt transcript.' });
  }

  const audioBase64 = req.file.buffer.toString('base64');
  const mimeType    = geminiMimeType(req.file.originalname, req.file.mimetype);

  try {
    const geminiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [
              {
                text: 'Transcribe this audio recording verbatim and completely. ' +
                      'Identify different speakers and label them (e.g. Speaker 1:, Speaker 2:, or use names if mentioned). ' +
                      'Insert a timestamp marker in [MM:SS] format every 2-3 minutes. ' +
                      'Output ONLY the transcript text — no preamble, no summary, no commentary.'
              },
              { inline_data: { mime_type: mimeType, data: audioBase64 } }
            ]
          }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
        })
      }
    );

    const geminiData = await geminiRes.json();
    if (!geminiRes.ok) {
      const msg = geminiData?.error?.message || `Gemini error ${geminiRes.status}`;
      return res.status(502).json({ error: 'Transcription failed: ' + msg });
    }

    const transcript = geminiData?.candidates?.[0]?.content?.parts
      ?.filter(p => p.text)?.map(p => p.text)?.join('') || '';

    if (!transcript) return res.status(502).json({ error: 'Gemini returned an empty transcript.' });

    res.json({ transcript, inputType: 'audio' });

  } catch (err) {
    console.error('[/api/upload]', err);
    res.status(500).json({ error: 'Transcription request failed: ' + err.message });
  }
});

// Keep the old /api/transcribe endpoint for backwards compatibility
app.post('/api/transcribe', upload.single('audio'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No audio file provided.' });
  // Re-route through the upload logic
  req.file.fieldname = 'file';
  const fakeReq = { ...req, file: req.file };
  // Just call the same logic inline
  const audioBase64 = req.file.buffer.toString('base64');
  const mimeType    = geminiMimeType(req.file.originalname, req.file.mimetype);
  try {
    const geminiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [
            { text: 'Transcribe this audio recording verbatim and completely. Identify different speakers. Output ONLY the transcript text.' },
            { inline_data: { mime_type: mimeType, data: audioBase64 } }
          ]}],
          generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
        })
      }
    );
    const d = await geminiRes.json();
    const transcript = d?.candidates?.[0]?.content?.parts?.filter(p=>p.text)?.map(p=>p.text)?.join('') || '';
    res.json({ transcript });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/analyse ──────────────────────────────────────────────────────
app.post('/api/analyse', async (req, res) => {
  const { transcript } = req.body;
  if (!transcript || typeof transcript !== 'string' || transcript.trim().length < 10) {
    return res.status(400).json({ error: 'No transcript provided.' });
  }

  const today = new Date().toLocaleDateString('en-GB', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });

  try {
    const claudeRes = await fetch('https://api.anthropic.com/v1/messages', {
      method:  'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model:      'claude-sonnet-4-6',
        max_tokens: 4096,
        system: 'You are a senior business analyst specialising in meeting documentation. ' +
                'Extract structured data from meeting transcripts to populate a Contact Report. ' +
                'Respond with ONLY a valid JSON object — no markdown fences, no preamble, no explanation.',
        messages: [{
          role: 'user',
          content: `Analyse this meeting transcript and extract all relevant information for a Contact Report.

Today's reference date: ${today}

TRANSCRIPT:
${transcript}

Return ONLY a JSON object with exactly these keys:
{
  "client": "Client company/organisation name",
  "date": "Meeting date extracted from transcript. If not mentioned, use today's date.",
  "time": "Meeting start time if mentioned, otherwise empty string",
  "timeEnded": "Meeting end time if mentioned, otherwise empty string",
  "duration": "Duration — estimated from timestamps or stated in the recording",
  "venue": "Meeting venue or 'Virtual / Online' if remote",
  "meetingType": "Classify: Sales Call, Client Briefing, Project Review, Status Update, Kick-off, Discovery Call, etc.",
  "attendees": "All attendees, one per line starting with •. Include name and role/company if known.",
  "clientAttendees": "Client-side attendees only, one per line",
  "agencyAttendees": "Agency/host-side attendees only, one per line",
  "objective": "2-3 sentence summary of the meeting purpose and context.",
  "discussionPoints": "Numbered list of main topics discussed. One per line. Be specific.",
  "decisions": "Numbered list of decisions or agreements reached. If none: 'No formal decisions recorded.'",
  "actionItems": "Numbered list. Format: [Owner Name] — Task description (Due: date or ASAP)",
  "nextSteps": "Numbered list of agreed next steps in sequence.",
  "followUpDate": "Next meeting or follow-up date if mentioned, otherwise 'To be confirmed'",
  "additionalNotes": "Important context, risks, concerns not captured above. If none: 'No additional notes.'"
}`
        }]
      })
    });

    const claudeData = await claudeRes.json();
    if (!claudeRes.ok) {
      const msg = claudeData?.error?.message || `Claude error ${claudeRes.status}`;
      return res.status(502).json({ error: 'Analysis failed: ' + msg });
    }

    const rawText = (claudeData.content || [])
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('');

    const cleaned = rawText.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const fields  = JSON.parse(cleaned);

    // Build plain-text report
    const divider = '─'.repeat(60);
    const generatedDate = new Date().toLocaleDateString('en-GB', { day:'2-digit', month:'long', year:'numeric' });
    const reportText = [
      'CONTACT REPORT',
      divider,
      '',
      `Client:           ${fields.client       || 'N/A'}`,
      `Meeting Date:     ${fields.date         || 'N/A'}`,
      `Meeting Time:     ${fields.time         || 'N/A'}`,
      `Duration:         ${fields.duration     || 'N/A'}`,
      `Venue:            ${fields.venue        || 'N/A'}`,
      `Meeting Type:     ${fields.meetingType  || 'N/A'}`,
      '',
      divider, 'ATTENDEES', divider,
      fields.attendees || 'N/A',
      '',
      divider, 'MEETING OBJECTIVE', divider,
      fields.objective || 'N/A',
      '',
      divider, 'KEY DISCUSSION POINTS', divider,
      fields.discussionPoints || 'N/A',
      '',
      divider, 'DECISIONS MADE', divider,
      fields.decisions || 'N/A',
      '',
      divider, 'ACTION ITEMS', divider,
      fields.actionItems || 'N/A',
      '',
      divider, 'NEXT STEPS', divider,
      fields.nextSteps || 'N/A',
      '',
      divider,
      `FOLLOW-UP DATE:   ${fields.followUpDate || 'N/A'}`,
      '',
      divider, 'ADDITIONAL NOTES', divider,
      fields.additionalNotes || 'N/A',
      '',
      divider,
      `Generated by Contact Report Generator · ${generatedDate}`,
    ].join('\n');

    res.json({ fields, reportText });

  } catch (err) {
    console.error('[/api/analyse]', err);
    if (err instanceof SyntaxError) {
      return res.status(502).json({ error: 'Failed to parse Claude response as JSON. Please try again.' });
    }
    res.status(500).json({ error: 'Analysis request failed: ' + err.message });
  }
});

// ── POST /api/export-docx ──────────────────────────────────────────────────
// Accepts: JSON { fields: object }
// Returns: .docx file download
app.post('/api/export-docx', async (req, res) => {
  const { fields } = req.body;
  if (!fields || typeof fields !== 'object') {
    return res.status(400).json({ error: 'No fields provided.' });
  }

  try {
    const buffer = await buildDocx(fields);
    const date   = fields.date ? fields.date.replace(/[^a-zA-Z0-9-]/g, '-') : 'contact-report';
    const filename = `Contact-Report-${date}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error('[/api/export-docx]', err);
    res.status(500).json({ error: 'Failed to generate DOCX: ' + err.message });
  }
});

// ── Health check ───────────────────────────────────────────────────────────
app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ── Catch-all → index.html ─────────────────────────────────────────────────
app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use((err, _req, res, _next) => {
  if (err.code === 'LIMIT_FILE_SIZE') return res.status(413).json({ error: 'File too large. Maximum size is 50 MB.' });
  console.error('[unhandled]', err);
  res.status(500).json({ error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`\n  Contact Report Generator`);
  console.log(`  ─────────────────────────`);
  console.log(`  Running at http://localhost:${PORT}`);
  console.log(`  Gemini (transcription) : ✓`);
  console.log(`  Claude  (analysis)     : ✓`);
  console.log(`  DOCX export            : ✓\n`);
});
