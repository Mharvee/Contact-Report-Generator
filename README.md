# Contact Report Generator

Turn meeting recordings into professional Contact Reports.
**Gemini Flash** transcribes the audio (free tier). **Claude** analyses the transcript and generates the report. All API keys stay on the server — never exposed to the client.

---

## Quick Start

### 1. Install dependencies
```bash
npm install
```

### 2. Add your API keys
```bash
cp .env.example .env
```
Edit `.env` and fill in both keys:
```
ANTHROPIC_API_KEY=sk-ant-...
GEMINI_API_KEY=AIza...
```

**Get your keys:**
- Anthropic: https://console.anthropic.com
- Gemini (free): https://aistudio.google.com/app/apikey

### 3. Run
```bash
npm start
```
Open http://localhost:3000

---

## Development (auto-restart on save)
```bash
npm run dev
```

---

## How it works

```
Browser → POST /api/transcribe (multipart audio)
            → Gemini Flash 1.5 (free tier)
            ← transcript text

Browser → POST /api/analyse (JSON transcript)
            → Claude claude-opus-4-5
            ← { fields, reportText }
```

1. User drops an audio/video file in the browser
2. Server sends the file to **Gemini 1.5 Flash** for transcription
3. Transcript is sent to **Claude** for structured extraction
4. Claude returns 12 fields (attendees, decisions, action items, etc.)
5. Fields are assembled into a formatted Contact Report
6. User can view, edit, copy, or download the report as `.txt`

---

## Supported formats
MP3, MP4, M4A, WAV, OGG, WebM, FLAC — up to 50 MB

---

## Deployment

### Environment variables required
| Variable | Description |
|----------|-------------|
| `ANTHROPIC_API_KEY` | Your Anthropic API key |
| `GEMINI_API_KEY` | Your Google Gemini API key |
| `PORT` | Server port (default: 3000) |

### Deploy to Railway / Render / Fly.io
Set the two environment variables in your platform's dashboard and deploy. The app serves the frontend as static files from `/public`.

### Deploy to a VPS
```bash
# Install Node 18+, then:
npm install --production
PORT=3000 node server.js

# Or use PM2 for process management:
npm install -g pm2
pm2 start server.js --name contact-report
pm2 save
```

---

## File structure
```
contact-report/
├── server.js          ← Express backend + API proxy
├── .env               ← Your API keys (never commit this)
├── .env.example       ← Key template
├── package.json
└── public/
    └── index.html     ← Frontend SPA
```
