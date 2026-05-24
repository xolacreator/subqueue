// ═══════════════════════════════════════════════════════════════
// Subqueue – Peptide Drive Agent (Google Apps Script)
// Deploy as: Extensions → Apps Script → Deploy → New deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

// ── Config ──────────────────────────────────────────────────────
// Store secrets via Script Properties (recommended — never hardcode keys):
//   Apps Script editor → Project Settings (⚙) → Script Properties → Add:
//     ANTHROPIC_API_KEY  =  sk-ant-...
//     DRIVE_FOLDER_ID    =  (folder ID from Drive URL)
//
// Alternatively set them here directly (less secure — visible in script history):
const ANTHROPIC_API_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || 'sk-ant-YOUR-KEY-HERE';
const DRIVE_FOLDER_ID   = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID')   || 'YOUR_FOLDER_ID_HERE';
const AI_MODEL          = 'claude-sonnet-4-6';
const MAX_DOC_CHARS     = 40000;  // chars per document included in context

// ── Entry points ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const peptide = (body.peptide || '').trim();
    if (!peptide) return jsonResponse({ success: false, error: 'No peptide name provided' });

    const context = buildContext(peptide);
    const data    = callClaude(peptide, context);
    return jsonResponse({ success: true, data });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// Allow preflight CORS from any origin (browsers send OPTIONS first)
function doGet(e) {
  return jsonResponse({ ok: true, message: 'Subqueue peptide agent is running' });
}

// ── Drive context builder ─────────────────────────────────────────
function buildContext(peptide) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const chunks  = [];
  let   total   = 0;

  // Walk all files in folder (PDFs, Docs, Sheets)
  const files = folder.getFiles();
  while (files.hasNext() && total < MAX_DOC_CHARS * 5) {
    const file = files.next();
    const mime  = file.getMimeType();
    let   text  = '';

    try {
      if (mime === MimeType.GOOGLE_DOCS) {
        text = DocumentApp.openById(file.getId()).getBody().getText();
      } else if (mime === MimeType.GOOGLE_SHEETS) {
        text = extractSheetText(SpreadsheetApp.openById(file.getId()));
      } else if (mime === MimeType.PDF) {
        // PDFs: convert to Google Doc first for text extraction
        const copy = Drive.Files.copy(
          { mimeType: MimeType.GOOGLE_DOCS },
          file.getId()
        );
        const doc  = DocumentApp.openById(copy.id);
        text       = doc.getBody().getText();
        DriveApp.getFileById(copy.id).setTrashed(true); // clean up temp
      }
    } catch (err) {
      // Skip unreadable files
      continue;
    }

    if (!text) continue;

    // Only include docs that mention the peptide (case-insensitive)
    if (text.toLowerCase().includes(peptide.toLowerCase())) {
      const trimmed = text.slice(0, MAX_DOC_CHARS);
      chunks.push(`--- ${file.getName()} ---\n${trimmed}`);
      total += trimmed.length;
    }
  }

  return chunks.join('\n\n');
}

function extractSheetText(ss) {
  const parts = [];
  ss.getSheets().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    parts.push(sheet.getName() + ':\n' + data.map(row => row.join('\t')).join('\n'));
  });
  return parts.join('\n\n');
}

// ── Anthropic call ────────────────────────────────────────────────
function callClaude(peptide, driveContext) {
  const systemPrompt = driveContext
    ? `You are a peptide pharmacology expert. You have access to the following reference documents from the user's knowledge base:\n\n${driveContext}\n\nUse this information as your primary source. Supplement with your training knowledge where the documents are silent.`
    : 'You are a peptide pharmacology expert with extensive knowledge of research peptides.';

  const userPrompt = `Generate complete, accurate data for the peptide: "${peptide}".

Return ONLY valid JSON with no markdown, no explanation, just the JSON object:
{
  "name": "string",
  "aka": "string (aliases / brand names)",
  "cat": "GH Secretagogue|Healing|Metabolic|Mitochondrial|Nootropic|Other",
  "tags": ["string"],
  "summary": "2-3 sentence clinical summary",
  "dosing": {
    "typical": "string",
    "range": "string",
    "freq": "string",
    "timing": "string (e.g. fasted, pre-sleep)",
    "cycle": "string (e.g. 12 weeks on, 4 weeks off)"
  },
  "recon": {
    "vial": "string (e.g. 5mg)",
    "dil": "string (e.g. 2mL bacteriostatic water)",
    "conc": "string (e.g. 2500mcg/mL)",
    "tbl": "5 units = 250mcg\\n10 units = 500mcg",
    "exp": 28
  },
  "def": {
    "dose": 250,
    "unit": "mcg",
    "freq": "5on2off",
    "time": "20:00"
  },
  "notes": "string (storage, synergies, stacking notes)",
  "warn": "string (contraindications, side effects – empty string if none)"
}`;

  const payload = {
    model: AI_MODEL,
    max_tokens: 1024,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body   = response.getContentText();

  if (status !== 200) {
    throw new Error(`Anthropic API error ${status}: ${body}`);
  }

  const msg = JSON.parse(body);
  const raw = msg.content?.[0]?.text || '';

  // Robust JSON extraction (find first { to last })
  const start = raw.indexOf('{');
  const end   = raw.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('No JSON found in AI response');
  return JSON.parse(raw.slice(start, end + 1));
}

// ── Helpers ───────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
