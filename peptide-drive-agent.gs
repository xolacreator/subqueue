// ═══════════════════════════════════════════════════════════════
// Subqueue – Peptide Drive Agent (Google Apps Script)
// Deploy as: Extensions → Apps Script → Deploy → New deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
//
// Script Properties (Project Settings ⚙ → Script Properties):
//   ANTHROPIC_API_KEY  =  sk-ant-...
//   DRIVE_FOLDER_IDS   =  folderId1,folderId2,folderId3   (comma-separated, or leave blank to search all Drive)
// ═══════════════════════════════════════════════════════════════

const AI_MODEL      = 'claude-sonnet-4-6';
const MAX_DOC_CHARS = 40000;  // chars per document included in context
const MAX_TOTAL     = 120000; // total context chars across all docs

// ── Entry points ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.peptide) {
      // Research mode: generate full library entry for a named peptide
      const context = buildContext(body.peptide);
      const data    = callClaudeResearch(body.peptide, context);
      return jsonResponse({ success: true, data });

    } else if (body.text) {
      // Extract mode: parse protocols/doses from pasted text
      const data = callClaudeExtract(body.text);
      return jsonResponse({ success: true, data });

    } else {
      return jsonResponse({ success: false, error: 'Send {peptide:"name"} or {text:"..."}' });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: 'Subqueue peptide agent running' });
}

// ── Drive context builder ─────────────────────────────────────────
function buildContext(keyword) {
  const props     = PropertiesService.getScriptProperties();
  const apiKey    = props.getProperty('ANTHROPIC_API_KEY');
  const folderIds = (props.getProperty('DRIVE_FOLDER_IDS') || '').split(',').map(s => s.trim()).filter(Boolean);

  const files   = folderIds.length ? getFilesFromFolders(folderIds) : searchDriveFiles(keyword);
  const chunks  = [];
  let   total   = 0;

  for (const file of files) {
    if (total >= MAX_TOTAL) break;
    const text = extractText(file);
    if (!text) continue;
    if (keyword && !text.toLowerCase().includes(keyword.toLowerCase())) continue;
    const trimmed = text.slice(0, MAX_DOC_CHARS);
    chunks.push(`--- ${file.getName()} ---\n${trimmed}`);
    total += trimmed.length;
  }

  return chunks.join('\n\n');
}

function getFilesFromFolders(folderIds) {
  const files = [];
  for (const id of folderIds) {
    try {
      const folder = DriveApp.getFolderById(id);
      // Also recurse one level into sub-folders
      const iter = folder.getFiles();
      while (iter.hasNext()) files.push(iter.next());
      const subs = folder.getFolders();
      while (subs.hasNext()) {
        const sub = subs.next();
        const subIter = sub.getFiles();
        while (subIter.hasNext()) files.push(subIter.next());
      }
    } catch (err) { /* skip inaccessible folder */ }
  }
  return files;
}

function searchDriveFiles(keyword) {
  // Full-text search across all Drive if no folders configured
  const query = keyword
    ? `fullText contains '${keyword.replace(/'/g, "\\'")}' and trashed = false`
    : `(mimeType='application/pdf' or mimeType='application/vnd.google-apps.document' or mimeType='application/vnd.google-apps.spreadsheet') and trashed = false`;
  const result = [];
  try {
    const iter = DriveApp.searchFiles(query);
    while (iter.hasNext() && result.length < 20) result.push(iter.next());
  } catch (err) { /* ignore */ }
  return result;
}

function extractText(file) {
  const mime = file.getMimeType();
  try {
    if (mime === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(file.getId()).getBody().getText();
    }
    if (mime === MimeType.GOOGLE_SHEETS) {
      return extractSheetText(SpreadsheetApp.openById(file.getId()));
    }
    if (mime === MimeType.PDF) {
      const copy = Drive.Files.copy({ mimeType: MimeType.GOOGLE_DOCS }, file.getId());
      const text = DocumentApp.openById(copy.id).getBody().getText();
      DriveApp.getFileById(copy.id).setTrashed(true);
      return text;
    }
  } catch (err) { /* skip unreadable */ }
  return '';
}

function extractSheetText(ss) {
  return ss.getSheets().map(sheet => {
    const rows = sheet.getDataRange().getValues();
    return sheet.getName() + ':\n' + rows.map(r => r.join('\t')).join('\n');
  }).join('\n\n');
}

// ── Claude calls ──────────────────────────────────────────────────
function callClaudeResearch(peptide, driveContext) {
  const systemPrompt = driveContext
    ? `You are a peptide pharmacology expert. Primary source — the user's reference documents:\n\n${driveContext}\n\nSupplement with your training knowledge where the documents are silent.`
    : 'You are a peptide pharmacology expert with extensive knowledge of research peptides.';

  const userPrompt = `Generate complete, accurate data for the peptide: "${peptide}".

Return ONLY valid JSON with no markdown:
{
  "name": "string",
  "aka": "string",
  "cat": "GH Secretagogue|Healing|Metabolic|Mitochondrial|Nootropic|Other",
  "tags": ["string"],
  "summary": "2-3 sentence clinical summary",
  "dosing": {
    "typical": "string",
    "range": "string",
    "freq": "string",
    "timing": "string",
    "cycle": "string"
  },
  "recon": {
    "vial": "string",
    "dil": "string",
    "conc": "string",
    "tbl": "5 units = 250mcg\\n10 units = 500mcg",
    "exp": 28
  },
  "def": { "dose": 250, "unit": "mcg", "freq": "5on2off", "time": "20:00" },
  "notes": "string",
  "warn": "string"
}`;

  return callClaude(systemPrompt, userPrompt, 1024);
}

function callClaudeExtract(text) {
  const systemPrompt = 'You are a medical protocol data extractor. Extract all medication and peptide dosing information accurately.';
  const userPrompt = `Extract all medication/peptide dosing protocols from the text below. Return ONLY valid JSON, no markdown:
{"protocols":[{"name":"string","dose":number_or_null,"unit":"mcg|mg|IU|mL|units or null","frequency":"daily|5on2off|weekly|eod|twice_weekly|monthly or null","vialSizeMg":number_or_null,"unitTable":"newline-separated or null","timing":"string or null","notes":"string or null","warnings":"string or null","confidence":0.0}],"summary":"brief"}

Text:
${text.slice(0, 6000)}`;

  return callClaude(systemPrompt, userPrompt, 2048);
}

function callClaude(system, user, maxTokens) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY not set in Script Properties');

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: AI_MODEL,
      max_tokens: maxTokens,
      system,
      messages: [{ role: 'user', content: user }]
    }),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body   = response.getContentText();
  if (status !== 200) throw new Error(`Anthropic API error ${status}: ${body}`);

  const msg = JSON.parse(body);
  const raw = (msg.content || []).map(c => c.text || '').join('');
  const s   = raw.indexOf('{'), e = raw.lastIndexOf('}');
  if (s === -1 || e === -1) throw new Error('No JSON in AI response: ' + raw.slice(0, 200));
  return JSON.parse(raw.slice(s, e + 1));
}

// ── Helpers ───────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
