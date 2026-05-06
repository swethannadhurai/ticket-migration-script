require('dotenv').config();
const axios = require('axios');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// 🔴 EDIT THIS LINE: Put the Ticket IDs from your Test Account here!
const TEST_TICKET_IDS = [1330];

// --- Config ---
const FRESHDESK_DOMAIN = process.env.FRESHDESK_DOMAIN;
const FRESHDESK_API_KEY = process.env.FRESHDESK_API_KEY;
const EXCEL_FOLDER = process.env.EXCEL_FOLDER || '';
const CF_DATE = (process.env.TICKET_CREATION_DATE_FIELD || '').trim();
const DRY_RUN = process.env.DRY_RUN === 'true';

const API_BASE = `https://${FRESHDESK_DOMAIN}/api/v2`;
const AUTH_HEADER = `Basic ${Buffer.from(`${FRESHDESK_API_KEY}:X`).toString('base64')}`;

// ── Utils ─────────────────────────────────────────────────────────────────────

function log(level, msg, data = '') {
  console.log(`[${level.toUpperCase()}] ${msg}`, data ? data : '');
}

async function apiGet(endpoint) {
  const res = await axios.get(`${API_BASE}${endpoint}`, { headers: { Authorization: AUTH_HEADER } });
  return res.data;
}

async function apiDeleteNote(ticketId, noteId) {
  if (DRY_RUN) { log('dry', `Delete note ${noteId} on ticket #${ticketId}`); return; }
  await axios.delete(`${API_BASE}/conversations/${noteId}`, {
    headers: { Authorization: AUTH_HEADER }
  });
}

async function apiCreatePublicNote(ticketId, bodyHtml) {
  if (DRY_RUN) { log('dry', `Create PUBLIC note on ticket #${ticketId}`); return; }
  await axios.post(`${API_BASE}/tickets/${ticketId}/notes`, {
    body: bodyHtml,
    private: false
  }, { headers: { Authorization: AUTH_HEADER } });
}

async function apiCreatePrivateNote(ticketId, bodyHtml) {
  if (DRY_RUN) { log('dry', `Create PRIVATE note on ticket #${ticketId}`); return; }
  await axios.post(`${API_BASE}/tickets/${ticketId}/notes`, {
    body: bodyHtml,
    private: true
  }, { headers: { Authorization: AUTH_HEADER } });
}

function stripHtml(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>?/gm, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function noteMatches(bodyText, sourceText) {
  const normNote = bodyText.toLowerCase().replace(/[^a-z0-9]/g, '');
  const normSource = sourceText.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!normNote || !normSource || normNote.length < 15 || normSource.length < 15) return false;
  const checkLen = Math.min(50, normSource.length);
  const prefixSource = normSource.substring(0, checkLen);
  const midSource = normSource.substring(Math.floor(normSource.length / 2), Math.floor(normSource.length / 2) + checkLen);
  return normNote.includes(prefixSource) || normNote.includes(midSource) || normSource.includes(normNote);
}

/** Split a combined notes block into individual entries by timestamp prefix.
 *  Pattern: DD/MM/YYYY HH:MM:SS - Author (Type)  at the start of a line.
 */
function splitIntoBlocks(text) {
  // Split only at positions where a timestamp starts at the beginning of a line
  // e.g. "12/09/2024 08:27:14 - Akbar Sharif I (Customer Visible Notes)"
  return text
    .split(/\n(?=\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2} - )/)
    .map(s => s.trim())
    .filter(Boolean);
}

// ── Excel ─────────────────────────────────────────────────────────────────────

function readExcelFile(filePath, targetSheetIndex = 0) {
  if (!fs.existsSync(filePath)) return [];
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[targetSheetIndex];
  if (!sheetName) return [];
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
}

function buildExcelLookup(rows) {
  const map = new Map();
  for (const row of rows) {
    const num = String(row['Number'] || row['number'] || '').trim();
    if (num) map.set(num.toUpperCase(), row);
    const sub = String(row['Subject'] || row['subject'] || '').trim();
    if (sub && sub !== num) map.set(sub.toUpperCase(), row);
  }
  return map;
}

function getCreatedDate(row) {
  const val = row['Created'] || row['created'] || row['Opened'] || row['opened'] || '';
  if (!val) return null;

  function formatDate(d) {
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    const hours = String(d.getUTCHours()).padStart(2, '0');
    const minutes = String(d.getUTCMinutes()).padStart(2, '0');
    const seconds = String(d.getUTCSeconds()).padStart(2, '0');
    return `${day}-${month}-${year} ${hours}:${minutes}:${seconds}`;
  }

  if (typeof val === 'number') {
    try {
      const d = xlsx.SSF.parse_date_code(val);
      const dt = new Date(Date.UTC(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0));
      return formatDate(dt);
    } catch (_) { }
  }
  const d = new Date(String(val).trim());
  if (!isNaN(d.valueOf())) return formatDate(d);
  return null;
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  log('info', 'Starting Test Environment Migration Check...');
  log('info', `DOMAIN: ${FRESHDESK_DOMAIN}`);
  log('info', `DRY_RUN: ${DRY_RUN}`);

  // 1. Load Excel
  const excelRows = [
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Service Request.xlsx'), 1),
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Customerservice_case.xlsx'), 1),
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Master Ticket (All Status).xlsx'), 0)
  ];
  const excelLookup = buildExcelLookup(excelRows);
  log('info', `Loaded ${excelRows.length} total target rows from Excel.`);

  // 2. Process each ticket
  for (const ticketId of TEST_TICKET_IDS) {
    console.log(`\n── Checking Test Ticket #${ticketId} ──`);
    try {
      const ticket = await apiGet(`/tickets/${ticketId}`);
      const subject = ticket.subject || '';

      let externalRef = '';
      if (ticket.custom_fields) {
        const extKey = Object.keys(ticket.custom_fields).find(key => key.startsWith('cf_external_reference'));
        if (extKey) externalRef = ticket.custom_fields[extKey] || '';
      }

      log('info', `Live Details:`, { subject, externalRef });

      // ── Match Excel row ──────────────────────────────────────────────────────
      let excelRow = null;
      if (externalRef) excelRow = excelLookup.get(externalRef.toUpperCase().trim());
      if (!excelRow) excelRow = excelLookup.get(subject.toUpperCase().trim());
      if (!excelRow) {
        for (const [key, row] of excelLookup) {
          if (subject.toUpperCase().includes(key) || key.includes(subject.toUpperCase())) {
            excelRow = row; break;
          }
        }
      }

      if (!excelRow) {
        log('warn', `NO EXCEL MATCH for Test Ticket #${ticketId}`);
        continue;
      }

      // ── Date diagnostics ─────────────────────────────────────────────────────
      const dateVal = excelRow['Created'] || excelRow['created'] || excelRow['Opened'] || '';
      log('info', `  Raw date value: "${dateVal}"`);

      // ── Date update ──────────────────────────────────────────────────────────
      const createdDate = getCreatedDate(excelRow);
      if (!createdDate) {
        log('warn', `  ⚠️  No parseable date found in Excel for ticket #${ticketId} — skipping date update`);
      } else if (DRY_RUN) {
        log('dry', `  Would update Ticket ${ticketId} ${CF_DATE} → ${createdDate}`);
      } else {
        let payloadFields = { [CF_DATE]: createdDate };
        let success = false;
        const maxRetries = 3;

        for (let i = 0; i < maxRetries && !success; i++) {
          try {
            await axios.put(`${API_BASE}/tickets/${ticketId}`, {
              custom_fields: payloadFields
            }, { headers: { Authorization: AUTH_HEADER } });
            log('info', `  ✅ Date updated: ${createdDate}`);
            success = true;
          } catch (dateErr) {
            if (dateErr.response && dateErr.response.data && dateErr.response.data.errors && i < maxRetries - 1) {
              let handled = false;
              for (const e of dateErr.response.data.errors) {
                if (e.field && e.field.startsWith('custom_fields.')) {
                  const badField = e.field.replace('custom_fields.', '');
                  delete payloadFields[badField];
                  handled = true;
                  log('warn', `  ⚠️  Dropping invalid field: ${badField}. Retrying without it...`);
                }
              }
              if (!handled) {
                log('error', `  ❌ Date update failed: ${dateErr.message}`);
                console.log('  Date API Error:', JSON.stringify(dateErr.response.data, null, 2));
                break;
              }
            } else {
              log('error', `  ❌ Date update failed: ${dateErr.message}`);
              if (dateErr.response && dateErr.response.data) {
                console.log('  Date API Error:', JSON.stringify(dateErr.response.data, null, 2));
              }
              break;
            }
          }
        }
      }

      // ── Notes logic ──────────────────────────────────────────────────────────
      const cvnText = String(excelRow['Customer Visible Notes'] || excelRow['Customer visible notes'] || excelRow['customer visible notes'] || '').trim();
      const intText = String(excelRow['Internal Notes'] || excelRow['Internal notes'] || excelRow['internal notes'] || '').trim();

      if (!cvnText && !intText) {
        log('info', `  No source notes in Excel for ticket #${ticketId}`);
        continue;
      }

      // Pre-split into individual blocks by timestamp
      const cvnBlocks = cvnText ? splitIntoBlocks(cvnText) : [];
      const intBlocks = intText ? splitIntoBlocks(intText) : [];

      const convos = await apiGet(`/tickets/${ticketId}/conversations`);
      console.log(`  Notes — API: ${convos.length} | CVN blocks: ${cvnBlocks.length} | INT blocks: ${intBlocks.length}`);

      let needsCvn = false;
      let needsInt = false;
      const deletedIds = new Set();
      let hasCvnNote = false;  // did we find ANY live note matching CVN?
      let hasIntNote = false;  // did we find ANY live note matching INT?

      for (const convo of convos) {
        if (deletedIds.has(convo.id)) continue; // already deleted this run
        const isPrivate = convo.private === true;
        const bodyText = convo.body_text || stripHtml(convo.body) || '';

        const containsCvn = cvnText && noteMatches(bodyText, cvnText);
        const containsInt = intText && noteMatches(bodyText, intText);

        if (!containsCvn && !containsInt) continue;

        // ── CVN check ──────────────────────────────────────────────────────────
        if (containsCvn) {
          hasCvnNote = true;
          // BAD if: private, combined with INT, OR source has >1 block (needs splitting into separate notes)
          const isBad = isPrivate || containsInt || cvnBlocks.length > 1;
          if (isBad) {
            const reason = isPrivate ? 'wrong visibility (private)' : containsInt ? 'combined with Internal Notes' : `needs splitting (${cvnBlocks.length} blocks in source)`;
            console.log(`  🔄 Note ${convo.id} is BAD — ${reason}. Deleting...`);
            try {
              await apiDeleteNote(ticketId, convo.id);
              deletedIds.add(convo.id);
              needsCvn = true;
              if (containsInt) needsInt = true;
            } catch (e) { log('error', `Delete failed: ${e.message}`); }
          } else {
            console.log(`  ✅ Note ${convo.id} is already a PERFECT public Customer Visible Note.`);
          }
        }

        // ── INT check ──────────────────────────────────────────────────────────
        if (containsInt && !(containsCvn && isPrivate)) {
          hasIntNote = true;
          const isBad = !isPrivate || containsCvn || intBlocks.length > 1;
          if (isBad) {
            const reason = !isPrivate ? 'wrong visibility (public)' : containsCvn ? 'combined with Customer Visible Notes' : `needs splitting (${intBlocks.length} blocks in source)`;
            console.log(`  🔄 Note ${convo.id} is BAD — ${reason}. Deleting...`);
            try {
              await apiDeleteNote(ticketId, convo.id);
              deletedIds.add(convo.id);
              needsInt = true;
            } catch (e) { log('error', `Delete failed: ${e.message}`); }
          } else {
            console.log(`  ✅ Note ${convo.id} is already a PERFECT private Internal Note.`);
          }
        }
      }

      // Also create notes that are completely missing from the live ticket
      if (cvnText && !hasCvnNote) { needsCvn = true; log('info', `  ℹ️  CVN notes are missing entirely — will create.`); }
      if (intText && !hasIntNote) { needsInt = true; log('info', `  ℹ️  INT notes are missing entirely — will create.`); }

      // ── Recreate: one separate note per block, oldest first ─────────────────
      if (needsCvn && cvnBlocks.length > 0) {
        const ordered = [...cvnBlocks].reverse(); // oldest first
        for (const block of ordered) {
          const html = `<p>${block.replace(/\n/g, '<br>')}</p>`;
          try {
            await apiCreatePublicNote(ticketId, html);
            console.log(`  ✨ Created PUBLIC note: "${block.substring(0, 80).replace(/\n/g, ' ')}"`);
          } catch (e) { log('error', `Create public note failed: ${e.message}`); }
        }
      }

      if (needsInt && intBlocks.length > 0) {
        const ordered = [...intBlocks].reverse(); // oldest first
        for (const block of ordered) {
          const html = `<p>${block.replace(/\n/g, '<br>')}</p>`;
          try {
            await apiCreatePrivateNote(ticketId, html);
            console.log(`  ✨ Created PRIVATE note: "${block.substring(0, 80).replace(/\n/g, ' ')}"`);
          } catch (e) { log('error', `Create private note failed: ${e.message}`); }
        }
      }

    } catch (err) {
      log('error', `Failed processing ticket ${ticketId}`, err.message);
      if (err.response && err.response.data) {
        console.log('API Error Details:', JSON.stringify(err.response.data, null, 2));
      }
    }
  }
}

main().catch(console.error);
