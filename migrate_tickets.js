/**
 * migrate_tickets.js
 * ─────────────────────────────────────────────────────────────────
 * For each Freshdesk ticket tagged with both "Prod_migration" AND "migrated":
 *
 *   1. Match the ticket to an Excel row via ServiceNow ticket number
 *      (XML subject → Excel "Number" column)
 *   2. Update "Ticket Creation Date" custom field from Excel "Created" column
 *   3. Fix notes:
 *        - Find XML notes whose content matches an Excel source note
 *        - DELETE those duplicate notes from Freshdesk
 *        - RE-CREATE them with correct visibility:
 *            "Customer Visible Notes" → public reply
 *            "Internal Notes"        → private note
 *        - Leave any post-migration notes (not in Excel) untouched
 *
 * Usage:
 *   1. Fill in .env with your Freshdesk credentials
 *   2. Set DRY_RUN=true first to preview actions without making any changes
 *   3. Run: node migrate_tickets.js
 * ─────────────────────────────────────────────────────────────────
 */

'use strict';
require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const xml2js = require('xml2js');
const axios = require('axios');

// ── Config ────────────────────────────────────────────────────────
const DOMAIN = (process.env.FRESHDESK_DOMAIN || '').trim();
const API_KEY = (process.env.FRESHDESK_API_KEY || '').trim();
const CF_DATE = 'cf_original_ticket_creation_date';
const DRY_RUN = process.env.DRY_RUN === 'true';
const XML_FOLDER = process.env.XML_FOLDER;
const EXCEL_FOLDER = process.env.EXCEL_FOLDER;

// Tags that must BOTH be present on a ticket (case-insensitive)
const REQUIRED_TAGS = ['prod_migration', 'migrated'];

// 🔴 EDIT THIS LINE: To test specific tickets in Prod, add their IDs here (e.g. ['622', '624']).
// Leave it completely empty `[]` to process ALL tickets.
const TEST_TICKET_IDS = ['1330'];

const RATE_LIMIT_MS = 1300;   // ~45 req/min (Freshdesk limit: 50/min)
const PROGRESS_FILE = path.join(__dirname, 'progress.json');
const LOG_FILE = path.join(__dirname, 'migration_log.json');

// ── Utilities ─────────────────────────────────────────────────────

const sleep = ms => new Promise(r => setTimeout(r, ms));

/** Strip all HTML and normalise whitespace for fuzzy text comparison */
function stripHtml(html = '') {
  return String(html || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&#?\w+;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Load progress file (for resuming interrupted runs) */
function loadProgress() {
  if (fs.existsSync(PROGRESS_FILE)) {
    return JSON.parse(fs.readFileSync(PROGRESS_FILE, 'utf8'));
  }
  return { done: [], failed: [] };
}

function saveProgress(p) {
  fs.writeFileSync(PROGRESS_FILE, JSON.stringify(p, null, 2));
}

const logEntries = [];
function log(level, msg, data = {}) {
  const entry = { time: new Date().toISOString(), level, msg, ...data };
  const dataStr = Object.keys(data).length ? JSON.stringify(data) : '';
  console.log(`[${level.toUpperCase().padEnd(5)}] ${msg} ${dataStr}`);
  logEntries.push(entry);
}

function saveLog() {
  fs.writeFileSync(LOG_FILE, JSON.stringify(logEntries, null, 2));
}

// ── Freshdesk API ─────────────────────────────────────────────────

const fd = axios.create({
  baseURL: `https://${DOMAIN}/api/v2`,
  auth: { username: API_KEY, password: 'X' },
  headers: { 'Content-Type': 'application/json' },
});

async function apiUpdateTicket(displayId, customFields) {
  if (DRY_RUN) {
    log('dry', `[DRY] Update ticket #${displayId}`, { customFields });
    return;
  }
  await sleep(RATE_LIMIT_MS);
  await fd.put(`/tickets/${displayId}`, { custom_fields: customFields });
  log('info', `Updated ticket #${displayId}`, { customFields });
}

async function apiDeleteConversation(noteId) {
  if (DRY_RUN) {
    log('dry', `[DRY] Delete conversation ${noteId}`);
    return;
  }
  await sleep(RATE_LIMIT_MS);
  await fd.delete(`/conversations/${noteId}`);
  log('info', `Deleted conversation ${noteId}`);
}

async function apiGetConversations(displayId) {
  await sleep(RATE_LIMIT_MS);
  const res = await fd.get(`/tickets/${displayId}/conversations`);
  return res.data || [];
}

async function apiCreatePrivateNote(displayId, body) {
  if (DRY_RUN) {
    log('dry', `[DRY] Create PRIVATE note on ticket #${displayId}`);
    return;
  }
  await sleep(RATE_LIMIT_MS);
  await fd.post(`/tickets/${displayId}/notes`, { private: true, body });
  log('info', `Created PRIVATE note on ticket #${displayId}`);
}

async function apiCreatePublicNote(displayId, body) {
  if (DRY_RUN) {
    log('dry', `[DRY] Create PUBLIC note on ticket #${displayId}`);
    return;
  }
  await sleep(RATE_LIMIT_MS);
  await fd.post(`/tickets/${displayId}/notes`, { private: false, body });
  log('info', `Created PUBLIC note on ticket #${displayId}`);
}

// ── Excel Reader ──────────────────────────────────────────────────

/**
 * Read ALL rows from the FIRST sheet of an Excel file.
 * (Local exports contain all rows in a single sheet even if Google Sheets
 *  shows multiple tabs.)
 */
function readExcelFile(filePath, targetSheetIndex = 0) {
  if (!fs.existsSync(filePath)) {
    log('warn', `Excel file not found: ${path.basename(filePath)}`);
    return [];
  }
  log('info', `Loading Excel: ${path.basename(filePath)}`);
  const wb = xlsx.readFile(filePath, { cellDates: true });
  if (targetSheetIndex >= wb.SheetNames.length) {
    log('warn', `Requested sheet index ${targetSheetIndex} out of bounds for ${path.basename(filePath)}`);
    return [];
  }
  const sheetName = wb.SheetNames[targetSheetIndex];
  log('info', `  Using sheet: "${sheetName}"`);
  const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  log('info', `  Rows loaded: ${rows.length}`);
  return rows;
}

/**
 * Build a lookup map: ServiceNow ticket number (upper-cased) → Excel row
 *
 * Tries both "Number" and "Subject" columns because:
 *   - Service Request.xlsx : Number = RITM... (matches Freshdesk subject)
 *   - CS file              : Number = CUST..., Subject = INC... (INC matches Freshdesk subject)
 */
function buildExcelLookup(rows) {
  const map = new Map();

  for (const row of rows) {
    // Primary key: Number column
    const num = String(row['Number'] || row['number'] || '').trim();
    if (num) map.set(num.toUpperCase(), row);

    // Secondary key: Subject column (for CS tickets where subject = INC number)
    const sub = String(row['Subject'] || row['subject'] || '').trim();
    if (sub && sub !== num) map.set(sub.toUpperCase(), row);
  }

  log('info', `Excel lookup built: ${map.size} keys`);
  return map;
}

/** Extract the "Created" date from an Excel row as an ISO string */
function getCreatedDate(row) {
  const val =
    row['Created'] ||
    row['created'] ||
    row['Opened'] ||
    row['opened'] ||
    '';

  if (!val) return null;

  function formatDate(year, month, day, hours, minutes) {
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:00.000Z`;
  }

  if (val instanceof Date) {
    let year = val.getFullYear();
    let month = val.getMonth() + 1;
    let day = val.getDate();
    let hours = val.getHours();
    let minutes = val.getMinutes();

    if (val.getSeconds() >= 30) {
      minutes += 1;
    }

    if (minutes === 60) {
      minutes = 0;
      hours += 1;
    }

    return formatDate(year, month, day, hours, minutes);
  }

  if (typeof val === 'number') {
    const d = xlsx.SSF.parse_date_code(val);

    let year = d.y;
    let month = d.m;
    let day = d.d;
    let hours = d.H || 0;
    let minutes = d.M || 0;
    let seconds = d.S || 0;

    if (seconds >= 30) {
      minutes += 1;
    }

    if (minutes === 60) {
      minutes = 0;
      hours += 1;
    }

    return formatDate(year, month, day, hours, minutes);
  }

  const parsed = new Date(String(val).trim());

  if (!isNaN(parsed.valueOf())) {
    let year = parsed.getFullYear();
    let month = parsed.getMonth() + 1;
    let day = parsed.getDate();
    let hours = parsed.getHours();
    let minutes = parsed.getMinutes();

    if (parsed.getSeconds() >= 30) {
      minutes += 1;
    }

    if (minutes === 60) {
      minutes = 0;
      hours += 1;
    }

    return formatDate(year, month, day, hours, minutes);
  }

  return null;
}



/**
 * Extract source notes from an Excel row.
 * Returns array of { text: string, isPublic: boolean }
 */
function getSourceNotes(row) {
  const notes = [];

  const cvn = String(
    row['Customer Visible Notes'] ||
    row['Customer visible notes'] ||
    row['customer visible notes'] || ''
  ).trim();
  if (cvn) notes.push({ text: cvn, isPublic: true });

  const intNotes = String(
    row['Internal Notes'] ||
    row['Internal notes'] ||
    row['internal notes'] || ''
  ).trim();
  if (intNotes) notes.push({ text: intNotes, isPublic: false });

  return notes;
}

// ── XML Parser ────────────────────────────────────────────────────

async function parseXmlFile(filePath) {
  log('info', `Parsing XML: ${path.basename(filePath)}`);
  const raw = fs.readFileSync(filePath, 'utf8');
  const result = await xml2js.parseStringPromise(raw, {
    explicitArray: false,  // simpler access; tags with 1 element unwrapped
    mergeAttrs: true,
  });
  const tickets = result['helpdesk-tickets']['helpdesk-ticket'];
  return Array.isArray(tickets) ? tickets : [tickets];
}

/**
 * Extract tag names from a parsed ticket.
 * xml2js with explicitArray:false gives tag names as either:
 *   ticket.tags.tag             → single tag object  { type, name }
 *   ticket.tags.tag             → array of tag objects
 *   ticket.tags (no .tag prop)  → empty
 */
function getTagNames(ticket) {
  try {
    const tagsNode = ticket['tags'];
    if (!tagsNode || !tagsNode['tag']) return [];
    const tagArr = Array.isArray(tagsNode['tag']) ? tagsNode['tag'] : [tagsNode['tag']];
    return tagArr
      .map(t => (t && t['name'] ? String(t['name']) : ''))
      .filter(Boolean);
  } catch (_) { return []; }
}

/** Returns true if the ticket has BOTH required migration tags (case-insensitive) */
function isMigratedTicket(ticket) {
  const names = getTagNames(ticket).map(n => n.toLowerCase().trim());
  return REQUIRED_TAGS.every(req => names.includes(req.toLowerCase()));
}

/** Get scalar value from xml2js parsed field (handles { _: value } and plain strings) */
function scalar(v) {
  if (v === undefined || v === null) return null;
  if (typeof v === 'object' && v !== null && '_' in v) return String(v._);
  return String(v);
}

/**
 * Extract all notes from a parsed ticket as flat array.
 * With explicitArray:false, note values are plain strings or {_,type} objects.
 */
function getTicketNotes(ticket) {
  const notes = [];
  try {
    const notesNode = ticket['notes'];
    if (!notesNode || !notesNode['helpdesk-note']) return notes;

    const raw = notesNode['helpdesk-note'];
    const arr = Array.isArray(raw) ? raw : [raw];

    for (const n of arr) {
      if (!n) continue;

      const deleted = scalar(n['deleted']) === 'true';
      if (deleted) continue;

      const id = scalar(n['id']);
      const body = scalar(n['body']) || '';
      const bodyHtml = scalar(n['body-html']) || '';
      const isPrivate = scalar(n['private']) === 'true';

      notes.push({ id, body, bodyHtml, isPrivate });
    }
  } catch (e) {
    log('warn', 'Error parsing notes', { error: e.message });
  }
  return notes;
}

// ── Note Matching ─────────────────────────────────────────────────

function noteMatches(xmlOrApiBodyText, sourceText) {
  // Try treating it as xmlNote object if it has bodyText or body properties
  const bodyText = typeof xmlOrApiBodyText === 'string' ? xmlOrApiBodyText :
    (xmlOrApiBodyText.body_text || xmlOrApiBodyText.bodyHtml || xmlOrApiBodyText.body || '');

  const normNote = bodyText.toLowerCase().replace(/[^a-z0-9]/g, '');
  const normSource = sourceText.toLowerCase().replace(/[^a-z0-9]/g, '');

  if (!normNote || !normSource || normNote.length < 15 || normSource.length < 15) return false;

  const checkLen = Math.min(50, normSource.length);
  const prefixSource = normSource.substring(0, checkLen);
  const midSource = normSource.substring(Math.floor(normSource.length / 2), Math.floor(normSource.length / 2) + checkLen);

  return normNote.includes(prefixSource) || normNote.includes(midSource) || normSource.includes(normNote);
}

/** Split a combined notes text into individual entries by timestamp prefix.
 *  Splits only at newline boundaries before a DD/MM/YYYY HH:MM:SS pattern.
 */
function splitIntoBlocks(text) {
  return text
    .split(/\n(?=\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2} - )/)
    .map(s => s.trim())
    .filter(Boolean);
}

// ── Main ──────────────────────────────────────────────────────────

async function main() {
  console.log('\n╔═══════════════════════════════════════════════════════╗');
  console.log('║      Freshdesk Migrated Ticket Update Script          ║');
  console.log(`║      DRY_RUN = ${String(DRY_RUN).padEnd(38)}║`);
  console.log('╚═══════════════════════════════════════════════════════╝\n');

  if (!DOMAIN || !API_KEY) {
    console.error('ERROR: Set FRESHDESK_DOMAIN and FRESHDESK_API_KEY in .env');
    process.exit(1);
  }
  if (!CF_DATE && !DRY_RUN) {
    log('warn', 'TICKET_CREATION_DATE_FIELD not set — date update will be skipped');
  }

  // ── 1. Load Excel ─────────────────────────────────────────────
  log('info', '── Step 1: Loading Excel source data ──');
  const allRows = [
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Service Request.xlsx'), 1),
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Customerservice_case.xlsx'), 1),
    ...readExcelFile(path.join(EXCEL_FOLDER, 'Master Ticket (All Status).xlsx'), 0),
  ];
  log('info', `Total Excel rows: ${allRows.length}`);
  const excelLookup = buildExcelLookup(allRows);

  // ── 2. Parse all XML files ────────────────────────────────────
  log('info', '\n── Step 2: Parsing XML ticket exports ──');
  const xmlFiles = fs.readdirSync(XML_FOLDER)
    .filter(f => /^Tickets\d+\.xml$/i.test(f))
    .sort();
  log('info', `XML files: ${xmlFiles.join(', ')}`);

  const migratedTickets = [];
  for (const xmlFile of xmlFiles) {
    const tickets = await parseXmlFile(path.join(XML_FOLDER, xmlFile));
    const filtered = tickets.filter(isMigratedTicket);
    log('info', `  ${xmlFile}: ${tickets.length} total → ${filtered.length} migrated`);
    migratedTickets.push(...filtered);
  }

  // --- OPTIONAL TEST FILTER ---
  let ticketsToProcess = migratedTickets;
  if (TEST_TICKET_IDS && TEST_TICKET_IDS.length > 0) {
    const stringIds = TEST_TICKET_IDS.map(String);
    ticketsToProcess = migratedTickets.filter(t => stringIds.includes(String(scalar(t['display-id']))));
    log('warn', `TEST_TICKET_IDS filter is ON. Found ${ticketsToProcess.length} matches out of ${TEST_TICKET_IDS.length} requested IDs.`);
  }

  log('info', `\nTotal migrated tickets to process: ${ticketsToProcess.length}`);

  if (ticketsToProcess.length === 0) {
    log('warn', 'No migrated tickets found. Check that tags match "Prod_migration" + "migrated".');
    log('info', 'Run: node check_data.js     for a detailed diagnosis.');
    saveLog();
    return;
  }

  // ── 3. Process each ticket ────────────────────────────────────
  log('info', '\n── Step 3: Processing tickets ──');
  const progress = loadProgress();
  const stats = { updated: 0, noteDeleted: 0, noteCreated: 0, skipped: 0, noMatch: 0 };

  for (const ticket of ticketsToProcess) {
    const displayId = scalar(ticket['display-id']);
    const subject = scalar(ticket['subject']) || '';

    let externalRef = '';
    if (ticket['custom_field']) {
      const extKey = Object.keys(ticket['custom_field']).find(key => key.startsWith('cf_external_reference'));
      if (extKey) {
        externalRef = scalar(ticket['custom_field'][extKey]) || '';
      }
    }

    if (!displayId) { log('warn', 'Ticket missing display-id, skipping'); continue; }
    if (progress.done.includes(displayId)) {
      log('info', `Ticket #${displayId} already processed — skipping`);
      stats.skipped++;
      continue;
    }

    console.log(`\n── Ticket #${displayId} | "${subject}"`);
    log('info', `Processing ticket #${displayId}`, { subject, externalRef });

    // ── Find matching Excel row ────────────────────────────────
    let excelRow = null;

    // 1. Exact match on External Reference
    if (externalRef) {
      excelRow = excelLookup.get(externalRef.toUpperCase().trim());
    }

    // 2. Fallback: Exact match on subject
    if (!excelRow) {
      excelRow = excelLookup.get(subject.toUpperCase());
    }

    // 3. Fallback: partial match (subject may contain SN number as part of longer string)
    if (!excelRow) {
      for (const [key, row] of excelLookup) {
        if (subject.toUpperCase().includes(key) || key.includes(subject.toUpperCase())) {
          excelRow = row;
          break;
        }
      }
    }

    if (!excelRow) {
      log('warn', `No Excel match for ticket #${displayId} (subject: "${subject}")`);
      stats.noMatch++;
      continue;
    }

    const matchedNumber = excelRow['Number'] || excelRow['Subject'] || '';
    log('info', `  Matched Excel row: ${matchedNumber}`);

    try {
      // ── 3a: Update Ticket Creation Date ───────────────────────
      const createdDate = getCreatedDate(excelRow);
      if (createdDate && CF_DATE) {
        await apiUpdateTicket(displayId, { [CF_DATE]: createdDate });
        stats.updated++;
        console.log(`  ✅ Date updated: ${createdDate}`);
      } else {
        log('warn', `  ⚠️  No date found for ticket #${displayId}`);
      }

      // ── 3b: Fix Notes ─────────────────────────────────────────
      const cvnText = String(excelRow['Customer Visible Notes'] || excelRow['Customer visible notes'] || excelRow['customer visible notes'] || '').trim();
      const intText = String(excelRow['Internal Notes'] || excelRow['Internal notes'] || excelRow['internal notes'] || '').trim();

      if (!cvnText && !intText) {
        console.log(`  Notes — no source notes in Excel.`);
      } else {
        // Pre-split source text into individual blocks by timestamp
        const cvnBlocks = cvnText ? splitIntoBlocks(cvnText) : [];
        const intBlocks = intText ? splitIntoBlocks(intText) : [];

        const liveNotes = await apiGetConversations(displayId);
        console.log(`  Notes — Live: ${liveNotes.length} | CVN blocks: ${cvnBlocks.length} | INT blocks: ${intBlocks.length}`);

        let needsCvn = false;
        let needsInt = false;

        for (const liveNote of liveNotes) {
          const isPrivateNote = liveNote.private === true;
          const bodyText = liveNote.body_text || stripHtml(liveNote.body) || '';

          const containsCvn = cvnText && noteMatches({ body: bodyText }, cvnText);
          const containsInt = intText && noteMatches({ body: bodyText }, intText);

          if (!containsCvn && !containsInt) continue; // Unrelated note

          // CVN: BAD if private, combined with INT, or source has >1 block (needs splitting)
          if (containsCvn) {
            const isBad = isPrivateNote || containsInt || cvnBlocks.length > 1;
            if (isBad) {
              const reason = isPrivateNote ? 'wrong visibility (private)' : containsInt ? 'combined with Internal Notes' : `needs splitting (${cvnBlocks.length} blocks)`;
              console.log(`  🔄 Note ${liveNote.id} is BAD — ${reason}. Deleting...`);
              await apiDeleteConversation(liveNote.id);
              stats.noteDeleted++;
              needsCvn = true;
              if (containsInt) needsInt = true;
            } else {
              console.log(`  ✅ Note ${liveNote.id} is already a PERFECT public Customer Visible Note.`);
            }
          }

          // INT: BAD if public, combined with CVN, or source has >1 block (needs splitting)
          if (containsInt && !(containsCvn && isPrivateNote)) {
            const isBad = !isPrivateNote || containsCvn || intBlocks.length > 1;
            if (isBad) {
              const reason = !isPrivateNote ? 'wrong visibility (public)' : containsCvn ? 'combined with Customer Visible Notes' : `needs splitting (${intBlocks.length} blocks)`;
              console.log(`  🔄 Note ${liveNote.id} is BAD — ${reason}. Deleting...`);
              await apiDeleteConversation(liveNote.id);
              stats.noteDeleted++;
              needsInt = true;
            } else {
              console.log(`  ✅ Note ${liveNote.id} is already a PERFECT private Internal Note.`);
            }
          }
        }

        // Recreate: one separate note per block, oldest first
        if (needsCvn && cvnBlocks.length > 0) {
          const ordered = [...cvnBlocks].reverse();
          for (const block of ordered) {
            const html = `<p>${block.replace(/\n/g, '<br>')}</p>`;
            await apiCreatePublicNote(displayId, html);
            stats.noteCreated++;
            console.log(`  ✨ Created PUBLIC note: "${block.substring(0, 80).replace(/\n/g, ' ')}"`);
          }
        }

        if (needsInt && intBlocks.length > 0) {
          const ordered = [...intBlocks].reverse();
          for (const block of ordered) {
            const html = `<p>${block.replace(/\n/g, '<br>')}</p>`;
            await apiCreatePrivateNote(displayId, html);
            stats.noteCreated++;
            console.log(`  ✨ Created PRIVATE note: "${block.substring(0, 80).replace(/\n/g, ' ')}"`);
          }
        }
      }

      progress.done.push(displayId);
      saveProgress(progress);

    } catch (err) {
      const errMsg = err.response?.data
        ? JSON.stringify(err.response.data)
        : err.message;
      log('error', `Failed on ticket #${displayId}`, { error: errMsg });
      progress.failed.push({ id: displayId, error: errMsg });
      saveProgress(progress);
    }
  }

  // ── Summary ───────────────────────────────────────────────────
  console.log('\n╔═══════════════════════════════════════════════════════╗');
  console.log('║                   MIGRATION COMPLETE                  ║');
  console.log('╠═══════════════════════════════════════════════════════╣');
  console.log(`║  Tickets date-updated  : ${String(stats.updated).padEnd(29)}║`);
  console.log(`║  Notes deleted         : ${String(stats.noteDeleted).padEnd(29)}║`);
  console.log(`║  Notes re-created      : ${String(stats.noteCreated).padEnd(29)}║`);
  console.log(`║  No Excel match        : ${String(stats.noMatch).padEnd(29)}║`);
  console.log(`║  Skipped (done before) : ${String(stats.skipped).padEnd(29)}║`);
  console.log(`║  Failed                : ${String(progress.failed.length).padEnd(29)}║`);
  console.log('╚═══════════════════════════════════════════════════════╝\n');

  saveLog();
  log('info', `Full log saved → ${LOG_FILE}`);
  if (progress.failed.length > 0) {
    log('warn', `Failed tickets saved in progress.json → check and re-run`);
  }
}

main().catch(err => {
  console.error('\nFATAL ERROR:', err.message);
  saveLog();
  process.exit(1);
});
