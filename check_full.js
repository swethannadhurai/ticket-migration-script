/**
 * check_full.js — Full pre-flight dry-run preview
 * ─────────────────────────────────────────────────────────
 * Shows exactly what migrate_tickets.js WOULD do:
 *   - How many migrated XML tickets found
 *   - How many match an Excel row
 *   - What creation date would be set
 *   - Which notes would be deleted / recreated as public or private
 *
 * Run: node check_full.js
 * ─────────────────────────────────────────────────────────
 */
'use strict';
require('dotenv').config();

const fs     = require('fs');
const path   = require('path');
const xlsx   = require('xlsx');
const xml2js = require('xml2js');

const XML_FOLDER   = process.env.XML_FOLDER;
const EXCEL_FOLDER = process.env.EXCEL_FOLDER;
const REQUIRED_TAGS = ['prod_migration', 'migrated'];

// ── Helpers (same logic as migrate_tickets.js) ────────────────────

function stripHtml(html = '') {
  return String(html || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ').replace(/&#?\w+;/g, ' ')
    .replace(/\s+/g, ' ').trim().toLowerCase();
}

function scalar(v) {
  if (v === undefined || v === null) return null;
  if (typeof v === 'object' && '_' in v) return String(v._);
  return String(v);
}

function getTagNames(ticket) {
  try {
    const tagsNode = ticket['tags'];
    if (!tagsNode || !tagsNode['tag']) return [];
    const arr = Array.isArray(tagsNode['tag']) ? tagsNode['tag'] : [tagsNode['tag']];
    return arr.map(t => (t && t['name'] ? String(t['name']) : '')).filter(Boolean);
  } catch (_) { return []; }
}

function isMigrated(ticket) {
  const names = getTagNames(ticket).map(n => n.toLowerCase().trim());
  return REQUIRED_TAGS.every(req => names.includes(req));
}

function getTicketNotes(ticket) {
  const notes = [];
  try {
    const notesNode = ticket['notes'];
    if (!notesNode || !notesNode['helpdesk-note']) return notes;
    const arr = Array.isArray(notesNode['helpdesk-note'])
      ? notesNode['helpdesk-note'] : [notesNode['helpdesk-note']];
    for (const n of arr) {
      if (!n || scalar(n['deleted']) === 'true') continue;
      notes.push({
        id:        scalar(n['id']),
        body:      scalar(n['body'])      || '',
        bodyHtml:  scalar(n['body-html']) || '',
        isPrivate: scalar(n['private'])   === 'true',
      });
    }
  } catch (_) {}
  return notes;
}

function noteMatches(xmlNote, sourceText) {
  const normXml    = stripHtml(xmlNote.bodyHtml) || xmlNote.body.toLowerCase().replace(/\s+/g, ' ').trim();
  const normSource = stripHtml(sourceText);
  if (!normXml || !normSource || normXml.length < 10 || normSource.length < 10) return false;
  const shorter = normXml.length <= normSource.length ? normXml : normSource;
  const longer  = normXml.length <= normSource.length ? normSource : normXml;
  return longer.includes(shorter);
}

function getCreatedDate(row) {
  const val = row['Created'] || row['created'] || row['Opened'] || '';
  if (!val) return null;
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'number') {
    const d = xlsx.SSF.parse_date_code(val);
    return new Date(Date.UTC(d.y, d.m - 1, d.d)).toISOString();
  }
  return String(val).trim() || null;
}

function getSourceNotes(row) {
  const notes = [];
  const cvn = String(row['Customer Visible Notes'] || row['Customer visible notes'] || '').trim();
  if (cvn) notes.push({ text: cvn, isPublic: true });
  const intN = String(row['Internal Notes'] || row['Internal notes'] || '').trim();
  if (intN) notes.push({ text: intN, isPublic: false });
  return notes;
}

// ── Main ──────────────────────────────────────────────────────────

async function main() {
  console.log('\n╔══════════════════════════════════════════════════════╗');
  console.log('║          FULL PRE-FLIGHT CHECK (no API calls)        ║');
  console.log('╚══════════════════════════════════════════════════════╝\n');

  // ── Load Excel ──────────────────────────────────────────────────
  console.log('── Loading Excel files...');
  const allRows = [];

  for (const file of ['Service Request.xlsx', 'sn_Closed-Resolved_Customerservice_case.xlsx']) {
    const fp = path.join(EXCEL_FOLDER, file);
    if (!fs.existsSync(fp)) { console.log(`  ⚠️  Not found: ${file}`); continue; }
    const wb   = xlsx.readFile(fp, { cellDates: true });
    const name = wb.SheetNames[0];
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
    console.log(`  ✅ ${file} → sheet "${name}" → ${rows.length} rows`);
    allRows.push(...rows);
  }

  // Build lookup
  const lookup = new Map();
  for (const row of allRows) {
    const num = String(row['Number'] || '').trim();
    const sub = String(row['Subject'] || '').trim();
    if (num) lookup.set(num.toUpperCase(), row);
    if (sub && sub !== num) lookup.set(sub.toUpperCase(), row);
  }
  console.log(`  Total lookup keys: ${lookup.size}\n`);

  // ── Parse XML ───────────────────────────────────────────────────
  console.log('── Parsing XML files...');
  const xmlFiles = fs.readdirSync(XML_FOLDER)
    .filter(f => /^Tickets\d+\.xml$/i.test(f)).sort();

  const migratedTickets = [];
  for (const f of xmlFiles) {
    const raw    = fs.readFileSync(path.join(XML_FOLDER, f), 'utf8');
    const result = await xml2js.parseStringPromise(raw, { explicitArray: false, mergeAttrs: true });
    const all    = result['helpdesk-tickets']['helpdesk-ticket'];
    const tickets = Array.isArray(all) ? all : [all];
    const mig    = tickets.filter(isMigrated);
    console.log(`  ${f}: ${tickets.length} total | ${mig.length} migrated`);
    migratedTickets.push(...mig);
  }
  console.log(`\n  ✅ TOTAL migrated tickets: ${migratedTickets.length}\n`);

  if (migratedTickets.length === 0) {
    console.log('❌ No migrated tickets found!');
    console.log('   Check that tags are "Prod_migration" + "migrated" (case-insensitive).');
    return;
  }

  // ── Match & Preview ─────────────────────────────────────────────
  console.log('── Matching tickets with Excel rows...\n');

  let matched = 0, noMatch = 0;
  const PREVIEW_LIMIT = 10; // show details for first N tickets

  for (let i = 0; i < migratedTickets.length; i++) {
    const t         = migratedTickets[i];
    const displayId = scalar(t['display-id']);
    const subject   = scalar(t['subject']) || '';
    const tags      = getTagNames(t).join(', ');

    // Excel match
    let excelRow = lookup.get(subject.toUpperCase());
    if (!excelRow) {
      for (const [key, row] of lookup) {
        if (subject.toUpperCase().includes(key) || key.includes(subject.toUpperCase())) {
          excelRow = row; break;
        }
      }
    }

    if (!excelRow) {
      if (i < PREVIEW_LIMIT) {
        console.log(`  ❌ #${displayId} | "${subject}" | tags:[${tags}] — NO EXCEL MATCH`);
      }
      noMatch++;
      continue;
    }

    matched++;
    const createdDate  = getCreatedDate(excelRow);
    const sourceNotes  = getSourceNotes(excelRow);
    const xmlNotes     = getTicketNotes(t);

    if (i < PREVIEW_LIMIT) {
      console.log(`  ✅ #${displayId} | "${subject}"`);
      console.log(`     Tags          : [${tags}]`);
      console.log(`     Created date  : ${createdDate || '⚠️  NOT FOUND'}`);
      console.log(`     XML notes     : ${xmlNotes.length}`);
      console.log(`     Source notes  : ${sourceNotes.length} (${sourceNotes.map(n => n.isPublic ? 'PUBLIC' : 'PRIVATE').join(', ')})`);

      // Preview which notes would be fixed
      for (const xmlNote of xmlNotes) {
        const src = sourceNotes.find(s => noteMatches(xmlNote, s.text));
        if (src) {
          const currentVis = xmlNote.isPrivate ? 'private' : 'public';
          const targetVis  = src.isPublic ? 'PUBLIC' : 'PRIVATE';
          const action     = currentVis === targetVis.toLowerCase() ? 'SAME (no vis change needed)' : `FIX ${currentVis} → ${targetVis}`;
          console.log(`     Note ${xmlNote.id}: duplicate → DELETE + RECREATE as ${targetVis} [${action}]`);
        } else {
          console.log(`     Note ${xmlNote.id}: NEW (keep untouched)`);
        }
      }
      console.log('');
    }
  }

  // ── Summary ─────────────────────────────────────────────────────
  console.log('╔══════════════════════════════════════════════════════╗');
  console.log('║                   CHECK SUMMARY                     ║');
  console.log('╠══════════════════════════════════════════════════════╣');
  console.log(`║  Migrated tickets found   : ${String(migratedTickets.length).padEnd(24)}║`);
  console.log(`║  Matched with Excel       : ${String(matched).padEnd(24)}║`);
  console.log(`║  No Excel match           : ${String(noMatch).padEnd(24)}║`);
  console.log('╚══════════════════════════════════════════════════════╝');

  if (noMatch > 0) {
    console.log('\n⚠️  Some tickets could not be matched to Excel rows.');
    console.log('   This may be normal if those tickets are from the CS file where');
    console.log('   the subject/number columns differ. Check the first few unmatched tickets above.');
  }

  console.log('\n✅ If the above looks correct, run the actual migration:');
  console.log('   node migrate_tickets.js\n');
}

main().catch(console.error);
