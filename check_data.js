/**
 * check_data.js  (v2 - deeper inspection)
 * ─────────────────────────────────────────────────────────
 * Run: node check_data.js
 */

require('dotenv').config();
const fs     = require('fs');
const path   = require('path');
const xlsx   = require('xlsx');
const xml2js = require('xml2js');

const XML_FOLDER   = process.env.XML_FOLDER;
const EXCEL_FOLDER = process.env.EXCEL_FOLDER;

// ── Excel: show ALL sheets and columns ─────────────────────

function inspectExcel(filePath) {
  console.log(`\n${'─'.repeat(70)}`);
  console.log(`FILE: ${path.basename(filePath)}`);
  if (!fs.existsSync(filePath)) { console.log('  ❌ File not found!'); return; }

  const wb = xlsx.readFile(filePath, { cellDates: true });
  console.log(`  Total sheets: ${wb.SheetNames.length}`);

  wb.SheetNames.forEach((name, i) => {
    const ws   = wb.Sheets[name];
    const rows = xlsx.utils.sheet_to_json(ws, { defval: '' });
    console.log(`\n  [Sheet ${i + 1}] "${name}"  →  ${rows.length} rows`);
    if (rows.length > 0) {
      const cols = Object.keys(rows[0]);
      console.log(`    Columns (${cols.length}):`);
      cols.forEach(c => console.log(`      - "${c}"`));

      // Sample first row
      console.log(`    Sample row 1:`);
      cols.slice(0, 10).forEach(c => {
        const val = rows[0][c];
        console.log(`      "${c}" = ${val}`);
      });
    }
  });
}

// ── XML: show raw tag structure & sample notes ──────────────

async function inspectXml() {
  console.log(`\n${'═'.repeat(70)}`);
  console.log('XML TAG & NOTE STRUCTURE INSPECTION');
  console.log(`${'═'.repeat(70)}`);

  const xmlFiles = fs.readdirSync(XML_FOLDER)
    .filter(f => /^Tickets\d+\.xml$/i.test(f))
    .sort();

  // Look only at first XML file for structure
  for (const xmlFile of xmlFiles.slice(0, 1)) {
    const raw    = fs.readFileSync(path.join(XML_FOLDER, xmlFile), 'utf8');
    const result = await xml2js.parseStringPromise(raw, {
      explicitArray: true,   // keep explicit arrays to see real structure
    });

    const all     = result['helpdesk-tickets']['helpdesk-ticket'];
    const tickets = Array.isArray(all) ? all : [all];

    console.log(`\n${xmlFile}: ${tickets.length} tickets`);

    // Show raw tags structure for first 5 tickets
    console.log('\n── RAW TAGS structure (first 5 tickets):');
    tickets.slice(0, 5).forEach((t, i) => {
      const displayId = t['display-id']?.[0]?._ || t['display-id']?.[0] || '?';
      const subject   = t['subject']?.[0] || '';
      const rawTags   = JSON.stringify(t['tags']);
      console.log(`\n  Ticket #${displayId} | "${subject}"`);
      console.log(`    Tags (raw): ${rawTags}`);
    });

    // Find tickets that DO have tags
    console.log('\n── Tickets WITH any tags:');
    let found = 0;
    for (const t of tickets) {
      const rawTags = JSON.stringify(t['tags'] || '');
      if (rawTags && rawTags !== '"undefined"' && rawTags !== '""' &&
          rawTags !== '[{"$":{"type":"array"}}]' &&
          rawTags !== '[{"$":{"type":"array"},"tag":[]}]') {
        const displayId = t['display-id']?.[0]?._ || t['display-id']?.[0] || '?';
        const subject   = t['subject']?.[0] || '';
        console.log(`\n  Ticket #${displayId} | "${subject}"`);
        console.log(`    Tags: ${rawTags}`);
        found++;
        if (found >= 5) break;
      }
    }

    if (found === 0) {
      console.log('  ⚠️  No tickets with tags found in ' + xmlFile);
      console.log('  Checking all XML files...');

      for (const xf of xmlFiles.slice(1)) {
        const r2  = fs.readFileSync(path.join(XML_FOLDER, xf), 'utf8');
        const pr2 = await xml2js.parseStringPromise(r2, { explicitArray: true });
        const ts  = pr2['helpdesk-tickets']['helpdesk-ticket'];
        const arr = Array.isArray(ts) ? ts : [ts];

        for (const t of arr) {
          const rawTags = JSON.stringify(t['tags'] || '');
          if (rawTags && rawTags.includes('tag') && rawTags.length > 50) {
            const displayId = t['display-id']?.[0]?._ || t['display-id']?.[0] || '?';
            const subject   = t['subject']?.[0] || '';
            console.log(`\n  [${xf}] Ticket #${displayId} | "${subject}"`);
            console.log(`    Tags: ${rawTags.substring(0, 400)}`);
            found++;
            if (found >= 5) break;
          }
        }
        if (found >= 5) break;
      }
    }

    // Show notes structure for first ticket WITH notes
    console.log('\n── RAW NOTES structure (first ticket with notes):');
    for (const t of tickets) {
      const notesNode = t['notes']?.[0];
      if (notesNode && notesNode['helpdesk-note']) {
        const displayId = t['display-id']?.[0]?._ || t['display-id']?.[0] || '?';
        console.log(`\n  Ticket #${displayId}`);
        const notes = notesNode['helpdesk-note'];
        const note  = notes[0];
        console.log('  Note keys:', Object.keys(note));
        console.log('  private:', note['private']);
        console.log('  body (first 100 chars):', String(note['body']?.[0] || '').substring(0, 100));
        break;
      }
    }
  }
}

async function main() {
  console.log('\n╔══════════════════════════════════════════════════════╗');
  console.log('║           DEEP DATA DIAGNOSTIC CHECK v2              ║');
  console.log('╚══════════════════════════════════════════════════════╝');

  inspectExcel(path.join(EXCEL_FOLDER, 'Service Request.xlsx'));
  inspectExcel(path.join(EXCEL_FOLDER, 'sn_Closed-Resolved_Customerservice_case.xlsx'));

  await inspectXml();
  console.log('\n✅ Diagnostic done.\n');
}

main().catch(console.error);
