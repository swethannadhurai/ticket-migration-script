/**
 * check_tags.js — quick check to count migrated tickets with fixed tag parsing
 * Run: node check_tags.js
 */
require('dotenv').config();
const fs     = require('fs');
const path   = require('path');
const xml2js = require('xml2js');

const XML_FOLDER    = process.env.XML_FOLDER;
const REQUIRED_TAGS = ['prod_migration', 'migrated'];

function getTagNames(ticket) {
  try {
    const tagsNode = ticket['tags'];
    if (!tagsNode || !tagsNode['tag']) return [];
    const tagArr = Array.isArray(tagsNode['tag']) ? tagsNode['tag'] : [tagsNode['tag']];
    return tagArr.map(t => (t && t['name'] ? String(t['name']) : '')).filter(Boolean);
  } catch (_) { return []; }
}

function isMigrated(ticket) {
  const names = getTagNames(ticket).map(n => n.toLowerCase().trim());
  return REQUIRED_TAGS.every(req => names.includes(req));
}

async function main() {
  const xmlFiles = fs.readdirSync(XML_FOLDER)
    .filter(f => /^Tickets\d+\.xml$/i.test(f)).sort();

  let total = 0;
  for (const f of xmlFiles) {
    const raw    = fs.readFileSync(path.join(XML_FOLDER, f), 'utf8');
    const result = await xml2js.parseStringPromise(raw, { explicitArray: false, mergeAttrs: true });
    const all    = result['helpdesk-tickets']['helpdesk-ticket'];
    const tickets = Array.isArray(all) ? all : [all];
    const migrated = tickets.filter(isMigrated);
    total += migrated.length;
    console.log(`${f}: ${tickets.length} total | ${migrated.length} migrated`);

    // Print first 3 samples
    migrated.slice(0, 3).forEach(t => {
      const id  = t['display-id']?._ || t['display-id'] || '?';
      const sub = t['subject'] || '';
      const tags = getTagNames(t).join(', ');
      console.log(`   → #${id} | "${sub}" | tags: [${tags}]`);
    });
  }
  console.log(`\nTOTAL migrated tickets: ${total}`);
}

main().catch(console.error);
