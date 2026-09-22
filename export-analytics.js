/**
 * Standalone Analytics Exporter for BD Meetings
 * Usage: node export-analytics.js
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('c:/Users/ben.arthur/Desktop/bd tracker/node_modules/xlsx');

const SOURCE_FILE = path.join(__dirname, 'BD MEETINGS 2026 .xlsx');
const OUTPUT_FILE = path.join(__dirname, `BD_Analytics_${new Date().toISOString().slice(0, 10)}.xlsx`);

const STAGES = [
  'New Meetings',
  'Follow Ups',
  'Contract Sent',
  'Invoice Sent',
  'Onboarded',
  'No-Show',
  'Dead Leads',
  'Temporary Inactive'
];

const EXCLUDED_AGENTS = ['russ', 'george', 'caroline', 'caroline richards'];

function isExcluded(name) {
  if (!name) return true;
  const clean = String(name).trim().toLowerCase();
  return EXCLUDED_AGENTS.some(e => clean === e || clean.startsWith(e) || clean.endsWith(e));
}

function parseBoolean(val) {
  if (typeof val === 'boolean') return val;
  const s = String(val || '').trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y';
}

function main() {
  if (!fs.existsSync(SOURCE_FILE)) {
    console.error(`Error: Source file not found: ${SOURCE_FILE}`);
    process.exit(1);
  }

  console.log(`Loading workbook: ${SOURCE_FILE}...`);
  const wb = xlsx.readFile(SOURCE_FILE);

  const openers = {};
  const allMeetings = [];

  STAGES.forEach(stageName => {
    const sheet = wb.Sheets[stageName];
    if (!sheet) return;

    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    if (!rows || rows.length < 2) return;

    const headers = (rows[0] || []).map(h => String(h || '').trim().toLowerCase());

    let openerIdx = headers.findIndex(h => h === 'opener' || h.includes('opener') || h === 'agent' || h === 'rep');
    if (openerIdx === -1) openerIdx = 3; // Default Col D

    const medbIdx = headers.findIndex(h => h.includes('medb') || h.includes('med b'));
    const ppoIdx = headers.findIndex(h => h === 'ppo' || h.includes('ppo'));
    const dateIdx = headers.findIndex(h => h.includes('date') || h.includes('created'));
    const compIdx = headers.findIndex(h => h.includes('company') || h.includes('business'));
    const personIdx = headers.findIndex(h => h.includes('authorized') || h.includes('contact') || h.includes('person') || h.includes('name'));

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const rawOpener = String(row[openerIdx] || '').trim();
      if (!rawOpener || isExcluded(rawOpener)) continue;

      const isMedB = medbIdx !== -1 ? parseBoolean(row[medbIdx]) : false;
      const isPPO = ppoIdx !== -1 ? parseBoolean(row[ppoIdx]) : false;
      const compName = compIdx !== -1 ? String(row[compIdx] || '').trim() : '';
      const personName = personIdx !== -1 ? String(row[personIdx] || '').trim() : '';
      const dateVal = dateIdx !== -1 ? String(row[dateIdx] || '').trim() : '';

      if (!openers[rawOpener]) {
        openers[rawOpener] = {
          opener: rawOpener,
          booked: 0,
          medBCount: 0,
          ppoCount: 0,
          stages: {}
        };
        STAGES.forEach(s => { openers[rawOpener].stages[s] = 0; });
      }

      openers[rawOpener].booked++;
      openers[rawOpener].stages[stageName] = (openers[rawOpener].stages[stageName] || 0) + 1;
      if (isMedB) openers[rawOpener].medBCount++;
      if (isPPO) openers[rawOpener].ppoCount++;

      allMeetings.push({
        date: dateVal,
        opener: rawOpener,
        stage: stageName,
        company: compName,
        contact: personName,
        medB: isMedB ? 'Yes' : 'No',
        ppo: isPPO ? 'Yes' : 'No'
      });
    }
  });

  // Build Summary Sheet
  const summaryHeaders = [
    'Agent / Opener',
    'Total Meetings',
    'Med B Count',
    'Med B %',
    'PPO Count',
    'PPO %',
    'Attended (Non-NoShow)',
    'Show Rate %',
    'Onboarded',
    'Close Rate %',
    ...STAGES
  ];

  const summaryRows = [];
  const totalStats = {
    booked: 0,
    medBCount: 0,
    ppoCount: 0,
    attended: 0,
    onboarded: 0,
    stages: {}
  };
  STAGES.forEach(s => { totalStats.stages[s] = 0; });

  const sortedOpeners = Object.values(openers).sort((a, b) => a.opener.localeCompare(b.opener));

  sortedOpeners.forEach(op => {
    const noShow = op.stages['No-Show'] || 0;
    const attended = Math.max(0, op.booked - noShow);
    const onboarded = op.stages['Onboarded'] || 0;
    const medBRate = op.booked > 0 ? (op.medBCount / op.booked) * 100 : 0;
    const ppoRate = op.booked > 0 ? (op.ppoCount / op.booked) * 100 : 0;
    const showRate = op.booked > 0 ? (attended / op.booked) * 100 : 0;
    const closeRate = op.booked > 0 ? (onboarded / op.booked) * 100 : 0;

    totalStats.booked += op.booked;
    totalStats.medBCount += op.medBCount;
    totalStats.ppoCount += op.ppoCount;
    totalStats.attended += attended;
    totalStats.onboarded += onboarded;
    STAGES.forEach(s => { totalStats.stages[s] += op.stages[s] || 0; });

    summaryRows.push([
      op.opener,
      op.booked,
      op.medBCount,
      `${medBRate.toFixed(1)}%`,
      op.ppoCount,
      `${ppoRate.toFixed(1)}%`,
      attended,
      `${showRate.toFixed(1)}%`,
      onboarded,
      `${closeRate.toFixed(1)}%`,
      ...STAGES.map(s => op.stages[s] || 0)
    ]);
  });

  // Totals Row
  const totMedBRate = totalStats.booked > 0 ? (totalStats.medBCount / totalStats.booked) * 100 : 0;
  const totPPORate = totalStats.booked > 0 ? (totalStats.ppoCount / totalStats.booked) * 100 : 0;
  const totShowRate = totalStats.booked > 0 ? (totalStats.attended / totalStats.booked) * 100 : 0;
  const totCloseRate = totalStats.booked > 0 ? (totalStats.onboarded / totalStats.booked) * 100 : 0;

  summaryRows.push([
    'TOTAL / TEAM',
    totalStats.booked,
    totalStats.medBCount,
    `${totMedBRate.toFixed(1)}%`,
    totalStats.ppoCount,
    `${totPPORate.toFixed(1)}%`,
    totalStats.attended,
    `${totShowRate.toFixed(1)}%`,
    totalStats.onboarded,
    `${totCloseRate.toFixed(1)}%`,
    ...STAGES.map(s => totalStats.stages[s] || 0)
  ]);

  const outWb = xlsx.utils.book_new();

  const wsSummary = xlsx.utils.aoa_to_sheet([summaryHeaders, ...summaryRows]);
  wsSummary['!cols'] = summaryHeaders.map((h, i) => {
    let max = h.length;
    summaryRows.forEach(r => {
      const len = r[i] !== undefined && r[i] !== null ? String(r[i]).length : 0;
      if (len > max) max = len;
    });
    return { wch: Math.min(Math.max(max + 3, 10), 30) };
  });
  xlsx.utils.book_append_sheet(outWb, wsSummary, 'Opener Analytics');

  // Meeting Details Sheet
  const meetingHeaders = ['Date', 'Opener', 'Stage', 'Company Name', 'Contact Person', 'Med B', 'PPO'];
  const meetingRows = allMeetings.map(m => [
    m.date, m.opener, m.stage, m.company, m.contact, m.medB, m.ppo
  ]);
  const wsMeetings = xlsx.utils.aoa_to_sheet([meetingHeaders, ...meetingRows]);
  wsMeetings['!cols'] = meetingHeaders.map((h, i) => ({ wch: 18 }));
  xlsx.utils.book_append_sheet(outWb, wsMeetings, 'Meeting Pipeline');

  xlsx.writeFile(outWb, OUTPUT_FILE);
  console.log(`\nSuccessfully created analytics export: ${OUTPUT_FILE}`);
  console.log(`Total Openers: ${sortedOpeners.length}`);
  console.log(`Total Meetings: ${totalStats.booked}`);
  console.log(`Med B Total: ${totalStats.medBCount} (${totMedBRate.toFixed(1)}%)`);
  console.log(`PPO Total: ${totalStats.ppoCount} (${totPPORate.toFixed(1)}%)`);
}

main();
