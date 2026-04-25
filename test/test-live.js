'use strict';
// Live tests against a running Outlook instance.
// Driven by directly invoking handleTool via the same code path as the MCP server.
const path = require('path');
const fs = require('fs');

// Inline-load index.js as a module (it exports compileFilter + buildQueryScript).
// To test the actual handler we need to expose it; for now, spawn PS via the build script.
const { compileFilter, buildQueryScript } = require('../index.js');
const { exec } = require('child_process');
const { promisify } = require('util');
const os = require('os');
const execAsync = promisify(exec);

const INIT = `
$ErrorActionPreference = 'Stop'
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  $ol = New-Object -ComObject Outlook.Application
}
$ns = $ol.GetNamespace('MAPI')
`;

async function ps(script) {
  const tmp = path.join(os.tmpdir(), `email-mcp-test-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
  fs.writeFileSync(tmp, script, 'utf8');
  try {
    const { stdout, stderr } = await execAsync(
      `powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "${tmp}"`,
      { timeout: 60000, maxBuffer: 32 * 1024 * 1024 }
    );
    if (stderr && stderr.trim()) process.stderr.write('[ps stderr] ' + stderr + '\n');
    return stdout.trim();
  } catch (err) {
    const msg = (err.stderr || err.stdout || err.message || '').toString();
    // keep tmp for debug
    throw new Error('PS failed (' + tmp + '): ' + msg.slice(0, 1000));
  } finally {
    // leave tmp on success too for now
  }
}

async function runQuery(opts) {
  const dasl = compileFilter(opts.filter || {}, { allow_slow: !!opts.allow_slow });
  const limit = opts.limit ?? 5;
  // Mirror handleTool's default-field rule
  let fields = opts.fields;
  if (!fields) {
    fields = limit <= 20
      ? ['entry_id','subject','from','received','unread','has_attachments']
      : ['subject','from','received'];
  }
  const script = buildQueryScript({
    daslFilter: dasl,
    account: opts.account || '',
    limit,
    offset: opts.offset ?? 0,
    order: opts.order_by || 'received_desc',
    fields,
  });
  const out = await ps(script);
  return JSON.parse(out);
}

let pass = 0, fail = 0;
async function t(name, fn) {
  try { await fn(); console.log('  ok  ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + (e.stack || e.message)); fail++; }
  // Outlook COM needs a beat between back-to-back PS spawns
  await new Promise(r => setTimeout(r, 500));
}

(async () => {
  console.log('live tests');

  await t('basic inbox query, limit 3', async () => {
    const r = await runQuery({ filter: {}, limit: 3 });
    if (!Array.isArray(r.results)) throw new Error('no results array');
    if (r.results.length === 0) throw new Error('expected at least one email in inbox');
    if (r.results.length > 3) throw new Error('limit not respected');
    const sample = r.results[0];
    if (!sample.entry_id || !sample.subject) throw new Error('missing default fields: ' + JSON.stringify(sample));
  });

  await t('large scan omits entry_id by default', async () => {
    const r = await runQuery({ filter: {}, limit: 50 });
    const sample = r.results[0];
    if (sample && sample.entry_id) throw new Error('entry_id should be omitted for limit > 20');
    if (sample && !sample.subject) throw new Error('subject should be present');
  });

  await t('explicit fields override default', async () => {
    const r = await runQuery({ filter: {}, limit: 50, fields: ['entry_id','subject'] });
    const sample = r.results[0];
    if (!sample.entry_id) throw new Error('entry_id should be present when explicitly requested');
  });

  await t('contains filter on subject', async () => {
    const r = await runQuery({
      filter: { subject: { $contains: 'meeting' } },
      limit: 10,
    });
    for (const e of r.results) {
      if (!/meeting/i.test(e.subject)) throw new Error('non-matching subject: ' + e.subject);
    }
  });

  await t('date range $gte', async () => {
    const r = await runQuery({
      filter: { received: { $gte: '2026-01-01' } },
      limit: 5,
    });
    for (const e of r.results) {
      if (e.received < '2026-01-01') throw new Error('older than filter: ' + e.received);
    }
  });

  await t('boolean tree: $or of two contains', async () => {
    const r = await runQuery({
      filter: { $or: [
        { subject: { $contains: 'meeting' } },
        { subject: { $contains: 'recap' } },
      ]},
      limit: 10,
    });
    for (const e of r.results) {
      if (!/meeting|recap/i.test(e.subject)) throw new Error('bad match: ' + e.subject);
    }
  });

  await t('unread filter (no error even if 0 results)', async () => {
    const r = await runQuery({
      filter: { unread: true },
      limit: 5,
    });
    if (!Array.isArray(r.results)) throw new Error('no results array');
    for (const e of r.results) {
      if (e.unread !== true) throw new Error('expected unread, got: ' + JSON.stringify(e));
    }
  });

  await t('pagination + has_more', async () => {
    const r1 = await runQuery({ filter: {}, limit: 5, offset: 0 });
    const r2 = await runQuery({ filter: {}, limit: 5, offset: 5 });
    if (r1.results[0].entry_id === r2.results[0].entry_id) {
      throw new Error('pagination did not advance');
    }
    if (r1.total_matched < 10) {
      // can't reliably test has_more without enough mail
      console.log('       (skipped has_more check, total_matched=' + r1.total_matched + ')');
      return;
    }
    if (!r1.has_more) throw new Error('expected has_more=true');
  });

  await t('asc order', async () => {
    const r = await runQuery({ filter: {}, limit: 3, order_by: 'received_asc' });
    if (r.results.length >= 2) {
      if (r.results[0].received > r.results[1].received) throw new Error('not ascending');
    }
  });

  await t('body field requires allow_slow', async () => {
    let threw = false;
    try { await runQuery({ filter: { body: { $contains: 'x' } }, limit: 1 }); }
    catch (e) { threw = true; }
    if (!threw) throw new Error('expected error without allow_slow');
  });

  await t('all-folders default scan returns results', async () => {
    const r = await runQuery({ filter: {}, limit: 5 });
    if (!Array.isArray(r.results) || r.results.length === 0) throw new Error('no results');
    // total_matched should reflect more than a single folder when scanning all mail folders
    if (typeof r.total_matched !== 'number' || r.total_matched < r.results.length) {
      throw new Error('bad total_matched: ' + r.total_matched);
    }
  });

  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail > 0 ? 1 : 0);
})();
