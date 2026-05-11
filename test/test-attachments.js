'use strict';
// Live tests for download_attachments. Exercises the same handleTool
// code path as the MCP server by invoking PowerShell directly.
//
// Seed: set ATTACH_TEST_ENTRY_ID env var to an email's EntryID that has
//   - at least one real (non-inline) attachment
//   - at least one inline image (e.g. an HTML signature with a logo)
// Optional: ATTACH_TEST_INLINE_ONLY_ENTRY_ID for an email with only inline images.
// Without seeds, individual tests skip (not fail).

const path = require('path');
const fs = require('fs');
const os = require('os');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);

const REPO_ROOT = path.resolve(__dirname, '..');
const DOWNLOADS_ROOT = path.join(REPO_ROOT, 'downloads');

async function ps(script) {
  const tmp = path.join(os.tmpdir(), `email-mcp-test-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
  fs.writeFileSync(tmp, '\uFEFF' + script, 'utf8');
  try {
    const { stdout, stderr } = await execAsync(
      `powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "${tmp}"`,
      { timeout: 60000, maxBuffer: 32 * 1024 * 1024, encoding: 'buffer' }
    );
    if (stderr && stderr.length) {
      const errText = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : stderr;
      if (errText.trim()) process.stderr.write('[ps stderr] ' + errText + '\n');
    }
    const out = Buffer.isBuffer(stdout) ? stdout.toString('utf8') : stdout;
    return out.trim();
  } finally {
    try { fs.unlinkSync(tmp); } catch {}
  }
}

// Will be filled in by Task 2: a function that returns the PS script for
// download_attachments given { entryId, includeInline, downloadsRoot }.
let buildDownloadScript;
try {
  ({ buildDownloadScript } = require('../index.js'));
} catch (_) { /* index.js may not export it yet on first run */ }

let pass = 0, fail = 0, skip = 0;
async function t(name, fn) {
  try { await fn(); console.log('  ok  ' + name); pass++; }
  catch (e) {
    if (e && e.skip) { console.log('  skip ' + name + ' — ' + e.message); skip++; }
    else { console.log('  FAIL ' + name + '\n       ' + (e.stack || e.message)); fail++; }
  }
  await new Promise(r => setTimeout(r, 500));
}
function skipIf(cond, msg) { if (cond) { const e = new Error(msg); e.skip = true; throw e; } }

(async () => {
  console.log('attachment download live tests');

  await t('scaffold loads', async () => {
    skipIf(!buildDownloadScript, 'buildDownloadScript not exported yet (Task 2 implements it)');
  });

  await t('buildDownloadScript escapes single quotes in entry_id', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    const script = buildDownloadScript({
      entryId: "abc'def",
      includeInline: false,
      downloadsRoot: DOWNLOADS_ROOT,
    });
    if (!script.includes("abc''def")) throw new Error('entry_id quote not escaped: ' + script.slice(0, 200));
  });

  await t('buildDownloadScript embeds includeInline as $true/$false literal', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    const yes = buildDownloadScript({ entryId: 'x', includeInline: true,  downloadsRoot: DOWNLOADS_ROOT });
    const no  = buildDownloadScript({ entryId: 'x', includeInline: false, downloadsRoot: DOWNLOADS_ROOT });
    if (!/\$includeInline\s*=\s*\$true/.test(yes))  throw new Error('expected $includeInline = $true in yes-script');
    if (!/\$includeInline\s*=\s*\$false/.test(no))  throw new Error('expected $includeInline = $false in no-script');
  });

  await t('buildDownloadScript embeds the absolute downloadsRoot', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    const script = buildDownloadScript({ entryId: 'x', includeInline: false, downloadsRoot: 'C:\\tmp\\dl' });
    if (!script.includes("'C:\\tmp\\dl'")) throw new Error('downloadsRoot not embedded: ' + script.slice(0, 200));
  });

  // --- Live tests against Outlook ---
  // These require a running Outlook profile and seed env vars.

  const SEED_ENTRY_ID = process.env.ATTACH_TEST_ENTRY_ID;
  const INLINE_ONLY_ENTRY_ID = process.env.ATTACH_TEST_INLINE_ONLY_ENTRY_ID;
  const BAD_ENTRY_ID = '0000000000000000000000000000000000000000000000000000000000000000';

  async function callDownload(entryId, includeInline) {
    const script = buildDownloadScript({
      entryId,
      includeInline: !!includeInline,
      downloadsRoot: DOWNLOADS_ROOT,
    });
    const out = await ps(script);
    return JSON.parse(out);
  }

  await t('bad entry_id returns {error}', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    const r = await callDownload(BAD_ENTRY_ID, false);
    if (r.error !== 'Email not found') throw new Error('expected Email not found, got: ' + JSON.stringify(r));
  });

  await t('seed email saves at least one real attachment', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    skipIf(!SEED_ENTRY_ID, 'set ATTACH_TEST_ENTRY_ID env var');
    const r = await callDownload(SEED_ENTRY_ID, false);
    if (r.error) throw new Error('unexpected error: ' + r.error);
    if (!r.folder) throw new Error('expected folder path, got: ' + JSON.stringify(r));
    if (!Array.isArray(r.saved) || r.saved.length === 0) {
      throw new Error('expected at least one saved file: ' + JSON.stringify(r));
    }
    // Verify files exist on disk and sizes are non-zero
    for (const s of r.saved) {
      const full = path.join(r.folder, s.filename);
      if (!fs.existsSync(full)) throw new Error('file missing on disk: ' + full);
      const st = fs.statSync(full);
      if (st.size === 0 && !s.error) throw new Error('zero-byte file with no error: ' + full);
    }
    // Marker file should exist
    if (!fs.existsSync(path.join(r.folder, '.entry_id'))) {
      throw new Error('.entry_id marker missing in ' + r.folder);
    }
  });

  await t('inline-only email returns saved=[] with skipped_inline > 0', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    skipIf(!INLINE_ONLY_ENTRY_ID, 'set ATTACH_TEST_INLINE_ONLY_ENTRY_ID env var');
    const r = await callDownload(INLINE_ONLY_ENTRY_ID, false);
    if (r.error) throw new Error('unexpected error: ' + r.error);
    if (r.folder !== null) throw new Error('expected folder=null when nothing saved, got: ' + r.folder);
    if (!Array.isArray(r.saved) || r.saved.length !== 0) {
      throw new Error('expected saved=[], got: ' + JSON.stringify(r.saved));
    }
    if (typeof r.skipped_inline !== 'number' || r.skipped_inline < 1) {
      throw new Error('expected skipped_inline >= 1, got: ' + r.skipped_inline);
    }
  });

  await t('include_inline=true saves inline images too', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    skipIf(!INLINE_ONLY_ENTRY_ID, 'set ATTACH_TEST_INLINE_ONLY_ENTRY_ID env var');
    const r = await callDownload(INLINE_ONLY_ENTRY_ID, true);
    if (r.error) throw new Error('unexpected error: ' + r.error);
    if (!r.folder) throw new Error('expected folder path with include_inline=true');
    if (!Array.isArray(r.saved) || r.saved.length === 0) {
      throw new Error('expected at least one saved file with include_inline=true');
    }
  });

  await t('re-download is idempotent (same folder, no hash suffix)', async () => {
    skipIf(!buildDownloadScript, 'not implemented');
    skipIf(!SEED_ENTRY_ID, 'set ATTACH_TEST_ENTRY_ID env var');
    const r1 = await callDownload(SEED_ENTRY_ID, false);
    const r2 = await callDownload(SEED_ENTRY_ID, false);
    if (r1.folder !== r2.folder) throw new Error('expected same folder on re-download: ' + r1.folder + ' vs ' + r2.folder);
  });

  console.log(`\n${pass} passed, ${fail} failed, ${skip} skipped`);
  process.exit(fail > 0 ? 1 : 0);
})();
