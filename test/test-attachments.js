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

  console.log(`\n${pass} passed, ${fail} failed, ${skip} skipped`);
  process.exit(fail > 0 ? 1 : 0);
})();
