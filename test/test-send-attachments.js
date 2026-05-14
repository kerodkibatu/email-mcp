'use strict';
// Unit tests for the attachment surface of send_email.
// Style mirrors test-translator.js: synchronous, assert-based, no live
// Outlook. Verifies JSONSchema, path validation, and PS script shape.

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  buildSendScript,
  normalizeAttachmentPath,
  validateAttachments,
  TOOLS,
} = require('../index.js');

let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); console.log('  ok  ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + (e.stack || e.message)); fail++; }
}

console.log('send_email attachments');

// --- JSONSchema ---

t('send_email tool advertises optional attachments array of strings', () => {
  const tool = TOOLS.find(x => x.name === 'send_email');
  assert.ok(tool, 'send_email tool not found');
  const props = tool.inputSchema.properties;
  assert.ok(props.attachments, 'attachments property missing');
  assert.strictEqual(props.attachments.type, 'array');
  assert.strictEqual(props.attachments.items.type, 'string');
  // Must remain optional — required list should not include attachments.
  const required = tool.inputSchema.required || [];
  assert.ok(!required.includes('attachments'), 'attachments must not be required');
});

// Regression: PR #4 added attachments support. Lock in that buildSendScript
// emits the Attachments.Add line so we never silently lose this wiring again.
t('buildSendScript regression: Attachments.Add is wired for non-empty list', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '',
    account: 'kerod@example.com',
    attachments: ['C:\\file.pdf'],
  });
  assert.ok(/\$mail\.Attachments\.Add\('C:\\file\.pdf'\)/.test(s),
    'expected $mail.Attachments.Add(\'C:\\\\file.pdf\') in script. Got:\n' + s);
});

// --- normalizeAttachmentPath ---

t('normalizeAttachmentPath rejects empty / non-string', () => {
  assert.throws(() => normalizeAttachmentPath(''));
  assert.throws(() => normalizeAttachmentPath(null));
  assert.throws(() => normalizeAttachmentPath(123));
});

t('normalizeAttachmentPath rejects relative paths', () => {
  assert.throws(() => normalizeAttachmentPath('relative\\file.txt'));
  assert.throws(() => normalizeAttachmentPath('./file.txt'));
});

t('normalizeAttachmentPath converts forward slashes to back slashes', () => {
  const out = normalizeAttachmentPath('C:/Users/Kerod/contract.pdf');
  assert.ok(!out.includes('/'), 'forward slashes should be normalized: ' + out);
  assert.ok(out.toLowerCase().includes('contract.pdf'));
});

// --- validateAttachments ---

t('validateAttachments returns [] for undefined / null', () => {
  assert.deepStrictEqual(validateAttachments(undefined), []);
  assert.deepStrictEqual(validateAttachments(null), []);
});

t('validateAttachments rejects non-array', () => {
  assert.throws(() => validateAttachments('C:\\file.pdf'));
  assert.throws(() => validateAttachments({}));
});

t('validateAttachments rejects missing files', () => {
  const bogus = path.join(os.tmpdir(), 'definitely-does-not-exist-' + Date.now() + '.pdf');
  assert.throws(() => validateAttachments([bogus]), /not found|Invalid/);
});

t('validateAttachments rejects directories', () => {
  assert.throws(() => validateAttachments([os.tmpdir()]), /not a regular file|Invalid/);
});

t('validateAttachments accepts a real file and returns normalized path', () => {
  const tmp = path.join(os.tmpdir(), 'send-attach-test-' + Date.now() + '.txt');
  fs.writeFileSync(tmp, 'hello');
  try {
    const out = validateAttachments([tmp.replace(/\\/g, '/')]);
    assert.strictEqual(out.length, 1);
    assert.ok(!out[0].includes('/'), 'normalized path should use backslashes');
  } finally {
    try { fs.unlinkSync(tmp); } catch {}
  }
});

// --- buildSendScript ---

t('buildSendScript omits Attachments.Add when none provided', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '',
    account: '',
    attachments: [],
  });
  assert.ok(!s.includes('Attachments.Add'), 'script should not call Attachments.Add when list empty');
});

t('buildSendScript emits one Attachments.Add per path', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '',
    account: '',
    attachments: ['C:\\one.pdf', 'C:\\two.pdf'],
  });
  const matches = s.match(/\$mail\.Attachments\.Add\(/g) || [];
  assert.strictEqual(matches.length, 2);
  assert.ok(s.includes("'C:\\one.pdf'"));
  assert.ok(s.includes("'C:\\two.pdf'"));
});

t('buildSendScript escapes single quotes in attachment paths', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '',
    account: '',
    attachments: ["C:\\O'Brien\\file.pdf"],
  });
  assert.ok(s.includes("'C:\\O''Brien\\file.pdf'"), 'single quote should be doubled');
});

t('buildSendScript still wires To / Subject / Body', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('Subject!', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>Body!</div>', 'utf8').toString('base64'),
    cc: 'c@d.com',
    account: '',
    attachments: [],
  });
  assert.ok(s.includes("$mail.To = 'a@b.com'"));
  assert.ok(s.includes("$mail.CC = 'c@d.com'"));
  assert.ok(s.includes('$mail.Send()'));
});

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
