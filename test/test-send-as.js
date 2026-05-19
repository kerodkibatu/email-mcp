'use strict';
// Unit tests for the send-as feature on send_email / reply_email / forward_email.
// Verifies tool schemas and the PowerShell snippet shape. No live Outlook.

const assert = require('assert');
const {
  buildSendScript,
  buildSendAsBlock,
  TOOLS,
} = require('../index.js');

let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); console.log('  ok  ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + (e.stack || e.message)); fail++; }
}

console.log('send_as');

// --- schema ---

for (const toolName of ['send_email', 'reply_email', 'forward_email']) {
  t(`${toolName} advertises optional send_as string`, () => {
    const tool = TOOLS.find(x => x.name === toolName);
    assert.ok(tool, toolName + ' not found');
    const props = tool.inputSchema.properties;
    assert.ok(props.send_as, 'send_as missing');
    assert.strictEqual(props.send_as.type, 'string');
    const required = tool.inputSchema.required || [];
    assert.ok(!required.includes('send_as'), 'send_as must not be required');
    assert.ok(/Send As/i.test(props.send_as.description), 'description should explain Send As');
    assert.ok(/permission/i.test(props.send_as.description), 'description should mention permission');
  });
}

// --- buildSendAsBlock ---

t('buildSendAsBlock returns empty string when sendAs is falsy', () => {
  assert.strictEqual(buildSendAsBlock('$mail', ''), '');
  assert.strictEqual(buildSendAsBlock('$mail', undefined), '');
  assert.strictEqual(buildSendAsBlock('$mail', null), '');
});

t('buildSendAsBlock emits both representing and sender prop sets', () => {
  const block = buildSendAsBlock('$mail', 'contact@x.com');
  // PR_SENT_REPRESENTING_NAME (0x0042001F) — represents
  assert.ok(block.includes('0x0042001F'), 'missing PR_SENT_REPRESENTING_NAME');
  // PR_SENDER_NAME (0x0C1A001F) — sender (suppresses "on behalf of")
  assert.ok(block.includes('0x0C1A001F'), 'missing PR_SENDER_NAME — without this the mail goes out as "on behalf of"');
  assert.ok(block.includes('$rcp.Resolve()'), 'must resolve target via Exchange');
  assert.ok(block.includes('$mail.PropertyAccessor'), 'must operate on the passed item var');
  assert.ok(block.includes('$mail.Save()'), 'must save after rewriting MAPI props');
});

t('buildSendAsBlock targets the item variable passed in', () => {
  const reply = buildSendAsBlock('$reply', 'contact@x.com');
  assert.ok(reply.includes('$reply.PropertyAccessor'));
  assert.ok(reply.includes('$reply.Save()'));
  assert.ok(!reply.includes('$mail.PropertyAccessor'));
});

t('buildSendAsBlock escapes single quotes in the address', () => {
  const block = buildSendAsBlock('$mail', "o'brien@x.com");
  assert.ok(block.includes("'o''brien@x.com'"), 'single quotes must be doubled');
});

// --- buildSendScript integration ---

t('buildSendScript without sendAs omits the send-as block entirely', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '', account: 'admin@x.com', attachments: [],
  });
  assert.ok(!s.includes('PR_SENT_REPRESENTING'), 'should not mention representing when no sendAs');
  assert.ok(!s.includes('0x0042001F'), 'should not set MAPI props when no sendAs');
  assert.ok(!s.includes('"sent_as"'), 'output JSON should omit sent_as field');
});

t('buildSendScript with sendAs injects the block and reports sent_as in output', () => {
  const s = buildSendScript({
    to: 'a@b.com',
    subjectB64: Buffer.from('hi', 'utf8').toString('base64'),
    htmlBodyB64: Buffer.from('<div>hi</div>', 'utf8').toString('base64'),
    cc: '', account: 'admin@x.com', attachments: [],
    sendAs: 'contact@x.com',
  });
  assert.ok(s.includes('0x0042001F'), 'PR_SENT_REPRESENTING_NAME missing');
  assert.ok(s.includes('0x0C1A001F'), 'PR_SENDER_NAME missing');
  assert.ok(s.includes('"sent_as":"contact@x.com"'), 'output JSON should include sent_as');
  // Block must come AFTER SendUsingAccount (so transport account is set first)
  // and BEFORE $mail.Send() (so the rewritten props are what actually go on the wire).
  const idxSendUsing = s.indexOf('SendUsingAccount');
  const idxRepresenting = s.indexOf('0x0042001F');
  const idxSend = s.indexOf('$mail.Send()');
  assert.ok(idxSendUsing >= 0 && idxRepresenting > idxSendUsing && idxSend > idxRepresenting,
    'order must be: SendUsingAccount → MAPI rewrite → Send()');
});

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
