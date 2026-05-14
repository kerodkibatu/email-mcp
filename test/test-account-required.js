'use strict';
// Unit tests for the `account` required + validated contract on send-style
// tools (send_email, reply_email, forward_email). Offline — uses a mocked
// account list and asserts on JSONSchema + validateAccountSelection.

const assert = require('assert');
const {
  validateAccountSelection,
  TOOLS,
} = require('../index.js');

let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); console.log('  ok  ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + (e.stack || e.message)); fail++; }
}

console.log('account required + validated');

// --- JSONSchema: account must be required on send_email, reply_email, forward_email ---

for (const toolName of ['send_email', 'reply_email', 'forward_email']) {
  t(`${toolName} declares account in required list`, () => {
    const tool = TOOLS.find(x => x.name === toolName);
    assert.ok(tool, `${toolName} not found`);
    const required = tool.inputSchema.required || [];
    assert.ok(required.includes('account'),
      `${toolName} must require 'account'. Got required: ${JSON.stringify(required)}`);
    const props = tool.inputSchema.properties;
    assert.ok(props.account, `${toolName} must declare account property`);
    assert.strictEqual(props.account.type, 'string');
  });
}

// --- validateAccountSelection: pure function tests ---

const MOCK_ACCOUNTS = [
  { name: 'kerod@towlydigital.com', entry_id: 'ABC1' },
  { name: 'kerod.kibatu@kyros.com', entry_id: 'ABC2' },
];

t('missing account (undefined) rejected with helpful error', () => {
  assert.throws(
    () => validateAccountSelection(undefined, MOCK_ACCOUNTS),
    err => /account is required/.test(err.message) && /towlydigital/.test(err.message) && /kyros/.test(err.message),
  );
});

t('empty-string account rejected', () => {
  assert.throws(
    () => validateAccountSelection('', MOCK_ACCOUNTS),
    /account is required/,
  );
});

t('whitespace-only account rejected', () => {
  assert.throws(
    () => validateAccountSelection('   ', MOCK_ACCOUNTS),
    /account is required/,
  );
});

t('unknown account rejected with available list', () => {
  assert.throws(
    () => validateAccountSelection('nonexistent@example.com', MOCK_ACCOUNTS),
    err => /Unknown account/.test(err.message) && /towlydigital/.test(err.message) && /kyros/.test(err.message),
  );
});

t('valid account accepted (full match)', () => {
  const m = validateAccountSelection('kerod@towlydigital.com', MOCK_ACCOUNTS);
  assert.strictEqual(m.entry_id, 'ABC1');
});

t('valid account accepted (substring match)', () => {
  const m = validateAccountSelection('towlydigital', MOCK_ACCOUNTS);
  assert.strictEqual(m.entry_id, 'ABC1');
});

t('substring match is case-insensitive', () => {
  const m = validateAccountSelection('KYROS', MOCK_ACCOUNTS);
  assert.strictEqual(m.entry_id, 'ABC2');
});

t('empty account list produces clear error', () => {
  assert.throws(
    () => validateAccountSelection('anything', []),
    /Unknown account.*\(none configured\)/,
  );
});

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
