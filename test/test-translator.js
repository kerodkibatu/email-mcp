'use strict';
const assert = require('assert');
const { compileFilter } = require('../index.js');

let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); console.log('  ok  ' + name); pass++; }
  catch (e) { console.log('  FAIL ' + name + '\n       ' + e.message); fail++; }
}

console.log('translator');

t('empty filter -> empty string', () => {
  assert.strictEqual(compileFilter({}, {}), '');
});

t('bare eq on string', () => {
  const out = compileFilter({ from: 'lynn@kyros.com' }, {});
  assert.match(out, /"urn:schemas:httpmail:fromemail" = 'lynn@kyros.com'/);
});

t('contains escapes single quote', () => {
  const out = compileFilter({ subject: { $contains: "O'Brien" } }, {});
  assert.match(out, /LIKE '%O''Brien%'/);
});

t('contains escapes LIKE wildcards', () => {
  const out = compileFilter({ subject: { $contains: '50%_off' } }, {});
  assert.match(out, /LIKE '%50\[%\]\[_\]off%'/);
});

t('$in expands to OR group', () => {
  const out = compileFilter({ from: { $in: ['a@x.com','b@x.com'] } }, {});
  assert.ok(out.includes("'a@x.com'") && out.includes("'b@x.com'") && out.includes(' OR '));
});

t('empty $in matches nothing', () => {
  const out = compileFilter({ from: { $in: [] } }, {});
  assert.strictEqual(out, '1 = 0');
});

t('date $gte formats correctly', () => {
  const out = compileFilter({ received: { $gte: '2026-01-25' } }, {});
  assert.match(out, /"urn:schemas:httpmail:datereceived" >= '2026-01-25 00:00'/);
});

t('date with time', () => {
  const out = compileFilter({ received: { $gte: '2026-01-25T14:30' } }, {});
  assert.match(out, />= '2026-01-25 14:30'/);
});

t('invalid date rejected', () => {
  assert.throws(() => compileFilter({ received: { $gte: 'yesterday' } }, {}));
});

t('$and combines with AND', () => {
  const out = compileFilter({
    $and: [
      { from: 'a@x.com' },
      { subject: { $contains: 'foo' } }
    ]
  }, {});
  assert.ok(out.includes(' AND '));
});

t('$or combines with OR', () => {
  const out = compileFilter({
    $or: [
      { subject: { $contains: 'a' } },
      { subject: { $contains: 'b' } }
    ]
  }, {});
  assert.ok(out.includes(' OR '));
});

t('$not wraps with NOT', () => {
  const out = compileFilter({ $not: { subject: { $contains: 'spam' } } }, {});
  assert.match(out, /^NOT \(/);
});

t('nested boolean tree', () => {
  const out = compileFilter({
    $and: [
      { from: { $in: ['a@x.com','b@x.com'] } },
      { $or: [
        { subject: { $contains: 'SSA' } },
        { subject: { $contains: 'visa' } }
      ]},
      { subject: { $not_contains: 'draft' } },
      { received: { $gte: '2026-01-25' } }
    ]
  }, {});
  // sanity: contains all the bits
  assert.ok(out.includes('SSA'));
  assert.ok(out.includes('visa'));
  assert.ok(out.includes('NOT'));
  assert.ok(out.includes('draft'));
  assert.ok(out.includes('2026-01-25'));
});

t('unread=true negates to read=0', () => {
  const out = compileFilter({ unread: true }, {});
  assert.match(out, /"urn:schemas:httpmail:read" = 0/);
});

t('unread=false negates to read=1', () => {
  const out = compileFilter({ unread: false }, {});
  assert.match(out, /"urn:schemas:httpmail:read" = 1/);
});

t('has_attachments bool', () => {
  const out = compileFilter({ has_attachments: true }, {});
  assert.match(out, /"urn:schemas:httpmail:hasattachment" = 1/);
});

t('body requires allow_slow', () => {
  assert.throws(() => compileFilter({ body: { $contains: 'foo' } }, {}));
  const out = compileFilter({ body: { $contains: 'foo' } }, { allow_slow: true });
  assert.match(out, /textdescription/);
});

t('unknown field rejected', () => {
  assert.throws(() => compileFilter({ banana: 'yes' }, {}));
});

t('unknown operator rejected', () => {
  assert.throws(() => compileFilter({ subject: { $weird: 'x' } }, {}));
});

t('$contains rejected on non-string field', () => {
  assert.throws(() => compileFilter({ received: { $contains: '2026' } }, {}));
});

t('multi-op on same field ANDs them', () => {
  const out = compileFilter({ received: { $gte: '2026-01-01', $lte: '2026-04-25' } }, {});
  assert.ok(out.includes('>=') && out.includes('<=') && out.includes('AND'));
});

t('$exists true / false', () => {
  assert.match(compileFilter({ cc: { $exists: true } }, {}), /IS NOT NULL/);
  assert.match(compileFilter({ cc: { $exists: false } }, {}), /IS NULL/);
});

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
