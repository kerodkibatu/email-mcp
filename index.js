#!/usr/bin/env node
'use strict';

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
} = require('@modelcontextprotocol/sdk/types.js');
const { exec } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const os = require('os');
const path = require('path');

const execAsync = promisify(exec);

async function ps(script) {
  const tmp = path.join(os.tmpdir(), `email-mcp-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
  fs.writeFileSync(tmp, script, 'utf8');
  try {
    const { stdout, stderr } = await execAsync(
      `powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "${tmp}"`,
      { timeout: 60000, maxBuffer: 32 * 1024 * 1024 }
    );
    if (stderr && stderr.trim()) process.stderr.write('[ps stderr] ' + stderr + '\n');
    return stdout.trim();
  } finally {
    try { fs.unlinkSync(tmp); } catch {}
  }
}

const INIT = `
$ErrorActionPreference = 'Stop'
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  $ol = New-Object -ComObject Outlook.Application
}
$ns = $ol.GetNamespace('MAPI')
`;

// ---------- query_emails: filter tree -> DASL ----------

const FIELD_DASL = {
  subject:         'urn:schemas:httpmail:subject',
  from:            'urn:schemas:httpmail:fromemail',
  from_name:       'urn:schemas:httpmail:from',
  to:              'urn:schemas:httpmail:to',
  cc:              'urn:schemas:httpmail:cc',
  body:            'urn:schemas:httpmail:textdescription',
  received:        'urn:schemas:httpmail:datereceived',
  sent:            'urn:schemas:mailheader:date',
  unread:          'urn:schemas:httpmail:read',  // negated; unread=true => read=0
  has_attachments: 'urn:schemas:httpmail:hasattachment',
  importance:      'urn:schemas:httpmail:importance',
  size:            'urn:schemas:httpmail:size',
};

const STRING_FIELDS = new Set(['subject', 'from', 'from_name', 'to', 'cc', 'body']);
const DATE_FIELDS   = new Set(['received', 'sent']);
const BOOL_FIELDS   = new Set(['unread', 'has_attachments']);
const NUM_FIELDS    = new Set(['importance', 'size']);

const SLOW_FIELDS   = new Set(['body']);

function escSqlString(s) {
  return String(s).replace(/'/g, "''");
}
function escLikeArg(s) {
  // DASL LIKE wildcards: % and _; escape with [%] [_]
  return escSqlString(s).replace(/%/g, '[%]').replace(/_/g, '[_]');
}
function fmtDate(v) {
  // accept ISO 'YYYY-MM-DD' or 'YYYY-MM-DDTHH:mm[:ss]'
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) throw new McpError(ErrorCode.InvalidParams, `Invalid date: ${v} (use YYYY-MM-DD or YYYY-MM-DDTHH:mm)`);
  const [, y, mo, d, hh = '00', mm = '00'] = m;
  return `${y}-${mo}-${d} ${hh}:${mm}`;
}
function fmtBool(v) { return v ? '1' : '0'; }
function fmtNum(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) throw new McpError(ErrorCode.InvalidParams, `Invalid number: ${v}`);
  return String(n);
}

function fmtValue(field, v) {
  if (STRING_FIELDS.has(field)) return `'${escSqlString(v)}'`;
  if (DATE_FIELDS.has(field))   return `'${fmtDate(v)}'`;
  if (BOOL_FIELDS.has(field))   return fmtBool(v);
  if (NUM_FIELDS.has(field))    return fmtNum(v);
  throw new McpError(ErrorCode.InvalidParams, `Unsupported field: ${field}`);
}

// `unread` is stored as `read` (negated). Translate the operator/value when emitting.
function negateBoolValue(v) { return v ? '0' : '1'; }

function compilePredicate(field, opOrVal, opts) {
  const dasl = FIELD_DASL[field];
  if (!dasl) throw new McpError(ErrorCode.InvalidParams, `Unknown field: ${field}`);
  if (SLOW_FIELDS.has(field) && !opts.allow_slow) {
    throw new McpError(ErrorCode.InvalidParams,
      `Field '${field}' is slow (full-text scan). Pass allow_slow: true to opt in.`);
  }

  const prop = `"${dasl}"`;
  const isUnread = field === 'unread';

  // Bare value -> $eq
  if (opOrVal === null || typeof opOrVal !== 'object' || Array.isArray(opOrVal)) {
    if (Array.isArray(opOrVal)) {
      // bare array -> $in
      return compileOp(prop, field, '$in', opOrVal, isUnread);
    }
    return compileOp(prop, field, '$eq', opOrVal, isUnread);
  }

  const keys = Object.keys(opOrVal);
  if (keys.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, `Empty operator object for field: ${field}`);
  }
  // multiple operators on same field => AND them
  const parts = keys.map(op => compileOp(prop, field, op, opOrVal[op], isUnread));
  return parts.length === 1 ? parts[0] : `(${parts.join(' AND ')})`;
}

function compileOp(prop, field, op, val, isUnread) {
  const v = (raw) => fmtValue(field, raw);
  const vUnread = (raw) => negateBoolValue(raw);

  switch (op) {
    case '$eq':
      if (isUnread) return `${prop} = ${vUnread(val)}`;
      return `${prop} = ${v(val)}`;
    case '$ne':
      if (isUnread) return `${prop} <> ${vUnread(val)}`;
      return `${prop} <> ${v(val)}`;
    case '$in': {
      if (!Array.isArray(val) || val.length === 0)
        return `1 = 0`; // empty $in matches nothing
      const parts = val.map(x => `${prop} = ${isUnread ? vUnread(x) : v(x)}`);
      return `(${parts.join(' OR ')})`;
    }
    case '$nin': {
      if (!Array.isArray(val) || val.length === 0) return `1 = 1`;
      const parts = val.map(x => `${prop} <> ${isUnread ? vUnread(x) : v(x)}`);
      return `(${parts.join(' AND ')})`;
    }
    case '$contains':
      if (!STRING_FIELDS.has(field)) throw new McpError(ErrorCode.InvalidParams, `$contains requires string field: ${field}`);
      return `${prop} LIKE '%${escLikeArg(val)}%'`;
    case '$not_contains':
      if (!STRING_FIELDS.has(field)) throw new McpError(ErrorCode.InvalidParams, `$not_contains requires string field: ${field}`);
      return `NOT (${prop} LIKE '%${escLikeArg(val)}%')`;
    case '$starts_with':
      if (!STRING_FIELDS.has(field)) throw new McpError(ErrorCode.InvalidParams, `$starts_with requires string field: ${field}`);
      return `${prop} LIKE '${escLikeArg(val)}%'`;
    case '$ends_with':
      if (!STRING_FIELDS.has(field)) throw new McpError(ErrorCode.InvalidParams, `$ends_with requires string field: ${field}`);
      return `${prop} LIKE '%${escLikeArg(val)}'`;
    case '$gte': return `${prop} >= ${v(val)}`;
    case '$lte': return `${prop} <= ${v(val)}`;
    case '$gt':  return `${prop} > ${v(val)}`;
    case '$lt':  return `${prop} < ${v(val)}`;
    case '$exists':
      return val ? `${prop} IS NOT NULL` : `${prop} IS NULL`;
    default:
      throw new McpError(ErrorCode.InvalidParams, `Unknown operator: ${op}`);
  }
}

function compileFilter(node, opts) {
  if (!node || typeof node !== 'object' || Array.isArray(node)) {
    throw new McpError(ErrorCode.InvalidParams, `Filter must be an object`);
  }
  const keys = Object.keys(node);
  if (keys.length === 0) return ''; // empty filter = match all

  const parts = keys.map(k => {
    if (k === '$and' || k === '$or') {
      if (!Array.isArray(node[k])) throw new McpError(ErrorCode.InvalidParams, `${k} requires an array`);
      if (node[k].length === 0) return k === '$and' ? '1 = 1' : '1 = 0';
      const sub = node[k].map(c => compileFilter(c, opts)).filter(s => s.length > 0);
      if (sub.length === 0) return '';
      const joiner = k === '$and' ? ' AND ' : ' OR ';
      return sub.length === 1 ? sub[0] : `(${sub.map(s => `(${s})`).join(joiner)})`;
    }
    if (k === '$not') {
      const inner = compileFilter(node[k], opts);
      if (!inner) return '';
      return `NOT (${inner})`;
    }
    // field predicate
    return compilePredicate(k, node[k], opts);
  }).filter(s => s.length > 0);

  if (parts.length === 0) return '';
  return parts.length === 1 ? parts[0] : parts.map(p => `(${p})`).join(' AND ');
}

// ---------- query_emails: PowerShell builder ----------

const ORDER_MAP = {
  received_desc: { prop: 'ReceivedTime', desc: true  },
  received_asc:  { prop: 'ReceivedTime', desc: false },
  sent_desc:     { prop: 'SentOn',       desc: true  },
  sent_asc:      { prop: 'SentOn',       desc: false },
  subject_asc:   { prop: 'Subject',      desc: false },
};

const VALID_FIELDS_OUT = new Set([
  'entry_id','subject','from','from_name','to','cc','received','sent',
  'unread','has_attachments','preview','importance','size'
]);

function psString(s) {
  // for single-quoted PS string literal: escape single quotes
  return `'${String(s).replace(/'/g, "''")}'`;
}
function psStringArray(arr) {
  return '@(' + arr.map(psString).join(',') + ')';
}

function buildQueryScript({ daslFilter, account, limit, offset, order, fields }) {
  const sortInfo = ORDER_MAP[order] || ORDER_MAP.received_desc;
  const sortDesc = sortInfo.desc ? '$true' : '$false';
  const sortProp = sortInfo.prop;

  // Use a here-string for the DASL so we don't have to worry about embedded quotes.
  const filterClause = daslFilter ? `@SQL=${daslFilter}` : '';
  const filterHere = filterClause
    ? `@'\n${filterClause}\n'@`
    : `''`;

  return `${INIT}
$ErrorActionPreference = 'Continue'
$accountFilter = ${psString(account || '')}
$daslFilter = ${filterHere}
$limit = ${Math.max(1, Math.min(500, limit))}
$offset = ${Math.max(0, offset)}
$sortProp = ${psString(sortProp)}
$sortDesc = ${sortDesc}
$fieldsList = ${psStringArray(fields)}

$stores = if ($accountFilter) {
  @($ns.Folders | Where-Object { $_.Name -like "*$accountFilter*" })
} else {
  @($ns.Folders)
}
if (-not $stores -or $stores.Count -eq 0) { $stores = @($ns.Folders.Item(1)) }

# Recursively collect all mail folders (DefaultItemType = 0 = olMail)
function Get-MailFolders($parent) {
  $list = New-Object System.Collections.ArrayList
  try {
    foreach ($f in $parent.Folders) {
      try {
        if ($f.DefaultItemType -eq 0) { [void]$list.Add($f) }
      } catch {}
      try {
        foreach ($sub in (Get-MailFolders $f)) { [void]$list.Add($sub) }
      } catch {}
    }
  } catch {}
  return ,$list
}

$mailFolders = New-Object System.Collections.ArrayList
foreach ($store in $stores) {
  foreach ($mf in (Get-MailFolders $store)) { [void]$mailFolders.Add($mf) }
}

$allItems = New-Object System.Collections.ArrayList
$totalMatched = 0

foreach ($f in $mailFolders) {
    $items = $f.Items
    try { $items.Sort("[$sortProp]", $sortDesc) } catch {}

    if ($daslFilter -and $daslFilter.Trim().Length -gt 0) {
      try { $items = $items.Restrict($daslFilter) } catch {
        [System.Console]::Error.WriteLine("Restrict failed in folder " + $f.Name + ": " + $_.Exception.Message)
        continue
      }
    }

    try { $totalMatched += [int]$items.Count } catch {}

    # Take only what we need from each folder (already sorted via COM Sort)
    $perFolderCap = $offset + $limit
    $taken = 0
    foreach ($item in $items) {
      if ($taken -ge $perFolderCap) { break }
      try {
        if ($item.Class -ne 43) { continue }   # 43 = olMail
        [void]$allItems.Add($item)
        $taken++
      } catch {}
    }
}

# Cross-folder final sort (Restrict+Sort was per-folder)
if ($mailFolders.Count -le 1) {
  $sorted = $allItems
} else {
  try {
    if ($sortProp -eq 'ReceivedTime') {
      $sorted = if ($sortDesc) { $allItems | Sort-Object -Property ReceivedTime -Descending } else { $allItems | Sort-Object -Property ReceivedTime }
    } elseif ($sortProp -eq 'SentOn') {
      $sorted = if ($sortDesc) { $allItems | Sort-Object -Property SentOn -Descending } else { $allItems | Sort-Object -Property SentOn }
    } else {
      $sorted = if ($sortDesc) { $allItems | Sort-Object -Property Subject -Descending } else { $allItems | Sort-Object -Property Subject }
    }
  } catch { $sorted = $allItems }
}

$sliced = @($sorted) | Select-Object -Skip $offset -First $limit

function Get-Field($i, $f) {
  switch ($f) {
    'entry_id'        { try { return $i.EntryID } catch { return '' } }
    'subject'         { try { return $i.Subject } catch { return '' } }
    'from'            { try { return $i.SenderEmailAddress } catch { return '' } }
    'from_name'       { try { return $i.SenderName } catch { return '' } }
    'to'              { try { return $i.To } catch { return '' } }
    'cc'              { try { return $i.CC } catch { return '' } }
    'received'        { try { if ($i.ReceivedTime) { return $i.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } } catch {}; return '' }
    'sent'            { try { if ($i.SentOn) { return $i.SentOn.ToString('yyyy-MM-dd HH:mm') } } catch {}; return '' }
    'unread'          { try { return [bool]$i.UnRead } catch { return $false } }
    'has_attachments' { try { return ($i.Attachments.Count -gt 0) } catch { return $false } }
    'preview'         { try { if ($i.Body) { return $i.Body.Substring(0, [Math]::Min(150, $i.Body.Length)).Trim() } } catch {}; return '' }
    'importance'      { try { return [int]$i.Importance } catch { return 1 } }
    'size'            { try { return [int]$i.Size } catch { return 0 } }
  }
}

$results = @()
foreach ($i in $sliced) {
  $obj = [ordered]@{}
  foreach ($f in $fieldsList) {
    $obj[$f] = Get-Field $i $f
  }
  $results += [PSCustomObject]$obj
}

$hasMore = $totalMatched -gt ($offset + $limit)
$nextOffset = if ($hasMore) { $offset + $limit } else { $null }

$output = [PSCustomObject]@{
  results = @($results)
  total_returned = @($results).Count
  total_matched = $totalMatched
  has_more = $hasMore
  next_offset = $nextOffset
}
$output | ConvertTo-Json -Depth 4 -Compress
`;
}

// ---------- TOOLS ----------

const TOOLS = [
  {
    name: 'list_accounts',
    description: 'List all email accounts configured in Outlook',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'query_emails',
    description: `Query emails using a MongoDB-style filter tree. Searches across ALL mail folders by default (Inbox, Sent, Archive, custom, etc.).

FIELDS (queryable): subject, from, from_name, to, cc, body (slow, requires allow_slow=true), received, sent, unread, has_attachments, importance, size.

OPERATORS: $eq $ne $in $nin $contains $not_contains $starts_with $ends_with $gte $lte $gt $lt $exists.
COMBINATORS: $and $or $not.

DATES: ISO format ("2026-01-25" or "2026-01-25T14:30").

DEFAULT FIELDS in response: small queries (limit <= 20) include entry_id; larger scans omit it for token efficiency. Override with the 'fields' param.

PATTERN: For broad searches, scan first with limit > 20 (no entry_ids), then re-query narrowly to get entry_ids for the specific items you want to act on (read/reply/mark).

EXAMPLE filter:
  { "$and": [
      { "from": { "$in": ["a@x.com","b@x.com"] } },
      { "subject": { "$contains": "report" } },
      { "received": { "$gte": "2026-01-01" } }
  ]}`,
    inputSchema: {
      type: 'object',
      properties: {
        filter: { type: 'object', description: 'Filter tree. Empty {} matches all.' },
        fields: { type: 'array', items: { type: 'string' }, description: 'Output fields. Default depends on limit.' },
        account: { type: 'string', description: 'Account email (substring match). Default: all accounts.' },
        limit: { type: 'number', description: 'Max results (default 20, max 500)' },
        offset: { type: 'number', description: 'Pagination offset' },
        order_by: { type: 'string', enum: Object.keys(ORDER_MAP), description: 'Sort order' },
        allow_slow: { type: 'boolean', description: 'Required to query the body field' },
      },
    },
  },
  {
    name: 'read_email',
    description: 'Read the full content of an email by its EntryID',
    inputSchema: {
      type: 'object',
      required: ['entry_id'],
      properties: {
        entry_id: { type: 'string', description: 'The EntryID of the email (from query_emails)' },
      },
    },
  },
  {
    name: 'send_email',
    description: 'Send a new email',
    inputSchema: {
      type: 'object',
      required: ['to', 'subject', 'body'],
      properties: {
        to: { type: 'string', description: 'Recipient email address (comma-separated for multiple)' },
        cc: { type: 'string', description: 'CC recipients (comma-separated, optional)' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body (plain text)' },
        account: { type: 'string', description: 'Send from this account email address (optional, uses default)' },
      },
    },
  },
  {
    name: 'reply_email',
    description: 'Reply to an email',
    inputSchema: {
      type: 'object',
      required: ['entry_id', 'body'],
      properties: {
        entry_id: { type: 'string', description: 'EntryID of the email to reply to' },
        body: { type: 'string', description: 'Reply body text' },
        reply_all: { type: 'boolean', description: 'Reply to all recipients (default: false)' },
      },
    },
  },
  {
    name: 'mark_as_read',
    description: 'Mark an email as read or unread',
    inputSchema: {
      type: 'object',
      required: ['entry_id'],
      properties: {
        entry_id: { type: 'string', description: 'EntryID of the email' },
        read: { type: 'boolean', description: 'true = mark read, false = mark unread (default: true)' },
      },
    },
  },
  {
    name: 'list_calendar',
    description: 'List upcoming calendar events',
    inputSchema: {
      type: 'object',
      properties: {
        account: { type: 'string', description: 'Account email address (optional)' },
        days: { type: 'number', description: 'How many days ahead to look (default: 7)' },
        count: { type: 'number', description: 'Max events to return (default: 20)' },
      },
    },
  },
];

// ---------- handlers ----------

async function handleTool(name, args) {
  switch (name) {
    case 'list_accounts': {
      const result = await ps(`
${INIT}
$accounts = @()
foreach ($store in $ns.Folders) {
  $accounts += [PSCustomObject]@{
    name = $store.Name
    entry_id = $store.EntryID
  }
}
$accounts | ConvertTo-Json -Depth 2 -Compress
`);
      return result || '[]';
    }

    case 'query_emails': {
      const filter = args.filter || {};
      const limit  = typeof args.limit === 'number' ? args.limit : 20;
      const offset = typeof args.offset === 'number' ? args.offset : 0;
      const order  = args.order_by || 'received_desc';
      const account = args.account || '';
      const allowSlow = !!args.allow_slow;

      // default fields
      let fields;
      if (Array.isArray(args.fields) && args.fields.length > 0) {
        const bad = args.fields.filter(f => !VALID_FIELDS_OUT.has(f));
        if (bad.length) throw new McpError(ErrorCode.InvalidParams, `Unknown output field(s): ${bad.join(', ')}. Valid: ${[...VALID_FIELDS_OUT].join(', ')}`);
        fields = args.fields;
      } else if (limit <= 20) {
        fields = ['entry_id','subject','from','received','unread','has_attachments'];
      } else {
        fields = ['subject','from','received'];
      }

      const dasl = compileFilter(filter, { allow_slow: allowSlow });

      const script = buildQueryScript({
        daslFilter: dasl,
        account,
        limit,
        offset,
        order,
        fields,
      });

      return await ps(script);
    }

    case 'read_email': {
      const entryId = args.entry_id.replace(/'/g, "''");
      return await ps(`
${INIT}
$item = $ns.GetItemFromID('${entryId}')
if (-not $item) { Write-Output '{"error":"Email not found"}'; exit 0 }
[PSCustomObject]@{
  entry_id        = $item.EntryID
  subject         = $item.Subject
  from            = $item.SenderEmailAddress
  from_name       = $item.SenderName
  to              = $item.To
  cc              = $item.CC
  received        = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
  sent            = if ($item.SentOn) { $item.SentOn.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
  unread          = $item.UnRead
  has_attachments = $item.Attachments.Count -gt 0
  attachments     = @($item.Attachments | ForEach-Object { $_.FileName })
  body            = $item.Body
} | ConvertTo-Json -Depth 3 -Compress
`);
    }

    case 'send_email': {
      const to = args.to.replace(/'/g, "''");
      const subject = args.subject.replace(/'/g, "''");
      const body = args.body.replace(/'/g, "''");
      const cc = (args.cc || '').replace(/'/g, "''");
      const account = (args.account || '').replace(/'/g, "''");

      return await ps(`
${INIT}
$mail = $ol.CreateItem(0)
$mail.Subject = '${subject}'
$mail.Body = '${body}'
$mail.To = '${to}'
${cc ? `$mail.CC = '${cc}'` : ''}
${account ? `
$accts = $ol.Session.Accounts
foreach ($acct in $accts) {
  if ($acct.SmtpAddress -like '*${account}*') {
    $mail.SendUsingAccount = $acct
    break
  }
}` : ''}
$mail.Send()
Write-Output '{"status":"sent","to":"${to}","subject":"${subject}"}'
`);
    }

    case 'reply_email': {
      const entryId = args.entry_id.replace(/'/g, "''");
      const body = args.body.replace(/'/g, "''");
      const replyAll = args.reply_all ? 'true' : 'false';

      return await ps(`
${INIT}
$item = $ns.GetItemFromID('${entryId}')
if (-not $item) { Write-Output '{"error":"Email not found"}'; exit 0 }
$reply = if ($${replyAll}) { $item.ReplyAll() } else { $item.Reply() }
$reply.Body = '${body}' + "\`r\`n\`r\`n" + $reply.Body
$reply.Send()
Write-Output '{"status":"sent"}'
`);
    }

    case 'mark_as_read': {
      const entryId = args.entry_id.replace(/'/g, "''");
      const read = args.read !== false;

      return await ps(`
${INIT}
$item = $ns.GetItemFromID('${entryId}')
if (-not $item) { Write-Output '{"error":"Email not found"}'; exit 0 }
$item.UnRead = $${!read}
$item.Save()
Write-Output '{"status":"ok","unread":${!read}}'
`);
    }

    case 'list_calendar': {
      const days = args.days || 7;
      const count = Math.min(args.count || 20, 50);
      const account = (args.account || '').replace(/'/g, "''");

      return await ps(`
${INIT}
$days = ${days}
$maxItems = ${count}
$targetAccount = '${account}'
$start = Get-Date
$end = $start.AddDays($days)

if ($targetAccount) {
  $store = $ns.Folders | Where-Object { $_.Name -like "*$targetAccount*" } | Select-Object -First 1
  if (-not $store) { $store = $ns.Folders.Item(1) }
  $cal = $store.Folders | Where-Object { $_.DefaultItemType -eq 1 } | Select-Object -First 1
  if (-not $cal) { $cal = $ns.GetDefaultFolder(9) }
} else {
  $cal = $ns.GetDefaultFolder(9)
}

$items = $cal.Items
$items.IncludeRecurrences = $true
$items.Sort('[Start]')
$filter = "[Start] >= '$($start.ToString('MM/dd/yyyy HH:mm'))' AND [Start] <= '$($end.ToString('MM/dd/yyyy HH:mm'))'"
$events = $items.Restrict($filter)

$results = @()
foreach ($e in $events) {
  if ($results.Count -ge $maxItems) { break }
  $results += [PSCustomObject]@{
    subject    = $e.Subject
    start      = $e.Start.ToString('yyyy-MM-dd HH:mm')
    end        = $e.End.ToString('yyyy-MM-dd HH:mm')
    location   = $e.Location
    organizer  = $e.Organizer
    all_day    = $e.AllDayEvent
    body       = $e.Body.Substring(0, [Math]::Min(200, $e.Body.Length)).Trim()
  }
}
$results | ConvertTo-Json -Depth 2 -Compress
`);
    }

    default:
      throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
  }
}

const server = new Server(
  { name: 'email', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  try {
    const result = await handleTool(name, args || {});
    return { content: [{ type: 'text', text: result }] };
  } catch (err) {
    if (err instanceof McpError) throw err;
    throw new McpError(ErrorCode.InternalError, err.message || String(err));
  }
});

// ---------- exports for testing ----------
module.exports = { compileFilter, buildQueryScript };

if (require.main === module) {
  (async () => {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    process.stderr.write('email MCP server running\n');
  })().catch((err) => {
    process.stderr.write('Fatal: ' + err.message + '\n');
    process.exit(1);
  });
}
