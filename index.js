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
      { timeout: 45000 }
    );
    if (stderr && stderr.trim()) process.stderr.write('[ps stderr] ' + stderr + '\n');
    return stdout.trim();
  } finally {
    try { fs.unlinkSync(tmp); } catch {}
  }
}

// Shared Outlook COM init block â€” reused in every script
const INIT = `
$ErrorActionPreference = 'Stop'
try {
  $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  $ol = New-Object -ComObject Outlook.Application
}
$ns = $ol.GetNamespace('MAPI')
`;

const TOOLS = [
  {
    name: 'list_accounts',
    description: 'List all email accounts configured in Outlook',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'list_emails',
    description: 'List recent emails from a folder. Defaults to Inbox of first account.',
    inputSchema: {
      type: 'object',
      properties: {
        account: { type: 'string', description: 'Account email address (optional, defaults to first account)' },
        folder: { type: 'string', description: 'Folder name: Inbox, Sent, Drafts, Deleted (default: Inbox)' },
        count: { type: 'number', description: 'Number of emails to return (default: 20, max: 50)' },
        unread_only: { type: 'boolean', description: 'Return only unread emails (default: false)' },
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
        entry_id: { type: 'string', description: 'The EntryID of the email (from list_emails)' },
      },
    },
  },
  {
    name: 'search_emails',
    description: 'Search emails across all folders by subject, sender, or body keyword',
    inputSchema: {
      type: 'object',
      required: ['query'],
      properties: {
        query: { type: 'string', description: 'Search term to look for in subject, sender, or body' },
        account: { type: 'string', description: 'Limit search to a specific account email (optional)' },
        count: { type: 'number', description: 'Max results to return (default: 15)' },
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
$accounts | ConvertTo-Json -Depth 2
`);
      return result || '[]';
    }

    case 'list_emails': {
      const folder = (args.folder || 'Inbox').replace(/'/g, "''");
      const count = Math.min(args.count || 20, 50);
      const account = args.account ? args.account.replace(/'/g, "''") : '';
      const unreadOnly = args.unread_only ? 'true' : 'false';

      const result = await ps(`
${INIT}
$folderMap = @{ 'Inbox'=6; 'Sent'=5; 'Drafts'=16; 'Deleted'=3; 'Outbox'=4; 'Junk'=23 }
$folderName = '${folder}'
$targetAccount = '${account}'
$unreadOnly = $${unreadOnly}
$count = ${count}

if ($targetAccount) {
  $store = $ns.Folders | Where-Object { $_.Name -like "*$targetAccount*" } | Select-Object -First 1
  if (-not $store) { $store = $ns.Folders.Item(1) }
} else {
  $store = $ns.Folders.Item(1)
}

if ($folderMap.ContainsKey($folderName)) {
  $mf = $store.Folders | Where-Object { $folderMap[$folderName] -ne $null } | Select-Object -First 1
  # Try by well-known index first
  try {
    $mf = $ns.GetDefaultFolder($folderMap[$folderName])
  } catch {
    $mf = $store.Folders | Where-Object { $_.Name -like "*$folderName*" } | Select-Object -First 1
  }
} else {
  $mf = $store.Folders | Where-Object { $_.Name -like "*$folderName*" } | Select-Object -First 1
}

if (-not $mf) { Write-Output '[]'; exit 0 }

$items = $mf.Items
$items.Sort('[ReceivedTime]', $true)

$emails = @()
$i = 1
foreach ($item in $items) {
  if ($emails.Count -ge $count) { break }
  if ($unreadOnly -and $item.UnRead -eq $false) { $i++; continue }
  if ($item.Class -ne 43) { $i++; continue }  # 43 = olMail
  $emails += [PSCustomObject]@{
    entry_id      = $item.EntryID
    subject       = $item.Subject
    from          = $item.SenderEmailAddress
    from_name     = $item.SenderName
    received      = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm')
    unread        = $item.UnRead
    has_attachments = $item.Attachments.Count -gt 0
    preview       = $item.Body.Substring(0, [Math]::Min(200, $item.Body.Length)).Trim()
  }
  $i++
}
$emails | ConvertTo-Json -Depth 2
`);
      return result || '[]';
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
} | ConvertTo-Json -Depth 3
`);
    }

    case 'search_emails': {
      const query = args.query.replace(/'/g, "''").replace(/"/g, '""');
      const count = Math.min(args.count || 15, 30);
      const account = args.account ? args.account.replace(/'/g, "''") : '';

      return await ps(`
${INIT}
$query = '${query}'
$maxResults = ${count}
$targetAccount = '${account}'
$results = @()

$stores = if ($targetAccount) {
  $ns.Folders | Where-Object { $_.Name -like "*$targetAccount*" }
} else {
  $ns.Folders
}

foreach ($store in $stores) {
  foreach ($folder in $store.Folders) {
    if ($results.Count -ge $maxResults) { break }
    try {
      $filter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%${query}%' OR ""urn:schemas:httpmail:fromemail"" LIKE '%${query}%' OR ""urn:schemas:httpmail:textdescription"" LIKE '%${query}%'"
      $found = $folder.Items.Restrict($filter)
      foreach ($item in $found) {
        if ($results.Count -ge $maxResults) { break }
        if ($item.Class -ne 43) { continue }
        $results += [PSCustomObject]@{
          entry_id  = $item.EntryID
          subject   = $item.Subject
          from      = $item.SenderEmailAddress
          from_name = $item.SenderName
          received  = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm')
          folder    = $folder.Name
          unread    = $item.UnRead
          preview   = $item.Body.Substring(0, [Math]::Min(150, $item.Body.Length)).Trim()
        }
      }
    } catch {}
  }
  if ($results.Count -ge $maxResults) { break }
}
$results | ConvertTo-Json -Depth 2
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
$mail = $ol.CreateItem(0)  # 0 = olMailItem
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
$reply.Body = '${body}' + "\r\n\r\n" + $reply.Body
$reply.Send()
Write-Output '{"status":"sent","reply_to":"' + $item.Subject + '"}'
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
$results | ConvertTo-Json -Depth 2
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

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write('email MCP server running\n');
}

main().catch((err) => {
  process.stderr.write('Fatal: ' + err.message + '\n');
  process.exit(1);
});
