# email-mcp

MCP server for email + calendar via **classic Outlook on Windows** (COM automation). No Azure app registration, no OAuth — it just drives the Outlook desktop client you're already signed into.

## Requirements

- Windows 10/11
- Classic Outlook desktop (configured with at least one account)
- Node.js 14+
- PowerShell (built-in)

> The "new" Outlook for Windows does **not** expose COM. If you're on the new Outlook and can't switch back, this won't work for you.

## Install / Run

Via `npx` straight from GitHub (no clone):

```bash
npx -y github:kerodkibatu/email-mcp
```

Or clone and run locally:

```bash
git clone https://github.com/kerodkibatu/email-mcp
cd email-mcp
npm install
node index.js
```

## Claude Code / Claude Desktop config

Add to your `.mcp.json` (or `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "email": {
      "command": "npx",
      "args": ["-y", "github:kerodkibatu/email-mcp"]
    }
  }
}
```

First launch will be slow (~10s) while npx clones + installs. Subsequent launches hit the cache.

## Tools

| Tool | Purpose |
|------|---------|
| `list_accounts` | List configured Outlook accounts |
| `list_emails` | List recent mail from a folder (Inbox/Sent/Drafts/Deleted/Outbox/Junk) |
| `read_email` | Read full body of an email by EntryID |
| `search_emails` | Search subject/sender/body across folders |
| `send_email` | Send a new mail (optionally from a specific account) |
| `reply_email` | Reply / Reply All to an email |
| `mark_as_read` | Flip read/unread state |
| `list_calendar` | List upcoming calendar events |

## How it works

Each tool call writes a short PowerShell script to a temp file and executes it. The script attaches to a running Outlook instance via `Marshal.GetActiveObject('Outlook.Application')`, or launches one if none is running, then drives the MAPI namespace to read/write mail.

This means:
- Outlook must be installed (doesn't need to be open, but first call will launch it)
- Whatever account is signed into Outlook is what the MCP sees — no separate auth
- Sent mail appears in the user's Sent folder exactly as if they sent it manually

## License

MIT — see [LICENSE](LICENSE).
