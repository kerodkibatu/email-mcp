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
| `query_emails` | MongoDB-style querying across all folders (`has_attachments`, `unread`, etc.) |
| `read_email` | Read full body of an email by EntryID |
| `send_email` | Send a new mail (optionally from a specific account, with file attachments) |
| `reply_email` | Reply / Reply All to an email |
| `forward_email` | Forward an email |
| `download_attachments` | Save real attachments to `~/Downloads/email-attachments/YYYY-MM-DD_<sender>_<subject>/` |
| `force_sync` | Trigger Send/Receive and block until all sync groups finish (or timeout) |
| `mark_as_read` | Flip read/unread state |
| `list_calendar` | List upcoming calendar events |

### Choosing the Sending Account

`send_email`, `reply_email`, and `forward_email` **require** an `account` parameter — a substring of the configured Outlook account name (typically the SMTP address). This is intentional: with multiple accounts configured (e.g. personal + work), defaulting to Outlook's primary account is a footgun — it's how personal mail leaks out of a work account or vice versa. Forcing the caller to name the account makes the send explicit.

If the supplied `account` doesn't match any configured account (case-insensitive substring), the tool errors and lists the available accounts. Run `list_accounts` first if you don't already know the name.

```json
{
  "to": "client@example.com",
  "subject": "Status update",
  "body": "Heads up — ...",
  "account": "kerod@towlydigital.com"
}
```

### Sending Attachments

`send_email` accepts an optional `attachments` array of absolute file paths. Each path must exist and point to a regular file; if any path is invalid, the tool returns an error and does not send. Forward and back slashes are both accepted on Windows; `~` and environment variables are **not** expanded — pass fully resolved paths.

```json
{
  "to": "kerod@example.com",
  "subject": "Signed contract",
  "body": "See attached.",
  "account": "kerod@towlydigital.com",
  "attachments": [
    "C:\\Users\\Kerod\\Desktop\\contract.pdf",
    "C:/Users/Kerod/Desktop/cover-letter.pdf"
  ]
}
```

### Downloading Attachments

The `download_attachments` tool extracts files from an email and saves them locally, returning the absolute folder path.
- **Location:** Files are saved under the user's Downloads folder: `~/Downloads/email-attachments/YYYY-MM-DD_sender-slug_subject-slug/`.
- **Inline Images:** Logos and signature images are filtered out by default to avoid clutter. Set `include_inline: true` if you specifically need them.
- **Idempotency:** Re-running the tool on the same email safely skips re-downloading if the files already exist.

## How it works

Each tool call writes a short PowerShell script to a temp file and executes it. The script attaches to a running Outlook instance via `Marshal.GetActiveObject('Outlook.Application')`, or launches one if none is running, then drives the MAPI namespace to read/write mail.

This means:
- Outlook must be installed (doesn't need to be open, but first call will launch it)
- Whatever account is signed into Outlook is what the MCP sees — no separate auth
- Sent mail appears in the user's Sent folder exactly as if they sent it manually

## License

MIT — see [LICENSE](LICENSE).
