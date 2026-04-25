# `query_emails` — Spec

Replaces `list_emails` + `search_emails` with one expressive tool. MongoDB-style filter tree → translated server-side to Outlook DASL `Items.Restrict()`.

## Scope

- **Replaces:** `list_emails`, `search_emails`
- **Keeps unchanged:** `list_accounts`, `read_email`, `send_email`, `reply_email`, `mark_as_read`, `list_calendar`
- **Reference:** `entry_id` remains the canonical email reference. No handle layer.

## Tool signature

```
query_emails(filter, fields?, folders?, accounts?, limit?, offset?, order_by?)
```

| Param | Type | Default |
|---|---|---|
| `filter` | object (operator tree, see below) | `{}` (matches all) |
| `fields` | string[] | see "Default fields" |
| `folders` | string[] | `["Inbox"]` |
| `accounts` | string[] | first account |
| `limit` | int | 20 (max 500) |
| `offset` | int | 0 |
| `order_by` | string | `"received_desc"` |

## Filter DSL

MongoDB-style. Combinators wrap field predicates.

**Combinators:** `$and`, `$or`, `$not`

**Field operators:**

| Op | Meaning | Applies to |
|---|---|---|
| `$eq`, `$ne` | exact match | all |
| `$in`, `$nin` | match any/none in list | all |
| `$contains`, `$not_contains` | substring (LIKE `%x%`) | string fields |
| `$starts_with`, `$ends_with` | prefix/suffix | string fields |
| `$gte`, `$lte`, `$gt`, `$lt` | comparison | dates, numbers |
| `$exists` | field is set / non-empty | all |

**Shorthand:** at top level, bare fields imply `$and` of `$eq` (or `$contains` for body/subject if value contains spaces — actually no, keep it strict: bare = `$eq`). Models can write the explicit form.

## Queryable fields

| Field | Type | Notes |
|---|---|---|
| `subject` | string | indexed, fast |
| `from` | string | sender SMTP address |
| `from_name` | string | display name |
| `to` | string | recipients (any match) |
| `cc` | string | |
| `body` | string | **slow** — full-text scan |
| `received` | date (ISO `YYYY-MM-DD` or `YYYY-MM-DDTHH:mm`) | |
| `sent` | date | |
| `unread` | bool | |
| `has_attachments` | bool | |
| `folder` | string | one of: Inbox, Sent, Drafts, Deleted, Outbox, Junk, or any custom folder name |
| `importance` | int | 0=low, 1=normal, 2=high |
| `size` | int | bytes |

## Default `fields`

- If `limit ≤ 20`: `["entry_id", "subject", "from", "received", "unread", "has_attachments"]`
- If `limit > 20`: `["subject", "from", "received"]` *(no entry_id — token-efficient scan mode)*

Agent can always override with explicit `fields`.

## Response shape

```json
{
  "results": [
    { "subject": "...", "from": "...", "received": "2026-04-22 14:30", ... }
  ],
  "total_returned": 42,
  "has_more": true,
  "next_offset": 50
}
```

## Examples

**Simple — last 20 from Lynn:**
```json
{ "filter": { "from": "lynn@kyros.com" } }
```

**Scan (no entry_ids):**
```json
{
  "filter": { "received": { "$gte": "2026-01-25" } },
  "fields": ["subject","from","received"],
  "limit": 200
}
```

**Boolean tree:**
```json
{
  "filter": {
    "$and": [
      { "from": { "$in": ["lynn@kyros.com","marina@kyros.com"] } },
      { "$or": [
        { "subject": { "$contains": "SSA" } },
        { "subject": { "$contains": "visa" } }
      ]},
      { "subject": { "$not_contains": "draft" } },
      { "received": { "$gte": "2026-01-25" } }
    ]
  },
  "limit": 100
}
```

**Two-phase (scan → hone):**
1. Scan to find candidate subjects (no entry_ids).
2. Narrow query on the specific subject + sender, `limit: 5` → entry_ids included by default.
3. Pass entry_id to `read_email` / `reply_email` / `mark_as_read`.

## DASL translator

Outlook `Items.Restrict(filter)` accepts SQL-ish strings:

```
@SQL="urn:schemas:httpmail:subject" LIKE '%SSA%'
  AND ("urn:schemas:httpmail:fromemail" = 'lynn@kyros.com'
       OR "urn:schemas:httpmail:fromemail" = 'marina@kyros.com')
  AND "urn:schemas:httpmail:datereceived" >= '2026-01-25 00:00'
```

**Field → DASL property map:**

| Field | DASL property |
|---|---|
| `subject` | `urn:schemas:httpmail:subject` |
| `from` | `urn:schemas:httpmail:fromemail` |
| `from_name` | `urn:schemas:httpmail:from` |
| `to` | `urn:schemas:httpmail:to` |
| `cc` | `urn:schemas:httpmail:cc` |
| `body` | `urn:schemas:httpmail:textdescription` (slow) |
| `received` | `urn:schemas:httpmail:datereceived` |
| `sent` | `urn:schemas:mailheader:date` |
| `unread` | `urn:schemas:httpmail:read` (negate) |
| `has_attachments` | `urn:schemas:httpmail:hasattachment` |
| `importance` | `urn:schemas:httpmail:importance` |

**Operator → DASL:**

- `$eq` → `=` (or `LIKE 'x'` for strings — DASL `=` is exact)
- `$ne` → `<>`
- `$in` → `(p = 'a' OR p = 'b' OR ...)` (parenthesized OR group)
- `$nin` → `(p <> 'a' AND p <> 'b' AND ...)`
- `$contains` → `LIKE '%x%'`
- `$not_contains` → `NOT (p LIKE '%x%')`
- `$starts_with` → `LIKE 'x%'`
- `$ends_with` → `LIKE '%x'`
- `$gte` → `>=`
- `$lte` → `<=`
- `$gt` → `>`
- `$lt` → `<`
- `$exists: true` → `p IS NOT NULL`
- `$and` / `$or` / `$not` → `AND` / `OR` / `NOT (...)`

**Translator implementation notes:**

- Recursive descent over the filter tree.
- Escape user strings: `'` → `''`. Reject control chars / nulls.
- Dates: parse ISO → render as `'YYYY-MM-DD HH:mm'` (Outlook accepts this; locale-stable enough on en-US Windows).
- `$contains` argument escaping: `%` and `_` are LIKE wildcards in DASL — escape with `[%]` / `[_]` to keep literal.
- Empty `filter: {}` → no Restrict call, return full sorted list.
- Cross-folder: iterate folders, run Restrict per folder, merge + sort + apply global limit/offset.

## Tool description (ships in MCP)

```
Query emails using a MongoDB-style filter. Replaces list_emails and search_emails.

FIELDS: subject, from, from_name, to, cc, body (slow), received, sent, unread,
has_attachments, folder, importance, size.

OPERATORS: $eq $ne $in $nin $contains $not_contains $starts_with $ends_with
$gte $lte $gt $lt $exists. Combinators: $and $or $not.

DATES: ISO format ("2026-01-25" or "2026-01-25T14:30").

DEFAULT FIELDS: small queries (limit ≤ 20) include entry_id; large scans omit
it for token efficiency. Override with `fields`.

PATTERN: For broad searches, scan first with limit > 20 (no entry_ids), then
re-query narrowly to get entry_ids for the specific items you want to act on.

EXAMPLE:
  filter: {
    $and: [
      { from: { $in: ["a@x.com","b@x.com"] } },
      { subject: { $contains: "report" } },
      { received: { $gte: "2026-01-01" } }
    ]
  }
```

## Out of scope (for now)

- Full-text body search optimization (slow path; document and move on)
- Calendar query DSL (`query_calendar`) — same pattern, separate iteration
- Raw DASL escape hatch — add only if a real query can't be expressed
- Custom folder enumeration tool (`list_folders`) — useful but not blocking

## Open questions

- Should `body` searches be opt-in (require `allow_slow: true`) to prevent accidental mailbox-wide scans? Lean yes.
- `order_by` values: `received_desc`, `received_asc`, `sent_desc`, `sent_asc`, `subject_asc`. Anything else worth supporting? Probably not.
- When `folders: ["*"]` (all folders), iteration cost is high. Cap at top-level folders by default; require explicit folder list to go deep.
