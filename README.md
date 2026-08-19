# email-mcp

Email + calendar via **classic Outlook on Windows**, driven through native COM (pywin32). Ships as a single `email` command that is both a **CLI** and a **stdio MCP server**. No Azure app registration, no OAuth, no PowerShell — it just drives the Outlook desktop client you're already signed into.

## Requirements

- **Windows** 10/11
- **Classic Outlook desktop**, configured with at least one account
- **Python 3.10+** and [`uv`](https://docs.astral.sh/uv/)

> The "new" Outlook for Windows does **not** expose COM. If you're on the new Outlook and can't switch back to classic, this won't work for you.

COM is reached only through the Outlook desktop client — there is no separate authentication. Whatever account is signed into Outlook is what the tool sees, and sent mail lands in the real Sent folder exactly as if you sent it by hand.

## Install / Run

Run straight from GitHub with `uvx` (no clone, no manual install):

```bash
uvx --from git+https://github.com/kerodkibatu/email-mcp email --help
```

`uv` resolves and caches the package on first run; later invocations are fast.

### CLI

Every tool is a subcommand of `email`. Output is JSON on stdout.

```bash
# List configured Outlook accounts
uvx --from git+https://github.com/kerodkibatu/email-mcp email list-accounts

# Query across all mail folders with a MongoDB-style filter (JSON string)
uvx --from git+https://github.com/kerodkibatu/email-mcp email query \
  --filter '{"$and":[{"from":{"$contains":"@kyros.com"}},{"unread":true}]}' \
  --limit 20 --order-by received_desc

# Read one email by EntryID
uvx --from git+https://github.com/kerodkibatu/email-mcp email read --entry-id "0000000..."

# Send a new mail (account is REQUIRED — see below)
uvx --from git+https://github.com/kerodkibatu/email-mcp email send \
  --to client@example.com \
  --subject "Status update" \
  --body "Heads up — ..." \
  --account kerod@towlydigital.com

# Save a new mail as a draft instead of sending (same flags as send, no --send-as)
uvx --from git+https://github.com/kerodkibatu/email-mcp email draft \
  --to client@example.com \
  --subject "Status update" \
  --body "Heads up — ..." \
  --account kerod@towlydigital.com
```

### MCP

`email mcp` boots a stdio MCP server exposing all 11 tools. Add it to your `.mcp.json` (or `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "email": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/kerodkibatu/email-mcp", "email", "mcp"]
    }
  }
}
```

The first launch is slower while `uv` resolves the package; subsequent launches hit the cache.

## Tools

| Tool | Purpose |
|------|---------|
| `list_accounts` | List configured Outlook accounts |
| `query_emails` | MongoDB-style querying across all folders (`has_attachments`, `unread`, etc.) |
| `read_email` | Read full body of an email by EntryID |
| `send_email` | Send a new mail (optionally from a specific account, with file attachments) |
| `draft_email` | Compose a new mail and save it as a draft instead of sending |
| `reply_email` | Reply / Reply All to an email |
| `forward_email` | Forward an email |
| `download_attachments` | Save real attachments to `~/Downloads/email-attachments/YYYY-MM-DD_<sender>_<subject>/` |
| `force_sync` | Trigger Send/Receive and wait briefly for sync groups to make headway |
| `mark_as_read` | Flip read/unread state |
| `list_calendar` | List upcoming calendar events |

### Choosing the Sending Account

`send_email`, `reply_email`, and `forward_email` **require** an `account` parameter — a substring of the configured Outlook account name (typically the SMTP address). This is intentional: with multiple accounts configured (e.g. personal + work), defaulting to Outlook's primary account is a footgun — it's how personal mail leaks out of a work account or vice versa. Forcing the caller to name the account makes the send explicit.

If the supplied `account` doesn't match any configured account (case-insensitive substring), the tool errors and lists the available accounts. Run `list_accounts` first if you don't already know the name. The transport account is set directly on the Outlook item via `SendUsingAccount`.

```json
{
  "to": "client@example.com",
  "subject": "Status update",
  "body": "Heads up — ...",
  "account": "kerod@towlydigital.com"
}
```

### Send As (EXPERIMENTAL — Exchange "Send As" permission required)

> **EXPERIMENTAL.** This drives classic Outlook through pywin32 COM only — there is no Azure/OAuth path. Behavior depends on tenant policy and it may bounce, silently downgrade, or leave the message in the Outbox. Treat it as best-effort.

`send_email`, `reply_email`, and `forward_email` accept an optional `send_as` parameter — an SMTP address to send AS. The recipient sees that address as the From, with **no** "on behalf of" disclosure. This is the true Exchange Send As, distinct from Send on Behalf.

Mechanically, the tool sets `SendUsingAccount` to the `account` you provide (the transport mailbox), then resolves `send_as` against Exchange and overwrites both the `PR_SENT_REPRESENTING_*` and `PR_SENDER_*` MAPI properties (via the item's `PropertyAccessor`) to point at that address before submitting. Overwriting the sender props — not just the representing props — is what collapses "on behalf of" into a pure Send As. Exchange validates the permission at submit time.

Requirements:
- The `account` user must have **Send As** permission on the `send_as` mailbox, granted server-side by an Exchange admin. The tool cannot grant or check this — it can only attempt the send.
- The `send_as` address must be resolvable by Exchange — typically a mailbox in the same tenant. External addresses (gmail.com, etc.) will fail at resolution with an error.
- If the permission is missing, behavior depends on tenant policy: Exchange may bounce the message, silently downgrade it to "on behalf of", or leave it in the Outbox.
- This feature only works inside an Exchange organization that has explicitly authorized it. It cannot be used to spoof external senders.

```json
{
  "to": "client@example.com",
  "subject": "Status update",
  "body": "Heads up — ...",
  "account": "admin@custom.com",
  "send_as": "contact@custom.com"
}
```

The response includes a `sent_as` field echoing the address when `send_as` was used.

### Sending Attachments

`send_email` accepts an optional `attachments` array of absolute file paths. Each path must exist and point to a regular file; if any path is invalid, the tool returns an error listing the offenders and does not send. Forward and back slashes are both accepted on Windows; `~` and environment variables are **not** expanded — pass fully resolved paths.

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
- **Idempotency:** Re-running the tool on the same email safely reuses the folder (tracked via an `.entry_id` marker) and disambiguates when two different emails slug to the same name.

## How it works

The `email` command attaches to a running Outlook instance via COM (falling back to launching one if none is running), then drives the MAPI namespace to read and write mail. All COM access goes through a single Outlook session; the tool never spawns PowerShell.

This means:
- Outlook must be installed (it doesn't need to be open — the first call will launch it)
- Whatever account is signed into Outlook is what the tool sees — no separate auth
- Sent mail appears in the user's Sent folder exactly as if they sent it manually

## License

MIT — see [LICENSE](LICENSE).
