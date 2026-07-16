# email-mcp — Python COM rebuild (CLI + MCP)

- **Date**: 2026-07-15
- **Status**: approved, in implementation
- **Goal**: Replace the JS/PowerShell-templating implementation with a native Python package that drives Outlook via COM (`pywin32`), exposing the same 10 tools through **both** a Typer CLI and an MCP server (`email mcp` subcommand). Faithful behavior parity. Big-bang replace.

## Motivation (driver = all three)

The current `index.js` writes a fresh PowerShell script to a temp file and spawns `powershell.exe` **per tool call**, re-attaching to Outlook COM each time. This is:
- **Slow** — process spawn + COM re-attach on every call.
- **Fragile** — behavior is emitted as templated PowerShell strings (quoting, UTF-8, base64 body smuggling, path escaping).
- **Hard to grow/test** — no real CLI, no unit surface except the string builders.

Native Python COM removes the codegen layer entirely, holds one persistent Outlook handle, and gives a real CLI.

## Architecture: shared core + two thin adapters

```
tools/email-mcp/
  pyproject.toml            # uv-managed; console script: email = email_mcp.cli:app
  src/email_mcp/
    __init__.py
    errors.py               # typed exceptions
    models.py               # dataclasses/TypedDicts for tool returns
    query.py                # MongoDB-style filter tree -> DASL compiler (PURE, fully unit-tested)
    outlook/
      session.py            # COM session: attach/launch Outlook, MAPI namespace, account resolution
      worker.py             # dedicated STA thread owning the Outlook handle (server only)
      folders.py            # recursive mail-folder walk, calendar folder resolution
    ops/
      accounts.py           # list_accounts, account selection validation
      messages.py           # query_emails, read_email, send/reply/forward, mark_as_read
      attachments.py        # download_attachments (inline detection, slug foldering, collision/traversal guards)
      sendas.py             # PropertyAccessor MAPI sender rewrite (EXPERIMENTAL)
      sync.py               # force_sync (SyncObjects, best-effort)
      calendar.py           # list_calendar
    cli.py                  # Typer app: 10 subcommands + `mcp` subcommand
    server.py               # FastMCP app: 10 @mcp.tool wrappers + main()
  tests/
    test_query_compiler.py  # exhaustive, no Outlook needed
    test_account_select.py
    test_attachment_paths.py
    conftest.py
  README.md
  docs/superpowers/specs/2026-07-15-python-rewrite-design.md   # this file
```

**`core` (query/outlook/ops) knows nothing about CLI or MCP.** Both `cli.py` and `server.py` are thin translation layers over the same functions. Every tool is exercisable by calling core directly.

Ops functions take a live `OutlookSession` (or the MAPI namespace) as their first argument — they never construct COM themselves. This keeps them thread-agnostic and testable with a fake session.

## COM lifecycle & threading (the real architecture decision)

COM objects are apartment-bound. Every thread that touches COM must call `pythoncom.CoInitialize()`. The two adapters differ:

- **CLI (`cli.py`)**: one-shot per process. `CoInitialize()` on the main thread → do work → `CoUninitialize()`. No worker thread. Simple.
- **MCP server (`server.py`)**: long-running. FastMCP dispatches tool handlers off an event loop / threadpool, but the Outlook handle is apartment-bound. **Solution: one dedicated STA worker thread owns the single Outlook handle; every tool call is marshaled onto it via a request queue and the result is marshaled back.** This yields the perf win (persistent handle, no re-attach) *and* correctness (single apartment, no cross-thread COM access). The async `@mcp.tool` handlers submit a callable to the worker and `await` its result.

`OutlookSession` lazily attaches on first use: `win32com.client.GetActiveObject("Outlook.Application")`, falling back to `win32com.client.Dispatch("Outlook.Application")` if none running. If COM is unavailable (e.g. "new" Outlook only), raise `NoCOMAvailable` with a clear remediation message.

## Feature parity — all 10 tools

Behavior is ported **verbatim** from `index.js`. Notes below capture the exact semantics that must survive.

### `list_accounts`
Iterate `namespace.Folders` (store roots) → `[{name, entry_id}]`.

### `query_emails`
- Filter tree → DASL `@SQL=` clause via `query.py` (see below).
- Store selection: `account` substring-matches store `Name`; else all stores; fallback to first store.
- Recursively collect folders where `DefaultItemType == 0` (olMail).
- Per folder: COM `Items.Sort("[<prop>]", desc)`, then `Items.Restrict(dasl)` if filter present; accumulate `Count` into `total_matched`; take up to `offset+limit` items where `item.Class == 43` (olMail).
- Cross-folder re-sort when >1 folder, then slice `[offset : offset+limit]`.
- `order_by`: `received_desc|received_asc|sent_desc|sent_asc|subject_asc` → (`ReceivedTime|SentOn|Subject`, desc?).
- Default fields: explicit `fields` wins (validated against the output whitelist); else `limit<=20` → `[entry_id,subject,from,received,unread,has_attachments]`; else `[subject,from,received]`.
- Output-field whitelist: `entry_id,subject,from,from_name,to,cc,received,sent,unread,has_attachments,preview,importance,size`.
- Field extraction mirrors `Get-Field` (dates `yyyy-MM-dd HH:mm`, `preview` = first 150 chars of Body, etc.), each guarded.
- Returns `{results, total_returned, total_matched, has_more, next_offset}`. `limit` clamped `[1,500]`, `offset >= 0`.

### `read_email`
`GetItemFromID(entry_id)` → `{entry_id,subject,from,from_name,to,cc,received,sent,unread,has_attachments,attachments[],body}`. `received` as `yyyy-MM-dd HH:mm:ss`.

### `send_email`
- **Account required**: fetch accounts, validate `account` via case-insensitive substring against store name; error lists available accounts. (Ported from `validateAccountSelection`.)
- Subject = plain string; body wrapped as `<div>{textToHtml(body)}</div>` set on `HTMLBody`. `textToHtml` escapes `& < > "` and converts newlines to `<br>`. (In Python we set the property directly — no base64 smuggling needed.)
- Attachments: each path normalized (accept `/` or `\`, must be absolute, no `~`/env expansion), must exist and be a regular file, else abort before send (ported from `normalizeAttachmentPath`/`validateAttachments`).
- Sending account: resolve `ol.Session.Accounts` where `SmtpAddress` matches substring; set `SendUsingAccount`. (In win32com direct assignment may work; if COM ignores it as it does in JS, use the same reflection/`_oleobj_` fallback. Verify live.)
- Optional `send_as` (see below).
- Returns `{status:"sent", to, from[, sent_as]}`.

### `reply_email` / `forward_email`
- Same account validation + `SendUsingAccount` resolution + optional `send_as`.
- reply: `item.ReplyAll()` or `item.Reply()`; prepend `<div>{html}</div>` to `HTMLBody`.
- forward: `item.Forward()`; set `To`/`CC`; optionally prepend body. Preserves subject (FW:), quoted body, attachments (COM does this).

### `send_as` (EXPERIMENTAL — see Testing)
Resolve recipient via `ol.Session.CreateRecipient(addr).Resolve()`; if unresolved, error. Read `PR_ENTRYID (0x0FFF0102)` and `PR_SEARCH_KEY (0x300B0102)` off the recipient; overwrite on the outgoing item's `PropertyAccessor` both the `PR_SENT_REPRESENTING_*` set (`0x0042001F,0x0065001F,0x0064001F,0x00410102,0x003B0102`) and the `PR_SENDER_*` set (`0x0C1A001F,0x0C1F001F,0x0C1E001F,0x0C190102,0x0C1D0102`), then `item.Save()`. The binary props (`*0102`) are byte arrays — **win32com marshaling of byte-array PropertyAccessor values differs from PowerShell and must be verified**; this is the single riskiest port.

### `download_attachments`
- `GetItemFromID`; inline detection: read attachment CID via `PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F')`; treat as inline if `cid:<id>` appears in `HTMLBody` (or, plaintext-only body, any CID-bearing attachment). Skip inline unless `include_inline`.
- Folder name `YYYY-MM-DD_<sender-slug>_<subject-slug>` under `~/Downloads/email-attachments/`; slug = lowercase, `[^a-z0-9]+`→`-`, trimmed, capped (30 sender / 50 subject). `.entry_id` marker file for reuse/collision (SHA1[:8] disambiguation suffix). Windows-reserved char sanitize; filename collision `(2)`,`(3)`; path-traversal guard (resolved dest must stay under folder).
- Returns `{folder, saved:[{filename,size[,error]}], skipped_inline[, note]}`.

### `mark_as_read`
`GetItemFromID`; `item.UnRead = not read`; `Save()`. Default `read=true`.

### `force_sync`
`namespace.SyncObjects`; filter by `account` substring on name; `.Start()` each (best-effort, async — Outlook exposes no sync barrier); if none configured, `SendAndReceive(False)`. Grace sleep = `min(timeout,5000)ms`. `timeout_sec` clamp `[1,600]`. Returns `{ok, started:[...], timed_out:false, note}`.

### `list_calendar`
Calendar folder: account store's folder where `DefaultItemType==1`, else `GetDefaultFolder(9)`. `Items.IncludeRecurrences=True`, `Sort('[Start]')`, `Restrict("[Start] >= 'MM/dd/yyyy HH:mm' AND [Start] <= ...")`. `days` default 7, `count` default 20 (max 50). Returns `[{subject,start,end,location,organizer,all_day,body(≤200)}]`.

## query.py — DASL compiler (pure, the crown-jewel unit surface)

Direct port of `compileFilter`/`compilePredicate`/`compileOp`:
- Field→DASL map, type classes (STRING/DATE/BOOL/NUM), `SLOW_FIELDS={body}` gated behind `allow_slow`.
- `unread` maps to `read` and negates the value.
- Operators: `$eq $ne $in $nin $contains $not_contains $starts_with $ends_with $gte $lte $gt $lt $exists`.
- Combinators: `$and $or $not`. Empty `$and`→`1 = 1`, empty `$or`→`1 = 0`, empty filter→`''` (match all), empty `$in`→`1 = 0`, empty `$nin`→`1 = 1`.
- Bare value → `$eq`; bare array → `$in`; multiple ops on one field → AND.
- String escaping: `'`→`''`; LIKE args additionally escape `%`→`[%]`, `_`→`[_]`. Dates accept `YYYY-MM-DD[ T]HH:mm[:ss]` → `YYYY-MM-DD HH:mm`.
- `$contains`/`$not_contains`/`$starts_with`/`$ends_with` require string fields.
- Raises typed `InvalidFilter` on bad field/op/date/number/shape.

This module has **no COM dependency** and gets exhaustive table-driven unit tests mirroring the JS behavior exactly (including the negated-unread and escaping edge cases).

## Error handling

`errors.py` typed exceptions: `EmailMcpError` (base), `NoCOMAvailable`, `OutlookNotRunning`, `AccountNotFound`, `InvalidFilter`, `InvalidAttachmentPath`, `EmailNotFound`, `SendAsUnresolved`.
- CLI: catch → stderr message + non-zero exit; `--json` flag prints structured error.
- MCP: catch → tool error with the same helpful text (account-not-found lists configured accounts, etc.).
- "new Outlook, no COM" is detected at attach time and reported with remediation, never a raw traceback.

## Packaging & install (uvx-first)

- `pyproject.toml`: `[project.scripts] email = "email_mcp.cli:app"`; deps `pywin32`, `typer`, `mcp` (FastMCP); Windows-only classifier; Python `>=3.10`.
- Public zero-install run: `uvx --from git+https://github.com/kerodkibatu/email-mcp email mcp`.
- `.mcp.json` local wiring: `email` server → `command: uvx`, `args: ["--from", "git+https://github.com/kerodkibatu/email-mcp", "email", "mcp"]`, replacing the old `node index.js` entry. (Matches how the JS version used `npx github:`. First launch resolves from GitHub, cached after. NOTE: until the rebuild is pushed to GitHub, `.mcp.json` will be wired to the **local** source for testing: `uv run --project <path> email mcp`, then flipped to the git form once pushed.)
- Catalog: add a `tools/catalog.md` entry (stack, install, CLI, the `email mcp` MCP endpoint).

## Testing & acceptance (honest)

- **Unit (no Outlook)**: `query.py` compiler (exhaustive), account-selection validator, attachment-path normalizer. These are the real regression net.
- **Live smoke, in order** — read-only first: `list_accounts`, `query_emails`, `read_email`, `list_calendar`, `download_attachments`. Then, **gated behind explicit user go-ahead**, the real send: `kerod5858@gmail.com` → `abemkibatu101@gmail.com`, subject `[TEST] email-mcp python rebuild`. Then `reply_email`/`forward_email`/`mark_as_read`/`force_sync` against that thread.
- **`send_as` is EXPERIMENTAL / unverified**: it needs Exchange Send-As permission + same-tenant resolution and *cannot* be validated with a gmail→gmail harness; it is also the hardest bit to port (binary MAPI byte arrays). Ported carefully, documented as best-effort; **not** part of acceptance.

**Done =** all read-only tools return correct live data, the gated test send lands in `abemkibatu101@gmail.com` from `kerod5858@gmail.com`, unit tests pass, `.mcp.json` points at the Python server, JS artifacts removed, catalog + memory updated.

## Migration (big-bang)

1. Move repo `projects/personal/email-mcp` → `tools/email-mcp` (done; history preserved).
2. Build Python package alongside, get read-only tools green live.
3. Remove JS artifacts (`index.js`, `node_modules/`, `package.json`, `package-lock.json`, `test/*.js`), keep `LICENSE`/`.gitignore`/`.git`.
4. Rewrite `.mcp.json` `email` server to the Python launch.
5. Update `tools/catalog.md` + `_Memory/personal/projects/email-mcp.md`.
6. Commit within the email-mcp repo; push is a separate explicit step (do not auto-push the public repo).

## Build approach (ultracode)

Foundation (serial): `errors.py`, `models.py`, `outlook/session.py` + `worker.py` contract. Then fan out leaf modules (`query.py`+tests, `ops/*`, `cli.py`, `server.py`) via a Workflow, each agent given the relevant `index.js` line ranges + the foundation contract. **Live-Outlook integration + debug loop stays serial in the main session** — parallel agents can't each drive the one desktop Outlook instance, and COM debugging is iterative against real state.
