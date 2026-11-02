# msteams-local-mcp

Read the **new Microsoft Teams (v2)** message cache **locally** — and expose it
over the **Model Context Protocol (MCP)** for AI assistants (Claude, etc.).

**No Microsoft Graph. No OAuth. No Azure app registration. No network.** It only
reads what the signed-in Teams desktop client already keeps on your own disk.

## Why

Reading your Teams messages through Microsoft Graph requires `Chat.Read` /
`ChannelMessage.Read.All`, which many tenants gate behind **admin consent** — so
if you're an external/guest member of a client's tenant, you often simply can't.
But the desktop client still caches your recent conversations locally, in the
clear, for the account you're signed into. This tool reads that cache.

Existing forensic parsers (e.g. [`forensicsim`](https://github.com/lxndrblz/forensicsim),
[`teams-decoder`](https://github.com/Sec42/teams-decoder)) target the **older**
Teams (classic / Teams 1.x) schema and return nothing on the current client.
`msteams-local-mcp` maps the **current** Teams 2.x schema (`react-web-client`:
`replychain-manager` → `replychains` → `messageMap`) on top of the excellent
[`ccl_chromium_reader`](https://github.com/cclgroupltd/ccl_chromium_reader) for the
low-level Chromium LevelDB + V8 decoding. Nothing here is tied to any tenant or
account — it enumerates every context in the cache.

## How it works

The new Teams client (`com.microsoft.teams2` on macOS, `MSTeams` MSIX on Windows)
is an Edge WebView2 app. It stores recent conversations in an IndexedDB database
backed by Chromium LevelDB, values serialized in V8. This tool:

1. locates that LevelDB (auto-discovery per OS, or `MSTEAMS_LEVELDB=/path`),
2. copies it to a temp dir (the running client holds a file lock),
3. enumerates each `(tenant, user)` context and reads `replychains` / `conversations`,
4. yields plain-text messages (sender, timestamp, content), skipping the rare
   record that fails to deserialize.

## Install

```bash
pipx install "git+https://github.com/KamorionLabs/msteams-local-mcp"
# or: uv tool install "git+https://github.com/KamorionLabs/msteams-local-mcp"
```

Requires Python ≥ 3.10 and a signed-in new-Teams desktop client.

## Use — CLI

```bash
msteams-local-dump accounts                     # tenant/user contexts + inferred label
msteams-local-dump conversations                # chats/channels
msteams-local-dump search "quarterly budget"    # substring search across messages
msteams-local-dump search "PR" --account <tenantId:userId> --limit 20
```

## Use — MCP

Run the server (stdio):

```bash
msteams-local-mcp
```

Register it with an MCP client. Example (Claude Desktop / Claude Code):

```json
{
  "mcpServers": {
    "msteams-local": { "command": "msteams-local-mcp" }
  }
}
```

Tools exposed (all read-only):

| Tool | Purpose |
|------|---------|
| `overview` | who messaged you: recent activity grouped by conversation (best first call) |
| `recent_messages` | messages received in the last N days, newest first — one call for "what did I miss" |
| `mentions` | @-mentions of you (with content + read state); `unread_only` filter |
| `unread_messages` | your unread @-mentions (the only reliable unread signal — see note) |
| `search_messages` | substring search across messages, optional `days` recency |
| `read_conversation` | messages of one conversation (newest last) |
| `list_conversations` | chats/channels (id, title, type); 1:1 and unnamed group chats use the title Teams shows |
| `list_accounts` | the tenant/user contexts, with an inferred org label |

All accept an optional `account` filter (a key from `list_accounts`) to scope to one
tenant/org. **Note on unread:** the local cache stores no general per-message read
marker, so true "all unread" can't be derived; `unread_messages` returns unread
@-mentions, and `recent_messages` covers "everything received while I was away".

Performance: the LevelDB is parsed once and cached (in memory + on disk, keyed on
the store's signature), so chained calls are near-instant and re-parsing happens only
when Teams writes new data or the snapshot format changes.

## Scope, limits & ethics

- **Read-only.** This tool never writes to Teams or sends anything.
- **Your own data.** It reads the local cache of the account **you** are signed
  into. Use it only on machines and accounts you're authorized to. Respect your
  employer's / clients' policies and applicable law.
- **Recent only.** The cache holds synced/recent conversations, not full history,
  and no server-side search.
- **Schema drift.** Teams changes its internal schema between major versions; the
  mapping may need updates. PRs welcome.
- Platforms: developed and tested on **macOS** (Teams 2.x, 2026). Windows paths
  are included; Windows testing/PRs welcome.

## License

MIT — see [LICENSE](LICENSE). Builds on `ccl_chromium_reader` (MIT).
