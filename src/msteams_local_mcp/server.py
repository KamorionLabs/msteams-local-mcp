"""MCP server exposing the local Microsoft Teams (v2) message cache read-only.

Runs over stdio. Tools:
  - ``list_accounts``      the (tenant, user) contexts present in the cache
  - ``list_conversations`` chats/channels, optionally per account
  - ``read_conversation``  messages of one conversation, newest last
  - ``search_messages``    substring search across cached messages (optional recency)
  - ``recent_messages``    messages received in the last N days (one call)

All data is read from the local disk only (no Graph, no network). Set
``MSTEAMS_LEVELDB`` to point at a specific LevelDB directory; otherwise the cache
is auto-discovered for the current OS.

Performance: the LevelDB is cold-copied and fully parsed ONCE, then cached in
memory for ``MSTEAMS_CACHE_TTL`` seconds (default 180). Without this, every tool
call re-copied ~30 MB and re-parsed ~thousands of messages — a single agent
request that chained many calls took minutes. With the cache, the first call pays
the parse cost and the rest are in-memory.
"""
from __future__ import annotations

import dataclasses
import glob
import json
import os
import pathlib
import re
import tempfile
import time
from collections import Counter
from typing import Optional

# The high-level server class was FastMCP in the mcp SDK 1.x and was renamed
# MCPServer in 2.0 (same API: name arg, .tool() decorator, .run() defaulting to
# stdio). Support both so the package works regardless of the installed SDK.
try:
    from mcp.server.fastmcp import FastMCP as _MCPServer  # mcp < 2
except ImportError:  # pragma: no cover
    from mcp.server.mcpserver import MCPServer as _MCPServer  # mcp >= 2

from .reader import Message, TeamsCacheReader, find_cache

mcp = _MCPServer("msteams-local")

_LEVELDB = os.environ.get("MSTEAMS_LEVELDB") or None
_TTL = float(os.environ.get("MSTEAMS_CACHE_TTL", "180"))
_MAXLEN = int(os.environ.get("MSTEAMS_MAX_CONTENT", "600"))  # truncate long bodies in output
_CACHE_FILE = pathlib.Path(
    os.environ.get("MSTEAMS_CACHE_DIR")
    or (pathlib.Path.home() / ".cache" / "msteams-local-mcp")
) / "snapshot.json"
# Bump whenever the parsed output changes, so stale on-disk snapshots are discarded.
_SNAPSHOT_VERSION = 2

# In-memory snapshot: parse once, reuse for _TTL seconds across tool calls.
_snap: dict = {
    "ts": 0.0, "sig": "", "messages": [], "conversations": [],
    "accounts": [], "mentions": [], "skipped": 0,
}


def _signature(leveldb: str) -> str:
    """Cheap fingerprint of the LevelDB — newest mtime + total size. Changes only
    when Teams actually writes new data, so we re-parse only then."""
    files = glob.glob(os.path.join(leveldb, "*"))
    if not files:
        return ""
    newest = max(os.path.getmtime(f) for f in files)
    total = sum(os.path.getsize(f) for f in files)
    return f"{newest:.0f}:{total}"


def _load_disk(sig: str) -> Optional[dict]:
    """Return a previously-parsed snapshot if its signature still matches."""
    try:
        data = json.loads(_CACHE_FILE.read_text())
    except (OSError, ValueError):
        return None
    if data.get("sig") != sig or data.get("version") != _SNAPSHOT_VERSION:
        return None
    data["messages"] = [Message(**m) for m in data.get("messages", [])]
    return data


def _save_disk(snap: dict) -> None:
    try:
        _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": _SNAPSHOT_VERSION,
            "sig": snap["sig"],
            "messages": [dataclasses.asdict(m) for m in snap["messages"]],
            "conversations": snap["conversations"],
            "accounts": snap["accounts"],
            "mentions": snap["mentions"],
            "skipped": snap["skipped"],
        }
        tmp = _CACHE_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload))
        os.replace(tmp, _CACHE_FILE)
        os.chmod(_CACHE_FILE, 0o600)  # holds message content — not world-readable
    except OSError:
        pass


def _epoch_ms(ts: str) -> Optional[float]:
    """Teams timestamps are epoch milliseconds (as strings like '1768557031836.0')."""
    if not ts:
        return None
    try:
        return float(ts)
    except (TypeError, ValueError):
        try:
            from datetime import datetime

            return datetime.fromisoformat(ts.replace("Z", "+00:00")).timestamp() * 1000
        except Exception:
            return None


def _snapshot() -> dict:
    """Return the parsed cache. Three tiers, cheapest first:
    1. in-memory, if still within the TTL and the LevelDB hasn't changed;
    2. on-disk parse, if the LevelDB signature matches (no re-parse, survives
       process/gateway restarts);
    3. full cold-copy + parse, only when Teams actually wrote new data.
    """
    now = time.monotonic()
    leveldb = _LEVELDB or find_cache()
    sig = _signature(leveldb) if leveldb else ""

    if _snap["messages"] and now - _snap["ts"] < _TTL and _snap["sig"] == sig:
        return _snap

    disk = _load_disk(sig) if sig else None
    if disk:
        _snap.update(ts=now, **disk)
        return _snap

    with TeamsCacheReader(leveldb) as r:
        messages = list(r.messages())
        conversations = list(r.conversations())
        mentions = list(r.mentions())
        skipped = r.skipped
    # Derive accounts (+ inferred org label) from the single messages pass.
    per_acc: dict[str, Counter] = {}
    for m in messages:
        c = per_acc.setdefault(m.account, Counter())
        mm = re.search(r"\(([^)]+)\)\s*$", m.sender)  # e.g. "Name (GP Rubix)"
        if mm:
            c[mm.group(1)] += 1
    accounts = [
        {
            "key": k,
            "tenant_id": k.split(":")[0] if ":" in k else k,
            "user_id": k.split(":")[1] if ":" in k else "",
            "label": (c.most_common(1)[0][0] if c else ""),
        }
        for k, c in per_acc.items()
    ]
    _snap.update(
        ts=now, sig=sig, messages=messages, conversations=conversations,
        accounts=accounts, mentions=mentions, skipped=skipped,
    )
    _save_disk(_snap)
    return _snap


def _msg_index(snap: dict) -> dict:
    """(account, message_id) -> Message, to attach content to mentions."""
    return {(m.account, m.message_id): m for m in snap["messages"]}


def _title_map(snap: dict) -> dict:
    """(account, conversation_id) -> display title (falls back to '')."""
    return {(c["account"], c["id"]): c.get("title", "") for c in snap["conversations"]}


def _fmt(m: Message) -> dict:
    content = m.content if len(m.content) <= _MAXLEN else m.content[:_MAXLEN] + "…"
    return {
        "sender": m.sender,
        "timestamp": m.timestamp,
        "content": content,
        "conversation_id": m.conversation_id,
        "account": m.account,
    }


@mcp.tool()
def list_accounts() -> list[dict]:
    """List the Teams accounts (tenant/user contexts) in the local cache.

    Each: ``key`` (use as the ``account`` filter elsewhere), ``tenant_id``,
    ``user_id`` and a best-effort ``label`` inferred from org names in messages.
    """
    return _snapshot()["accounts"]


@mcp.tool()
def list_conversations(account: Optional[str] = None, limit: int = 100) -> list[dict]:
    """List chats/channels (id, title, type), optionally filtered by ``account``."""
    out = []
    for conv in _snapshot()["conversations"]:
        if account and conv.get("account") != account:
            continue
        if conv.get("title") or conv.get("last_message"):
            out.append(conv)
        if len(out) >= limit:
            break
    return out


@mcp.tool()
def read_conversation(conversation_id: str, limit: int = 50, account: Optional[str] = None) -> list[dict]:
    """Return up to ``limit`` most recent messages of a conversation (newest last)."""
    msgs = [
        m
        for m in _snapshot()["messages"]
        if m.conversation_id == conversation_id and (not account or m.account == account)
    ]
    msgs.sort(key=lambda m: _epoch_ms(m.timestamp) or 0)
    return [_fmt(m) for m in msgs[-limit:]]


@mcp.tool()
def search_messages(
    query: str, account: Optional[str] = None, days: Optional[int] = None, limit: int = 50
) -> list[dict]:
    """Case-insensitive substring search across cached messages.

    Optional ``days`` restricts to messages received in the last N days.
    """
    q = query.lower()
    cutoff = (time.time() - days * 86400) * 1000 if days else None
    hits = []
    for m in _snapshot()["messages"]:
        if account and m.account != account:
            continue
        if cutoff is not None:
            e = _epoch_ms(m.timestamp)
            if e is None or e < cutoff:
                continue
        if q in m.content.lower() or q in m.sender.lower():
            hits.append(_fmt(m))
            if len(hits) >= limit:
                break
    return hits


@mcp.tool()
def recent_messages(days: int = 7, account: Optional[str] = None, limit: int = 200) -> list[dict]:
    """Messages received in the last ``days`` days, newest first — in ONE call.

    Ideal for 'what did I get while I was away'. Filter by ``account`` (see
    ``list_accounts``) to scope to one tenant/org.
    """
    cutoff = (time.time() - days * 86400) * 1000
    hits = []
    for m in _snapshot()["messages"]:
        if account and m.account != account:
            continue
        e = _epoch_ms(m.timestamp)
        if e is None or e < cutoff:
            continue
        hits.append((e, m))
    hits.sort(key=lambda t: t[0], reverse=True)
    return [_fmt(m) for _e, m in hits[:limit]]


def _mentions_impl(days, unread_only, account, limit):
    snap = _snapshot()
    idx = _msg_index(snap)
    titles = _title_map(snap)
    cutoff = (time.time() - days * 86400) * 1000 if days else None
    out = []
    for mt in snap["mentions"]:
        if account and mt["account"] != account:
            continue
        if unread_only and mt["is_read"]:
            continue
        e = _epoch_ms(mt["timestamp"])
        if cutoff is not None and (e is None or e < cutoff):
            continue
        m = idx.get((mt["account"], mt["message_id"]))
        out.append(
            {
                "timestamp": mt["timestamp"],
                "is_read": mt["is_read"],
                "conversation": titles.get((mt["account"], mt["conversation_id"]), "")
                or mt["conversation_id"],
                "sender": m.sender if m else "",
                "content": (m.content[:_MAXLEN] if m else ""),
                "account": mt["account"],
            }
        )
    out.sort(key=lambda d: _epoch_ms(d["timestamp"]) or 0, reverse=True)
    return out[:limit]


@mcp.tool()
def mentions(
    days: Optional[int] = None, unread_only: bool = False, account: Optional[str] = None, limit: int = 100
) -> list[dict]:
    """@-mentions of you (someone @mentioned you), newest first, with content.

    ``unread_only`` keeps only not-yet-read mentions; ``days`` restricts recency;
    ``account`` scopes to one tenant (see ``list_accounts``). This is the only
    reliable read/unread signal in the local store.
    """
    return _mentions_impl(days, unread_only, account, limit)


@mcp.tool()
def unread_messages(days: int = 30, account: Optional[str] = None, limit: int = 100) -> list[dict]:
    """Unread items needing attention = your UNREAD @-mentions.

    ⚠️ The local Teams cache has NO general read marker (per-message read state is
    not stored on disk), so true 'all unread' can't be derived. Unread @-mentions
    are the reliable signal. For 'everything that arrived while I was away', use
    ``recent_messages(days=...)`` instead (you read nothing while away → recent ≈ unread).
    """
    return _mentions_impl(days, True, account, limit)


@mcp.tool()
def overview(days: int = 7, account: Optional[str] = None) -> list[dict]:
    """Who messaged you: recent activity grouped by conversation, newest first.

    Per conversation: title, message count, distinct senders, last message + time.
    Great first call to get the lay of the land before drilling in.
    """
    snap = _snapshot()
    titles = _title_map(snap)
    cutoff = (time.time() - days * 86400) * 1000
    convs: dict = {}
    for m in snap["messages"]:
        if account and m.account != account:
            continue
        e = _epoch_ms(m.timestamp)
        if e is None or e < cutoff:
            continue
        g = convs.setdefault(
            (m.account, m.conversation_id), {"count": 0, "senders": set(), "last_e": 0.0, "last": None}
        )
        g["count"] += 1
        if m.sender:
            g["senders"].add(m.sender)
        if e > g["last_e"]:
            g["last_e"], g["last"] = e, m
    rows = []
    for (acc, cid), g in convs.items():
        last = g["last"]
        rows.append(
            {
                "conversation": titles.get((acc, cid), "") or cid,
                "count": g["count"],
                "senders": sorted(g["senders"])[:8],
                "last_sender": last.sender if last else "",
                "last_message": (last.content[:200] if last else ""),
                "last_time": last.timestamp if last else "",
                "account": acc,
            }
        )
    rows.sort(key=lambda r: _epoch_ms(r["last_time"]) or 0, reverse=True)
    return rows


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
