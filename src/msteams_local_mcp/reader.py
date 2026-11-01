"""Read the *new* Microsoft Teams (v2) message cache from the local disk.

The new Teams client (``com.microsoft.teams2`` on macOS, ``MSTeams`` MSIX on
Windows) is an Edge WebView2 app that keeps recent conversations in an IndexedDB
database backed by Chromium LevelDB, with values serialized in the V8 format.
This module reads that store **locally and read-only** — no Microsoft Graph, no
OAuth, no network. It only exposes data the signed-in user already has on disk.

Schema (react-web-client, observed on Teams 2.x, 2026):
  - one IndexedDB database per (tenant, user) context, named
    ``Teams:<manager>:react-web-client:<tenantId>:<userObjectId>:<locale>``
  - messages live in ``Teams:replychain-manager:…`` → object store
    ``replychains`` → each record has a ``messageMap`` {messageId: message}
  - chat/channel metadata lives in ``Teams:conversation-manager:…`` →
    object store ``conversations``

Low-level LevelDB + V8 parsing is delegated to ``ccl_chromium_reader`` (MIT), the
reference forensic library. This module maps the *current* Teams schema on top of
it (the older ``forensicsim`` mapping targets a previous schema and returns
nothing on 2.x). Nothing here is specific to any tenant or account.
"""
from __future__ import annotations

import dataclasses
import glob
import html
import os
import pathlib
import re
import shutil
import sys
import tempfile
from collections.abc import Iterator
from typing import Any, Optional

from ccl_chromium_reader import ccl_chromium_indexeddb as _idb

_TAG_RE = re.compile(r"<[^>]+>")


def _text(value: Any) -> str:
    """Best-effort plain text from an HTML/str/bytes message body."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", "replace")
    if not isinstance(value, str):
        value = str(value)
    return _TAG_RE.sub("", html.unescape(value)).strip()


def _conversation_title(record: dict) -> str:
    """Display title of a ``conversations`` record, or "" if it has none.

    Channels and named group chats have ``threadProperties.topic``. 1:1 and
    unnamed group chats have no topic, so use ``chatTitle``, the title computed
    by the Teams client. A title with an unresolved member is left out.
    """
    props = record.get("threadProperties")
    topic = _text((props.get("topic") if isinstance(props, dict) else "") or "")
    if topic:
        return topic
    chat = record.get("chatTitle")
    if not isinstance(chat, dict) or chat.get("hasUnknownUserInShortTitle"):
        return ""
    return _text(chat.get("shortTitle") or "") or _text(chat.get("longTitle") or "")


# --- Cache discovery ---------------------------------------------------------

def default_cache_globs() -> list[str]:
    """Candidate globs for the Teams v2 IndexedDB LevelDB, per OS."""
    home = pathlib.Path.home()
    if sys.platform == "darwin":
        base = (
            home
            / "Library/Containers/com.microsoft.teams2/Data/Library/"
            "Application Support/Microsoft/MSTeams/EBWebView"
        )
        return [str(base / "*/IndexedDB/https_teams.microsoft.com_0.indexeddb.leveldb")]
    if sys.platform.startswith("win"):
        local = os.environ.get("LOCALAPPDATA", str(home / "AppData/Local"))
        return [
            os.path.join(
                local,
                "Packages",
                "MSTeams_*",
                "LocalCache/Microsoft/MSTeams/EBWebView",
                "*/IndexedDB/https_teams.microsoft.com_0.indexeddb.leveldb",
            )
        ]
    # Linux (unofficial clients) — best effort.
    return [
        str(
            home
            / ".config"
            / "*eams*"
            / "*/IndexedDB/https_teams.microsoft.com_0.indexeddb.leveldb"
        )
    ]


def find_cache() -> Optional[str]:
    """Return the first Teams v2 LevelDB path found, or None."""
    for pattern in default_cache_globs():
        hits = sorted(glob.glob(pattern))
        if hits:
            return hits[0]
    return None


# --- Model -------------------------------------------------------------------

@dataclasses.dataclass
class Account:
    """One (tenant, user) context found in the cache."""

    tenant_id: str
    user_id: str
    # Human label inferred from the data (org/display names), not hardcoded.
    label: str = ""

    @property
    def key(self) -> str:
        return f"{self.tenant_id}:{self.user_id}"


@dataclasses.dataclass
class Message:
    account: str
    conversation_id: str
    reply_chain_id: str
    message_id: str
    sender: str
    timestamp: str
    content: str
    content_type: str


# --- Reader ------------------------------------------------------------------

class TeamsCacheReader:
    """Read messages/conversations from a Teams v2 IndexedDB LevelDB.

    The database is copied to a temporary directory before reading, because the
    running Teams client holds a lock on the live files.
    """

    def __init__(self, leveldb_path: Optional[str] = None, *, copy: bool = True):
        path = leveldb_path or find_cache()
        if not path or not os.path.isdir(path):
            raise FileNotFoundError(
                "Teams v2 cache not found. Pass a leveldb path explicitly, or make "
                "sure the new Teams client has been signed in at least once. Looked "
                f"in: {default_cache_globs()}"
            )
        self._src = path
        self._tmp: Optional[str] = None
        if copy:
            self._tmp = tempfile.mkdtemp(prefix="msteams-cache-")
            dst = os.path.join(self._tmp, os.path.basename(path))
            shutil.copytree(path, dst)
            path = dst
        self._db = _idb.WrappedIndexDB(path)
        self._skipped = 0

    def close(self) -> None:
        if self._tmp and os.path.isdir(self._tmp):
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None

    def __enter__(self) -> "TeamsCacheReader":
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    @property
    def skipped(self) -> int:
        """Number of records skipped because they could not be deserialized."""
        return self._skipped

    def _on_bad(self, _key, _raw):  # ccl bad_deserializer_data_handler
        self._skipped += 1
        return None

    def _stores(self, manager: str, store: str):
        """Yield (Account, WrappedObjectStore) for each context of a manager."""
        for meta in self._db.database_ids:
            name = meta.name or ""
            if f"Teams:{manager}:" not in name:
                continue
            parts = name.split(":")
            # …:react-web-client:<tenantId>:<userObjectId>:<locale>
            try:
                tenant_id, user_id = parts[-3], parts[-2]
            except IndexError:
                continue
            db = self._db[meta.dbid_no]
            try:
                obj = db.get_object_store_by_name(store)
            except Exception:
                continue
            yield Account(tenant_id=tenant_id, user_id=user_id), obj

    # -- public API --

    def accounts(self) -> list[Account]:
        """List the (tenant, user) contexts that hold messages, labeled by the
        most frequent org/display name seen in their messages."""
        out: list[Account] = []
        for acc, obj in self._stores("replychain-manager", "replychains"):
            names: dict[str, int] = {}
            seen = 0
            for rec in obj.iterate_records(bad_deserializer_data_handler=self._on_bad):
                if rec.value is None:
                    continue
                mm = rec.value.get("messageMap") or {}
                if isinstance(mm, dict):
                    for msg in mm.values():
                        who = _text(msg.get("imDisplayName"))
                        m = re.search(r"\(([^)]+)\)\s*$", who)  # e.g. "X (GP Rubix)"
                        if m:
                            names[m.group(1)] = names.get(m.group(1), 0) + 1
                seen += 1
                if seen >= 200:  # enough to infer a label cheaply
                    break
            acc.label = max(names, key=names.get) if names else ""
            out.append(acc)
        return out

    def messages(self, account: Optional[str] = None) -> Iterator[Message]:
        """Yield all messages, optionally restricted to one account key
        (``<tenantId>:<userObjectId>``)."""
        for acc, obj in self._stores("replychain-manager", "replychains"):
            if account and acc.key != account:
                continue
            for rec in obj.iterate_records(bad_deserializer_data_handler=self._on_bad):
                if rec.value is None:
                    continue
                mm = rec.value.get("messageMap") or {}
                if not isinstance(mm, dict):
                    continue
                for mid, msg in mm.items():
                    content = _text(msg.get("content"))
                    if not content:
                        continue
                    yield Message(
                        account=acc.key,
                        conversation_id=str(rec.value.get("conversationId", "")),
                        reply_chain_id=str(rec.value.get("replyChainId", "")),
                        message_id=str(msg.get("id", mid)),
                        sender=_text(msg.get("imDisplayName") or msg.get("from")),
                        timestamp=str(
                            msg.get("originalArrivalTime")
                            or msg.get("composetime")
                            or ""
                        ),
                        content=content,
                        content_type=str(msg.get("messageType") or ""),
                    )

    def conversations(self, account: Optional[str] = None) -> Iterator[dict]:
        """Yield chat/channel metadata records (title, members, id, last time…)."""
        for acc, obj in self._stores("conversation-manager", "conversations"):
            if account and acc.key != account:
                continue
            for rec in obj.iterate_records(bad_deserializer_data_handler=self._on_bad):
                if rec.value is None:
                    continue
                v = rec.value
                last = v.get("lastMessage") if isinstance(v.get("lastMessage"), dict) else {}
                yield {
                    "account": acc.key,
                    "id": str(v.get("id", "")),
                    "title": _conversation_title(v),
                    "type": str(v.get("type", "")),
                    "last_message_time": str(v.get("lastMessageTimeUtc", "")),
                    "last_message": _text(last.get("content") if last else ""),
                }

    def mentions(self, account: Optional[str] = None) -> Iterator[dict]:
        """Yield @-mention entries (read state included). The only reliable read
        marker in the cache: general read/unread is NOT stored locally, but each
        mention carries ``is_read``. Content is not here — join on
        (conversation_id, message_id) with :meth:`messages`."""
        for acc, obj in self._stores("messaging-slice-manager", "mentions-metadata-items"):
            if account and acc.key != account:
                continue
            for rec in obj.iterate_records(bad_deserializer_data_handler=self._on_bad):
                if rec.value is None or not isinstance(rec.value, dict):
                    continue
                v = rec.value
                yield {
                    "account": acc.key,
                    "conversation_id": str(v.get("sourceThreadId", "")),
                    "message_id": str(v.get("sourceMessageId", "")),
                    "reply_chain_id": str(v.get("sourceReplyChainId", "")),
                    "timestamp": str(v.get("timestamp", "")),
                    "is_read": str(v.get("isRead")) == "True",
                }
