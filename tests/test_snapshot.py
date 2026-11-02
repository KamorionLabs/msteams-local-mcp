"""On-disk snapshot versioning, with synthetic data."""
import json
import time

import pytest

from msteams_local_mcp import server
from msteams_local_mcp.reader import Message


class FakeReader:
    """Stands in for TeamsCacheReader and counts the full parses."""

    parses = 0
    skipped = 0

    def __init__(self, _leveldb):
        type(self).parses += 1

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        pass

    def messages(self):
        return [Message("a", "c", "r", "m", "s", "1", "hi", "text")]

    def conversations(self):
        return [{"account": "a", "id": "c", "title": "Ann Lee"}]

    def mentions(self):
        return []


@pytest.fixture
def leveldb(tmp_path, monkeypatch):
    path = tmp_path / "leveldb"
    path.mkdir()
    (path / "000001.log").write_text("x")
    monkeypatch.setattr(server, "_LEVELDB", str(path))
    monkeypatch.setattr(server, "_CACHE_FILE", tmp_path / "snapshot.json")
    monkeypatch.setattr(server, "TeamsCacheReader", FakeReader)
    monkeypatch.setattr(FakeReader, "parses", 0)
    monkeypatch.setattr(server, "_snap", {
        "ts": 0.0, "sig": "", "messages": [], "conversations": [],
        "accounts": [], "mentions": [], "skipped": 0,
    })
    return str(path)


def _disk_snapshot(sig, **extra):
    return {
        "sig": sig, "messages": [], "conversations": [{"account": "a", "id": "c", "title": ""}],
        "accounts": [], "mentions": [], "skipped": 0, **extra,
    }


@pytest.mark.parametrize("extra", [{}, {"version": server._SNAPSHOT_VERSION - 1}])
def test_snapshot_of_another_version_is_parsed_again(leveldb, extra):
    server._CACHE_FILE.write_text(json.dumps(_disk_snapshot(server._signature(leveldb), **extra)))

    snap = server._snapshot()

    assert FakeReader.parses == 1
    assert snap["conversations"][0]["title"] == "Ann Lee"
    assert json.loads(server._CACHE_FILE.read_text())["version"] == server._SNAPSHOT_VERSION


def test_snapshot_of_the_current_version_is_reused(leveldb):
    server._CACHE_FILE.write_text(json.dumps(
        _disk_snapshot(server._signature(leveldb), version=server._SNAPSHOT_VERSION)
    ))

    snap = server._snapshot()

    assert FakeReader.parses == 0
    assert snap["conversations"][0]["title"] == ""


def test_snapshot_with_unexpected_message_fields_is_parsed_again(leveldb):
    server._CACHE_FILE.write_text(json.dumps({
        **_disk_snapshot(server._signature(leveldb), version=server._SNAPSHOT_VERSION),
        "messages": [{"unexpected": "field"}],
    }))

    snap = server._snapshot()

    assert FakeReader.parses == 1
    assert snap["conversations"][0]["title"] == "Ann Lee"


@pytest.mark.parametrize("tool", [server.overview, server.mentions], ids=["overview", "mentions"])
def test_titled_rows_lead_to_read_conversation(monkeypatch, tool):
    now = str(time.time() * 1000)
    snap = {
        "messages": [Message("a", "c", "r", "m", "Ann Lee", now, "hi", "text")],
        "conversations": [{"account": "a", "id": "c", "title": "Ann Lee"}],
        "mentions": [{"account": "a", "conversation_id": "c", "message_id": "m",
                      "reply_chain_id": "r", "timestamp": now, "is_read": False}],
    }
    monkeypatch.setattr(server, "_snapshot", lambda: snap)

    row = tool()[0]

    assert row["conversation"] == "Ann Lee"
    assert [m["content"] for m in server.read_conversation(row["conversation_id"])] == ["hi"]
