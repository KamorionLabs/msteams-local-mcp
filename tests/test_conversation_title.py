"""Titles of conversation records, with synthetic data."""
from types import SimpleNamespace

import pytest

from msteams_local_mcp.reader import Account, TeamsCacheReader, _conversation_title


class Undefined:
    """Behaves like the V8 undefined value of ccl_chromium_reader: falsy."""

    def __bool__(self):
        return False

    def __repr__(self):
        return "<Undefined>"


@pytest.mark.parametrize(
    "record, expected",
    [
        ({"threadProperties": {"topic": "Team chat"}, "chatTitle": {"shortTitle": "Ann"}}, "Team chat"),
        ({"threadProperties": {"topic": None}, "chatTitle": {"shortTitle": "Ann Lee"}}, "Ann Lee"),
        ({"threadProperties": {"topic": Undefined()}, "chatTitle": {"shortTitle": "Ann Lee"}}, "Ann Lee"),
        ({"chatTitle": {"shortTitle": "", "longTitle": "Ann Lee, Bob Ray"}}, "Ann Lee, Bob Ray"),
        ({"chatTitle": {"shortTitle": Undefined(), "longTitle": Undefined()}}, ""),
        ({"chatTitle": {"shortTitle": "Unknown", "hasUnknownUserInShortTitle": True}}, ""),
        ({"chatTitle": {"shortTitle": "Unknown", "longTitle": "Ann Lee, Bob Ray",
                        "hasUnknownUserInShortTitle": True}}, "Ann Lee, Bob Ray"),
        ({"chatTitle": Undefined()}, ""),
        ({"threadProperties": Undefined()}, ""),
        ({}, ""),
    ],
)
def test_conversation_title(record, expected):
    assert _conversation_title(record) == expected


def test_conversations_yields_the_chat_title():
    records = [
        SimpleNamespace(value={"id": "19:a_b@unq.gbl.spaces", "threadProperties": {"topic": None},
                               "chatTitle": {"shortTitle": "Ann Lee"}}),
        SimpleNamespace(value=None),
    ]
    reader = TeamsCacheReader.__new__(TeamsCacheReader)
    reader._stores = lambda manager, store: [
        (Account(tenant_id="t", user_id="u"), SimpleNamespace(iterate_records=lambda **_: records))
    ]

    rows = list(reader.conversations())

    assert [(row["id"], row["title"]) for row in rows] == [("19:a_b@unq.gbl.spaces", "Ann Lee")]


@pytest.mark.parametrize("root", [["item"], Undefined()])
@pytest.mark.parametrize(
    "read",
    [
        lambda reader: list(reader.conversations()),
        lambda reader: list(reader.messages()),
        lambda reader: [acc.label for acc in reader.accounts() if acc.label],
    ],
    ids=["conversations", "messages", "accounts"],
)
def test_records_that_are_not_objects_are_skipped(read, root):
    reader = TeamsCacheReader.__new__(TeamsCacheReader)
    reader._stores = lambda manager, store: [
        (Account(tenant_id="t", user_id="u"),
         SimpleNamespace(iterate_records=lambda **_: [SimpleNamespace(value=root)]))
    ]

    assert read(reader) == []
