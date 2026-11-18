"""Object store lookup, with synthetic databases."""
from types import SimpleNamespace

from msteams_local_mcp.reader import Account, TeamsCacheReader


def _database(*names):
    return SimpleNamespace(object_store_names=names, get_object_store_by_id=lambda number: f"store {number}")


class _IndexedDb:
    """Stands in for ccl's WrappedIndexDB: database_ids and lookup by number."""

    def __init__(self, databases):
        self.database_ids = [SimpleNamespace(dbid_no=n, name=name) for n, (name, _db) in enumerate(databases)]
        self._databases = [db for _name, db in databases]

    def __getitem__(self, number):
        return self._databases[number]


def _reader(databases):
    reader = TeamsCacheReader.__new__(TeamsCacheReader)
    reader._db = _IndexedDb(databases)
    return reader


def test_numbered_copies_of_a_store_are_read():
    reader = _reader([
        ("Teams:replychain-manager:react-web-client:t:u:en-us",
         _database("replychains", "replychains-2", "replychains-old", None)),
        ("Teams:streams-replychain-manager:react-web-client:t:u:en-us", _database("replychains")),
    ])

    assert list(reader._stores("replychain-manager", "replychains")) == [
        (Account(tenant_id="t", user_id="u"), "store 1"),
        (Account(tenant_id="t", user_id="u"), "store 2"),
    ]


def test_database_without_the_store_yields_nothing():
    reader = _reader([("Teams:replychain-manager:react-web-client:t:u:en-us", _database("other"))])

    assert list(reader._stores("replychain-manager", "replychains")) == []
