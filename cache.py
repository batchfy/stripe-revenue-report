import time
from typing import Any

from sqlitedict import SqliteDict


class StripeCache:
    """SQLite-backed key-value store for Stripe API objects."""

    def __init__(self, path: str = "stripe_cache.sqlite"):
        self._db = SqliteDict(path, autocommit=True)

    def __enter__(self) -> "StripeCache":
        return self

    def __exit__(self, *_) -> None:
        self._db.close()

    def get(self, key: str) -> Any | None:
        record = self._db.get(key)
        return record["data"] if record is not None else None

    def set(self, key: str, value: Any) -> None:
        self._db[key] = {"ts": time.time(), "data": value}

    def delete(self, key: str) -> bool:
        if key in self._db:
            del self._db[key]
            return True
        return False
