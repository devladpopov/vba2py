"""VBA type emulation: Empty, Null, Collection, Array."""

from __future__ import annotations

from typing import Any, Dict, Iterator, List, Optional, Union


# --- Sentinel types ---


class _VBAEmpty:
    """Sentinel representing VBA Empty (uninitialized Variant)."""

    _instance: Optional["_VBAEmpty"] = None

    def __new__(cls) -> "_VBAEmpty":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "Empty"

    def __str__(self) -> str:
        return ""

    def __bool__(self) -> bool:
        return False

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _VBAEmpty):
            return True
        # Empty == 0, Empty == ""
        if other == 0 or other == "":
            return True
        return False

    def __hash__(self) -> int:
        return hash(None)


class _VBANull:
    """Sentinel representing VBA Null.

    Null propagates through expressions: Null + anything = Null.
    """

    _instance: Optional["_VBANull"] = None

    def __new__(cls) -> "_VBANull":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "Null"

    def __str__(self) -> str:
        return "Null"

    def __bool__(self) -> bool:
        raise TypeError("Invalid use of Null")

    # Null propagation: any operation with Null returns Null
    def __add__(self, other: Any) -> "_VBANull":
        return self

    def __radd__(self, other: Any) -> "_VBANull":
        return self

    def __sub__(self, other: Any) -> "_VBANull":
        return self

    def __rsub__(self, other: Any) -> "_VBANull":
        return self

    def __mul__(self, other: Any) -> "_VBANull":
        return self

    def __rmul__(self, other: Any) -> "_VBANull":
        return self

    def __truediv__(self, other: Any) -> "_VBANull":
        return self

    def __rtruediv__(self, other: Any) -> "_VBANull":
        return self

    def __eq__(self, other: object) -> bool:
        # Null = Null is Null in VBA (not True), but for Python usability
        # we allow identity check via `is`
        return isinstance(other, _VBANull)

    def __hash__(self) -> int:
        return hash("VBANull")


EMPTY = _VBAEmpty()
NULL = _VBANull()


# --- VBACollection ---


class VBACollection:
    """1-based, ordered collection with optional string keys.

    Emulates VBA Collection object.
    """

    def __init__(self) -> None:
        self._items: List[Any] = []
        self._keys: Dict[str, int] = {}  # key -> index in _items (0-based internal)

    def add(
        self,
        item: Any,
        key: Optional[str] = None,
        before: Optional[Union[int, str]] = None,
        after: Optional[Union[int, str]] = None,
    ) -> None:
        """Add an item to the collection.

        Args:
            item: The item to add.
            key: Optional unique string key.
            before: Insert before this index (1-based) or key.
            after: Insert after this index (1-based) or key.
        """
        if key is not None and key in self._keys:
            raise ValueError(f"Key '{key}' already exists in collection")

        if before is not None and after is not None:
            raise ValueError("Cannot specify both before and after")

        insert_idx: int
        if before is not None:
            insert_idx = self._resolve_index(before)
        elif after is not None:
            insert_idx = self._resolve_index(after) + 1
        else:
            insert_idx = len(self._items)

        self._items.insert(insert_idx, item)

        # Update key mappings for shifted items
        for k, v in self._keys.items():
            if v >= insert_idx:
                self._keys[k] = v + 1

        if key is not None:
            self._keys[key] = insert_idx

    def remove(self, index: Union[int, str]) -> None:
        """Remove item by 1-based index or string key."""
        internal_idx = self._resolve_index(index)

        # Remove key if it exists for this index
        key_to_remove: Optional[str] = None
        for k, v in self._keys.items():
            if v == internal_idx:
                key_to_remove = k
                break
        if key_to_remove is not None:
            del self._keys[key_to_remove]

        self._items.pop(internal_idx)

        # Update remaining key mappings
        for k, v in list(self._keys.items()):
            if v > internal_idx:
                self._keys[k] = v - 1

    def item(self, index: Union[int, str]) -> Any:
        """Get item by 1-based index or string key."""
        internal_idx = self._resolve_index(index)
        return self._items[internal_idx]

    @property
    def count(self) -> int:
        """Number of items in the collection."""
        return len(self._items)

    def _resolve_index(self, index: Union[int, str]) -> int:
        """Convert 1-based int or string key to 0-based internal index."""
        if isinstance(index, str):
            if index not in self._keys:
                raise KeyError(f"Key '{index}' not found in collection")
            return self._keys[index]
        # 1-based integer index
        if index < 1 or index > len(self._items):
            raise IndexError(f"Subscript out of range: {index}")
        return index - 1

    def __iter__(self) -> Iterator[Any]:
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __getitem__(self, index: Union[int, str]) -> Any:
        return self.item(index)


# --- VBAArray ---


class VBAArray:
    """1-based (by default) array wrapper emulating VBA arrays.

    Usage:
        arr = VBAArray([10, 20, 30])  # indices 1..3
        arr[1]  # 10
        arr[3]  # 30
    """

    def __init__(self, data: Optional[List[Any]] = None, base: int = 1) -> None:
        self._data: List[Any] = list(data) if data else []
        self._base: int = base

    def __getitem__(self, index: int) -> Any:
        internal = index - self._base
        if internal < 0 or internal >= len(self._data):
            raise IndexError(f"Subscript out of range: {index}")
        return self._data[internal]

    def __setitem__(self, index: int, value: Any) -> None:
        internal = index - self._base
        if internal < 0 or internal >= len(self._data):
            raise IndexError(f"Subscript out of range: {index}")
        self._data[internal] = value

    def __len__(self) -> int:
        return len(self._data)

    def __iter__(self) -> Iterator[Any]:
        return iter(self._data)

    def __repr__(self) -> str:
        return f"VBAArray({self._data}, base={self._base})"

    def redim(self, new_size: int, preserve: bool = False) -> None:
        """ReDim the array. If preserve=True, keep existing data up to new_size."""
        if new_size < 0:
            raise ValueError("Invalid procedure call or argument")
        if preserve:
            if new_size >= len(self._data):
                self._data.extend([EMPTY] * (new_size - len(self._data)))
            else:
                self._data = self._data[:new_size]
        else:
            self._data = [EMPTY] * new_size

    @property
    def ubound(self) -> int:
        """Upper bound (1-based by default)."""
        if not self._data:
            raise ValueError("Subscript out of range: array is empty")
        return self._base + len(self._data) - 1

    @property
    def lbound(self) -> int:
        """Lower bound."""
        return self._base
