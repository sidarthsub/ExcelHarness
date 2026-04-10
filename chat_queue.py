"""Thread-safe FIFO for inbound chat messages.

The add-in sends chat messages over WSS; the WSS handler calls enqueue().
The harness's Builder loop calls drain_all() between turns.
"""
import threading


class ChatQueue:
    def __init__(self):
        self._messages: list[str] = []
        self._lock = threading.Lock()

    def enqueue(self, text: str) -> None:
        with self._lock:
            self._messages.append(text)

    def drain_all(self) -> list[str]:
        with self._lock:
            drained = self._messages[:]
            self._messages.clear()
            return drained
