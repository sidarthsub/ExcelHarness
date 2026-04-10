"""Tests for the in-memory chat FIFO used by the harness."""
import pytest

from chat_queue import ChatQueue


def test_enqueue_and_drain():
    q = ChatQueue()
    q.enqueue("hello")
    q.enqueue("world")
    drained = q.drain_all()
    assert drained == ["hello", "world"]


def test_drain_when_empty():
    q = ChatQueue()
    assert q.drain_all() == []


def test_drain_clears_queue():
    q = ChatQueue()
    q.enqueue("a")
    q.drain_all()
    assert q.drain_all() == []


def test_thread_safe_concurrent_writes():
    """enqueue and drain should not lose messages under concurrent writes."""
    import threading
    q = ChatQueue()
    N = 1000

    def writer():
        for i in range(N):
            q.enqueue(f"msg-{i}")

    threads = [threading.Thread(target=writer) for _ in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    all_msgs = q.drain_all()
    assert len(all_msgs) == N * 4
