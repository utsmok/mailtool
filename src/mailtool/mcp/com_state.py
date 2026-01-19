"""COM State Management for MCP Server

This module provides thread-safe COM initialization tracking for the MCP server.
COM must be initialized once per thread that accesses COM objects.

This module is shared between server.py and resources.py to ensure consistent
COM state across all MCP tools and resources.
"""

import logging
import threading

import pythoncom

# Configure logging
logger = logging.getLogger(__name__)

# Track which threads have initialized COM
# COM must be initialized once per thread that accesses COM objects
_com_initialized_threads: set[int] = set()
_com_lock = threading.Lock()


def ensure_com_initialized() -> None:
    """Ensure COM is initialized for the current thread.

    This function is called by every MCP tool and resource to ensure COM is available
    in the calling thread. It's safe to call CoInitialize multiple times
    in the same thread - each call must be matched with CoUninitialize.

    Note: We track initialized threads to avoid double-initialization.
    The same thread can call this multiple times, but COM is only initialized once.

    Thread Safety:
        This function is thread-safe. Multiple threads can call it concurrently,
        and each thread will initialize COM exactly once.
    """
    thread_id = threading.get_ident()

    with _com_lock:
        if thread_id not in _com_initialized_threads:
            logger.debug(f"Initializing COM for thread {thread_id}")
            pythoncom.CoInitialize()
            _com_initialized_threads.add(thread_id)
            logger.debug(f"COM initialized for thread {thread_id}")


def get_initialized_thread_count() -> int:
    """Get the number of threads that have initialized COM.

    Returns:
        int: Number of threads with COM initialized
    """
    with _com_lock:
        return len(_com_initialized_threads)


def is_com_initialized_for_thread(thread_id: int | None = None) -> bool:
    """Check if COM is initialized for a specific thread.

    Args:
        thread_id: Thread ID to check (defaults to current thread)

    Returns:
        bool: True if COM is initialized for the thread
    """
    if thread_id is None:
        thread_id = threading.get_ident()

    with _com_lock:
        return thread_id in _com_initialized_threads
