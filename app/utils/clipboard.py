"""
Windows clipboard utilities for COM automation operations.
"""
from app.config import logger


def clear_clipboard():
    """Clear the Windows clipboard to avoid stale data issues."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except ImportError:
        logger.debug("win32clipboard not available (non-Windows platform)")
    except Exception as e:
        logger.debug("Failed to clear clipboard: %s", e)
