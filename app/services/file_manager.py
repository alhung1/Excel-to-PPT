"""
File management service — upload tracking and cleanup.
"""
import os
import time
import threading
from pathlib import Path
from typing import Dict, Optional

from app.config import logger, UPLOAD_DIR, OUTPUT_DIR, FILE_CLEANUP_MAX_AGE


# ---------------------------------------------------------------------------
# In-memory file registry
# ---------------------------------------------------------------------------
class FileManager:
    """Thread-safe registry for uploaded files with auto-cleanup."""

    def __init__(self):
        self._files: Dict[str, dict] = {}
        self._lock = threading.Lock()

    # -- CRUD ---------------------------------------------------------------

    def register(self, file_id: str, file_type: str, path: str, filename: str):
        with self._lock:
            self._files[file_id] = {
                "type": file_type,
                "path": path,
                "filename": filename,
                "created_at": time.time(),
            }

    def get(self, file_id: str) -> Optional[dict]:
        with self._lock:
            return self._files.get(file_id)

    def remove(self, file_id: str) -> bool:
        with self._lock:
            info = self._files.pop(file_id, None)
        if info:
            p = Path(info["path"])
            if p.exists():
                p.unlink()
                logger.info("Removed file: %s", p)
            return True
        return False

    @property
    def count(self) -> int:
        with self._lock:
            return len(self._files)

    # -- Cleanup ------------------------------------------------------------

    def cleanup_old_files(self, max_age: float = None):
        """Delete uploaded & output files older than *max_age* seconds."""
        if max_age is None:
            max_age = FILE_CLEANUP_MAX_AGE

        now = time.time()
        removed = 0

        # Clean tracked uploads
        with self._lock:
            expired = [
                fid
                for fid, info in self._files.items()
                if now - info.get("created_at", now) > max_age
            ]

        for fid in expired:
            self.remove(fid)
            removed += 1

        # Clean untracked files in upload/output dirs
        for directory in [UPLOAD_DIR, OUTPUT_DIR]:
            for item in directory.iterdir():
                try:
                    age = now - item.stat().st_mtime
                    if age > max_age:
                        if item.is_dir():
                            import shutil
                            shutil.rmtree(item)
                        else:
                            item.unlink()
                        removed += 1
                except Exception as e:
                    logger.warning("Cleanup error for %s: %s", item, e)

        if removed:
            logger.info("Cleanup: removed %d old files/directories", removed)


# Singleton instance
file_manager = FileManager()


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def get_directory_size_mb(path: Path) -> float:
    """Return total size of a directory in megabytes."""
    total = 0
    try:
        for f in path.rglob("*"):
            if f.is_file():
                total += f.stat().st_size
    except Exception:
        pass
    return total / (1024 * 1024)
