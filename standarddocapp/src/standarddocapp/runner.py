"""Background job runner that streams log records into a queue.

Runs an arbitrary callable on a background daemon thread while sending all
log records (from the targeted logger names) into a queue.Queue so the Tk
main loop can drain them via after() without blocking the UI.
"""
from __future__ import annotations

import logging
import queue
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional


@dataclass
class JobMessage:
    kind: str  # "log" | "done" | "error"
    text: str = ""
    return_code: Optional[int] = None


class _QueueLogHandler(logging.Handler):
    """logging.Handler that pushes formatted records onto a queue.Queue."""

    def __init__(self, q: "queue.Queue[JobMessage]") -> None:
        super().__init__()
        self.q = q
        self.setFormatter(
            logging.Formatter(
                fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
                datefmt="%H:%M:%S",
            )
        )

    def emit(self, record: logging.LogRecord) -> None:  # noqa: D401
        try:
            self.q.put_nowait(JobMessage(kind="log", text=self.format(record)))
        except Exception:
            # Last-resort: never let logging crash the worker.
            pass


class JobRunner:
    """Run one background job at a time.

    Parameters
    ----------
    logger_names:
        Names of loggers whose records should be captured into the queue.
        Pass ``("stdharvest",)`` for the harvest tab and ``("stdsearch",)``
        for the search tab.
    """

    def __init__(self, logger_names: Iterable[str]) -> None:
        self.logger_names = tuple(logger_names)
        self.queue: "queue.Queue[JobMessage]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._handler: Optional[_QueueLogHandler] = None
        self._lock = threading.Lock()
        self._log_file: Optional[Path] = None
        self._file_handler: Optional[logging.FileHandler] = None

    @property
    def is_running(self) -> bool:
        t = self._thread
        return t is not None and t.is_alive()

    def start(
        self,
        target: Callable[[], int],
        log_file_dir: Optional[Path] = None,
        log_file_prefix: str = "job",
    ) -> None:
        """Run ``target`` on a background thread; ``target`` returns an int rc."""
        if self.is_running:
            raise RuntimeError("A job is already running.")

        with self._lock:
            self._handler = _QueueLogHandler(self.queue)
            self._handler.setLevel(logging.INFO)
            for name in self.logger_names:
                lg = logging.getLogger(name)
                lg.setLevel(logging.INFO)
                lg.addHandler(self._handler)

            if log_file_dir is not None:
                log_file_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                self._log_file = log_file_dir / f"{log_file_prefix}_{stamp}.log"
                self._file_handler = logging.FileHandler(
                    self._log_file, encoding="utf-8"
                )
                self._file_handler.setLevel(logging.INFO)
                self._file_handler.setFormatter(
                    logging.Formatter(
                        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S",
                    )
                )
                for name in self.logger_names:
                    logging.getLogger(name).addHandler(self._file_handler)
                self.queue.put_nowait(
                    JobMessage(kind="log", text=f"Log file: {self._log_file}")
                )

        def _runner() -> None:
            rc = 1
            try:
                rc = int(target() or 0)
                self.queue.put_nowait(
                    JobMessage(kind="done", text=f"Job finished (rc={rc})", return_code=rc)
                )
            except SystemExit as exc:
                rc = int(exc.code) if isinstance(exc.code, int) else 1
                self.queue.put_nowait(
                    JobMessage(kind="done", text=f"Job exited (rc={rc})", return_code=rc)
                )
            except Exception as exc:  # noqa: BLE001
                tb = traceback.format_exc()
                self.queue.put_nowait(JobMessage(kind="log", text=tb))
                self.queue.put_nowait(
                    JobMessage(
                        kind="error",
                        text=f"Job failed: {exc}",
                        return_code=2,
                    )
                )
            finally:
                self._detach_handlers()

        self._thread = threading.Thread(target=_runner, daemon=True)
        self._thread.start()

    def _detach_handlers(self) -> None:
        with self._lock:
            if self._handler is not None:
                for name in self.logger_names:
                    try:
                        logging.getLogger(name).removeHandler(self._handler)
                    except Exception:
                        pass
                self._handler = None
            if self._file_handler is not None:
                for name in self.logger_names:
                    try:
                        logging.getLogger(name).removeHandler(self._file_handler)
                    except Exception:
                        pass
                try:
                    self._file_handler.close()
                except Exception:
                    pass
                self._file_handler = None

    @property
    def log_file(self) -> Optional[Path]:
        return self._log_file

    def drain(self, max_items: int = 200) -> list[JobMessage]:
        """Pop up to ``max_items`` messages from the queue (non-blocking)."""
        out: list[JobMessage] = []
        for _ in range(max_items):
            try:
                out.append(self.queue.get_nowait())
            except queue.Empty:
                break
        return out
