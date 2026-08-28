"""Safe background runner for the dashboard's one fixed scraper operation.

The manager deliberately exposes no command-selection API.  It runs the
server-configured ``main.py`` with the current Python interpreter, keeps one
job active at a time, validates the resulting workbooks, and restores the
previous spreadsheet/CSV bytes if anything fails.
"""

from __future__ import annotations

from collections import deque
import csv
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from typing import Any, Callable
import uuid

from dotenv import dotenv_values

try:  # The dashboard is currently run on macOS; keep a safe fallback for tests.
    import fcntl
except ImportError:  # pragma: no cover - Windows fallback
    fcntl = None


OUTPUT_PATTERNS = ("US_Event_*.xlsx", "US_Event_*.csv")
MAX_CAPTURED_LINES = 400
MAX_PUBLIC_DETAILS = 8_000

_SECRET_LABEL_RE = re.compile(
    r"(?im)^(\s*(?:authorization|password|passwd|passphrase|secret|"
    r"access[_ -]?token|refresh[_ -]?token|token|cookie|session(?:[_ -]?id)?|"
    r"auth(?:[_ -]?state)?|email|username|user)\s*[:=]\s*).*$"
)
_ENV_SECRET_LINE_RE = re.compile(
    r"(?im)^(\s*[A-Z][A-Z0-9_]*(?:EMAIL|USERNAME|USER|PASSWORD|PASS|TOKEN|"
    r"SECRET|COOKIE|SESSION|AUTH)[A-Z0-9_]*\s*=\s*).*$"
)
_BEARER_RE = re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._~+/=-]+")
_BASIC_AUTH_RE = re.compile(r"(?i)\bBasic\s+[A-Za-z0-9+/=]+")
_EMAIL_RE = re.compile(
    r"(?i)\b[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?"
    r"(?:\.[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?)+\b"
)
_QUERY_SECRET_RE = re.compile(
    r"(?i)([?&](?:password|passwd|token|secret|cookie|session|auth)[^=&#\s]*=)"
    r"[^&#\s]+"
)
_INLINE_SECRET_RE = re.compile(
    r"(?i)((?:authorization|password|passwd|passphrase|secret|access[_ -]?token|"
    r"refresh[_ -]?token|token|cookie|session(?:[_ -]?id)?|auth(?:[_ -]?state)?|"
    r"email|username|user)[^:=\n]{0,24}[:=]\s*)"
    r"(?:\"[^\"\n]*\"|'[^'\n]*'|[^\s,;}\]]+)"
)
_SECRET_ENV_NAME_RE = re.compile(
    r"(?i)(?:EMAIL|PASSWORD|PASS|TOKEN|SECRET|COOKIE|SESSION|AUTH)"
)


def _now_local() -> datetime:
    return datetime.now().astimezone()


def _iso(value: datetime | None) -> str:
    return value.isoformat(timespec="seconds") if value else ""


def _display_timestamp(value: datetime | None) -> str:
    if value is None:
        return ""
    hour = value.hour % 12 or 12
    period = "AM" if value.hour < 12 else "PM"
    return f"{value.strftime('%B')} {value.day}, {value.year} at {hour}:{value.minute:02d} {period}"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class RefreshManager:
    """Coordinate and report one local spreadsheet-refresh subprocess."""

    def __init__(
        self,
        *,
        project_dir: str | Path,
        update_script: str | Path,
        loader: Callable[..., dict[str, Any]],
        state_file: str | Path | None = None,
        lock_file: str | Path | None = None,
        popen_factory: Callable[..., subprocess.Popen[str]] = subprocess.Popen,
    ) -> None:
        self.project_dir = Path(project_dir).resolve()
        self.update_script = Path(update_script).resolve()
        self.loader = loader
        self.state_file = Path(
            state_file or (self.project_dir / "dashboard_state.json")
        ).resolve()
        self.lock_file = Path(
            lock_file or (self.project_dir / ".dashboard_refresh.lock")
        ).resolve()
        self._popen_factory = popen_factory
        self._state_lock = threading.RLock()
        self._thread: threading.Thread | None = None
        self._observed_external_refresh = False
        self._external_success_baseline = ""
        persisted = self._read_persisted_state()
        self._status: dict[str, Any] = {
            "status": "idle",
            "message": "Ready to refresh gig data.",
            "stage": "",
            "job_id": "",
            "started_at": "",
            "finished_at": "",
            "last_success_at": persisted.get("last_success_at", ""),
            "last_success_display": persisted.get("last_success_display", ""),
            "new_count": persisted.get("new_count"),
            "details": "",
        }

    def get_status(self) -> dict[str, Any]:
        """Return a detached, JSON-safe public status snapshot."""

        with self._state_lock:
            payload = dict(self._status)
            observed_before_probe = self._observed_external_refresh
        if payload.get("status") == "idle" or observed_before_probe:
            # Serialize a same-manager status probe with start_refresh so a
            # harmless GET cannot momentarily make a new POST look duplicated.
            with self._state_lock:
                another_process_holds_lock = self._another_process_holds_lock()
            if another_process_holds_lock:
                with self._state_lock:
                    if not self._observed_external_refresh:
                        self._external_success_baseline = str(
                            self._status.get("last_success_at") or ""
                        )
                    self._observed_external_refresh = True
                payload.update(
                    {
                        "status": "running",
                        "message": "Gig data is already being refreshed.",
                        "stage": "running",
                        "details": "",
                    }
                )
            else:
                # Read after observing the unlocked state so a just-finished
                # external manager's atomic success metadata cannot be missed.
                persisted = self._read_persisted_state()
                with self._state_lock:
                    observed_external = self._observed_external_refresh
                    baseline = self._external_success_baseline
                    self._observed_external_refresh = False
                    self._external_success_baseline = ""
                persisted_at = str(persisted.get("last_success_at") or "")
                if observed_external:
                    if persisted_at > baseline:
                        payload.update(
                            {
                                "status": "success",
                                "message": "Gig data refreshed successfully.",
                                "stage": "complete",
                                "finished_at": persisted_at,
                                "details": "",
                            }
                        )
                        payload.update(persisted)
                    else:
                        payload.update(
                            {
                                "status": "error",
                                "message": "Refresh failed. Your previous gig data is still available.",
                                "stage": "",
                                "details": "The external refresh ended without validated updated data.",
                            }
                        )
                    with self._state_lock:
                        self._status.update(payload)
                elif persisted_at > str(payload.get("last_success_at") or ""):
                    payload.update(persisted)
        return payload

    def start_refresh(self) -> tuple[bool, dict[str, Any]]:
        """Start the fixed updater in a daemon thread and return immediately."""

        with self._state_lock:
            if (
                self._status.get("status") == "running"
                and self._thread is not None
                and self._thread.is_alive()
            ):
                payload = dict(self._status)
                payload["message"] = "Gig data is already being refreshed."
                return False, payload

            lock_handle = self._acquire_process_lock()
            if lock_handle is None:
                if not self._observed_external_refresh:
                    self._external_success_baseline = str(
                        self._status.get("last_success_at") or ""
                    )
                self._observed_external_refresh = True
                payload = dict(self._status)
                payload.update(
                    {
                        "status": "running",
                        "message": "Gig data is already being refreshed.",
                        "stage": "running",
                        "details": "",
                    }
                )
                return False, payload

            started = _now_local()
            job_id = uuid.uuid4().hex
            self._status.update(
                {
                    "status": "running",
                    "message": "Starting the calendar refresh...",
                    "stage": "starting",
                    "job_id": job_id,
                    "started_at": _iso(started),
                    "finished_at": "",
                    "new_count": None,
                    "details": "",
                }
            )
            thread = threading.Thread(
                target=self._run_job,
                args=(job_id, lock_handle),
                name=f"us-event-refresh-{job_id[:8]}",
                daemon=True,
            )
            self._thread = thread
            try:
                thread.start()
            except Exception:
                self._release_process_lock(lock_handle)
                self._thread = None
                self._status.update(
                    {
                        "status": "error",
                        "message": "Refresh failed. Your previous gig data is still available.",
                        "stage": "",
                        "finished_at": _iso(_now_local()),
                        "details": "The background refresh could not be started.",
                    }
                )
                return False, dict(self._status)
            return True, dict(self._status)

    def _run_job(self, job_id: str, lock_handle: Any) -> None:
        rollback_dir: Path | None = None
        manifest: dict[str, str] = {}
        captured: deque[str] = deque(maxlen=MAX_CAPTURED_LINES)
        failure: Exception | None = None

        try:
            rollback_dir, manifest = self._snapshot_outputs()
            if not self.update_script.is_file():
                raise FileNotFoundError("The configured calendar update script was not found.")

            command = [sys.executable, str(self.update_script)]
            popen_options: dict[str, Any] = {
                "cwd": str(self.project_dir),
                "stdin": subprocess.DEVNULL,
                "stdout": subprocess.PIPE,
                "stderr": subprocess.STDOUT,
                "text": True,
                "encoding": "utf-8",
                "errors": "replace",
                "bufsize": 1,
                "shell": False,
            }
            # Retaining the flock descriptor in the child means the lock remains
            # held even if Flask is unexpectedly terminated while main.py runs.
            if os.name == "posix" and fcntl is not None:
                popen_options["pass_fds"] = (lock_handle.fileno(),)

            process = self._popen_factory(command, **popen_options)
            output_stream = process.stdout
            try:
                if output_stream is not None:
                    for line in output_stream:
                        captured.append(line.rstrip("\n"))
                        self._update_progress_from_line(job_id, line)
                return_code = process.wait()
            except BaseException:
                self._terminate_process(process)
                raise
            finally:
                if output_stream is not None:
                    try:
                        output_stream.close()
                    except OSError:
                        self._terminate_process(process)
                        raise
            if return_code != 0:
                raise RuntimeError(f"The calendar updater exited with code {return_code}.")

            with self._state_lock:
                if self._status.get("job_id") == job_id:
                    self._status.update(
                        {
                            "stage": "validating",
                            "message": "Checking the updated gig data...",
                        }
                    )
            snapshot = self.loader(self.project_dir, strict=True)
            if snapshot.get("errors"):
                raise RuntimeError("The updated spreadsheet data did not pass validation.")
            if not snapshot.get("files"):
                raise RuntimeError("The updater completed without a readable gig workbook.")
            self._validate_csv_exports(snapshot, manifest)

            new_count = int(snapshot.get("summary", {}).get("new", 0) or 0)
            finished = _now_local()
            persisted = {
                "last_success_at": _iso(finished),
                "last_success_display": _display_timestamp(finished),
                "new_count": max(0, new_count),
            }
            self._write_persisted_state(persisted)
            with self._state_lock:
                if self._status.get("job_id") == job_id:
                    self._status.update(
                        {
                            "status": "success",
                            "message": "Gig data refreshed successfully.",
                            "stage": "complete",
                            "finished_at": _iso(finished),
                            "last_success_at": persisted["last_success_at"],
                            "last_success_display": persisted[
                                "last_success_display"
                            ],
                            "new_count": persisted["new_count"],
                            "details": "",
                        }
                    )
        except Exception as exc:  # The public endpoint reports a safe terminal state.
            failure = exc
            rollback_errors = self._restore_outputs(rollback_dir, manifest)
            finished = _now_local()
            raw_details = "\n".join(captured)
            if raw_details:
                raw_details = f"{raw_details}\n{type(exc).__name__}: {exc}"
            else:
                raw_details = f"{type(exc).__name__}: {exc}"
            details = self._sanitize_output(raw_details)
            if rollback_errors:
                details = self._sanitize_output(
                    f"{details}\nRollback warning: {'; '.join(rollback_errors)}"
                )
                message = "Refresh failed, and the previous files could not be fully restored."
            else:
                message = "Refresh failed. Your previous gig data is still available."
            with self._state_lock:
                if self._status.get("job_id") == job_id:
                    self._status.update(
                        {
                            "status": "error",
                            "message": message,
                            "stage": "",
                            "finished_at": _iso(finished),
                            "new_count": None,
                            "details": details,
                        }
                    )
        finally:
            if failure is None:
                self._discard_rollback(rollback_dir)
            self._release_process_lock(lock_handle)

    def _update_progress_from_line(self, job_id: str, line: str) -> None:
        text = line.casefold()
        stage = "running"
        message = "Refreshing gig data..."
        if any(token in text for token in ("spreadsheet", "xlsx", "csv", "export")):
            stage, message = "spreadsheets", "Updating spreadsheets..."
        elif any(token in text for token in ("calendar", "gig", "event")):
            stage, message = "calendar", "Reading the calendar..."
        elif any(token in text for token in ("login", "sign in", "auth", "connect")):
            stage, message = "connecting", "Connecting to US Event Management..."
        with self._state_lock:
            if self._status.get("job_id") == job_id and self._status.get("status") == "running":
                self._status.update({"stage": stage, "message": message})

    def _output_paths(self) -> dict[str, Path]:
        paths: dict[str, Path] = {}
        for pattern in OUTPUT_PATTERNS:
            for path in self.project_dir.glob(pattern):
                if path.is_file():
                    paths[path.name] = path
        return paths

    def _validate_csv_exports(
        self, snapshot: dict[str, Any], manifest: dict[str, str]
    ) -> None:
        """Read every monthly CSV fully and cross-check it with its workbook."""

        current_paths = self._output_paths()
        expected_csv_names = {
            name for name in manifest if name.casefold().endswith(".csv")
        }
        missing = sorted(name for name in expected_csv_names if name not in current_paths)
        if missing:
            raise RuntimeError(
                "The updater did not produce the expected CSV export files."
            )

        workbook_counts = {
            str(metadata.get("name") or ""): int(metadata.get("row_count") or 0)
            for metadata in snapshot.get("files", [])
            if metadata.get("name")
        }
        if expected_csv_names:
            missing_pairs = sorted(
                f"{Path(workbook_name).stem}.csv"
                for workbook_name in workbook_counts
                if f"{Path(workbook_name).stem}.csv" not in current_paths
            )
            if missing_pairs:
                raise RuntimeError(
                    "The updater did not produce a CSV export for every workbook."
                )
        for name, path in sorted(current_paths.items()):
            if not name.casefold().endswith(".csv"):
                continue
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.reader(handle, strict=True)
                header = next(reader, None)
                if not header:
                    raise RuntimeError(f"CSV validation failed for {name}: missing header.")
                normalized_header = [
                    re.sub(r"[^a-z0-9]+", "", value.casefold()) for value in header
                ]
                if len(set(normalized_header)) != len(normalized_header):
                    raise RuntimeError(
                        f"CSV validation failed for {name}: duplicate columns."
                    )
                if "date" not in normalized_header or "rawcalendarlisting" not in normalized_header:
                    raise RuntimeError(
                        f"CSV validation failed for {name}: required columns are missing."
                    )
                row_count = 0
                for row in reader:
                    if len(row) != len(header):
                        raise RuntimeError(
                            f"CSV validation failed for {name}: malformed row."
                        )
                    row_count += 1

            workbook_name = f"{path.stem}.xlsx"
            if workbook_name not in workbook_counts:
                raise RuntimeError(
                    f"CSV validation failed for {name}: matching workbook is missing."
                )
            if row_count != workbook_counts[workbook_name]:
                raise RuntimeError(
                    f"CSV validation failed for {name}: row count does not match the workbook."
                )

    def _snapshot_outputs(self) -> tuple[Path, dict[str, str]]:
        self.project_dir.mkdir(parents=True, exist_ok=True)
        rollback_dir = Path(
            tempfile.mkdtemp(
                prefix=".dashboard_refresh_rollback_", dir=str(self.project_dir)
            )
        )
        manifest: dict[str, str] = {}
        try:
            for name, source in self._output_paths().items():
                backup = rollback_dir / name
                shutil.copy2(source, backup)
                manifest[name] = _sha256(backup)
            return rollback_dir, manifest
        except Exception:
            self._discard_rollback(rollback_dir)
            raise

    def _restore_outputs(
        self, rollback_dir: Path | None, manifest: dict[str, str]
    ) -> list[str]:
        if rollback_dir is None:
            return []
        errors: list[str] = []
        current = self._output_paths()
        for name, path in current.items():
            if name not in manifest:
                try:
                    path.unlink()
                except OSError as exc:
                    errors.append(f"could not remove new {name}: {type(exc).__name__}")

        for name, expected_digest in manifest.items():
            backup = rollback_dir / name
            target = self.project_dir / name
            restore_temp = self.project_dir / f".{name}.{uuid.uuid4().hex}.restore"
            try:
                shutil.copy2(backup, restore_temp)
                os.replace(restore_temp, target)
                if _sha256(target) != expected_digest:
                    raise OSError("restored content verification failed")
            except OSError as exc:
                errors.append(f"could not restore {name}: {type(exc).__name__}")
                try:
                    restore_temp.unlink(missing_ok=True)
                except OSError:
                    pass

        if not errors:
            self._discard_rollback(rollback_dir)
        return errors

    @staticmethod
    def _discard_rollback(rollback_dir: Path | None) -> None:
        if rollback_dir is None:
            return
        try:
            shutil.rmtree(rollback_dir)
        except OSError:
            pass

    def _acquire_process_lock(self):
        self.project_dir.mkdir(parents=True, exist_ok=True)
        handle = self.lock_file.open("a+", encoding="utf-8")
        if fcntl is None:  # pragma: no cover - the supported local target is macOS
            return handle
        try:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            handle.close()
            return None
        return handle

    def _another_process_holds_lock(self) -> bool:
        handle = self._acquire_process_lock()
        if handle is None:
            return True
        self._release_process_lock(handle)
        return False

    @staticmethod
    def _release_process_lock(lock_handle: Any) -> None:
        if lock_handle is None:
            return
        try:
            if fcntl is not None:
                fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)
        finally:
            lock_handle.close()

    def _read_persisted_state(self) -> dict[str, Any]:
        try:
            payload = json.loads(self.state_file.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return {}
        if not isinstance(payload, dict):
            return {}
        last_success_at = str(payload.get("last_success_at") or "")
        last_success_display = str(payload.get("last_success_display") or "")
        new_count = payload.get("new_count")
        if not isinstance(new_count, int) or isinstance(new_count, bool) or new_count < 0:
            new_count = None
        return {
            "last_success_at": last_success_at,
            "last_success_display": last_success_display,
            "new_count": new_count,
        }

    def _write_persisted_state(self, payload: dict[str, Any]) -> None:
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.state_file.with_name(
            f".{self.state_file.name}.{uuid.uuid4().hex}.tmp"
        )
        try:
            with temporary.open("w", encoding="utf-8") as handle:
                json.dump(payload, handle, indent=2, sort_keys=True)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            try:
                os.chmod(temporary, 0o600)
            except OSError:
                pass
            os.replace(temporary, self.state_file)
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass

    def _sanitize_output(self, value: str) -> str:
        text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)

        try:
            env_values = dotenv_values(self.project_dir / ".env")
        except Exception:
            env_values = {}
        literal_secrets_set = {
            str(secret)
            for secret in env_values.values()
            if secret is not None and str(secret)
        }
        literal_secrets_set.update(
            str(secret)
            for name, secret in os.environ.items()
            if _SECRET_ENV_NAME_RE.search(name) and secret
        )
        literal_secrets_set.update(self._auth_state_secrets())
        literal_secrets = sorted(
            literal_secrets_set,
            key=len,
            reverse=True,
        )
        for secret in literal_secrets:
            text = text.replace(secret, "[REDACTED]")

        text = _BEARER_RE.sub("Bearer [REDACTED]", text)
        text = _BASIC_AUTH_RE.sub("Basic [REDACTED]", text)
        text = _SECRET_LABEL_RE.sub(r"\1[REDACTED]", text)
        text = _ENV_SECRET_LINE_RE.sub(r"\1[REDACTED]", text)
        text = _QUERY_SECRET_RE.sub(r"\1[REDACTED]", text)
        text = _INLINE_SECRET_RE.sub(r"\1[REDACTED]", text)
        text = _EMAIL_RE.sub("[REDACTED EMAIL]", text)
        text = text.strip()
        if len(text) > MAX_PUBLIC_DETAILS:
            text = f"...{text[-MAX_PUBLIC_DETAILS:]}"
        return text

    def _auth_state_secrets(self) -> set[str]:
        """Collect cookie/storage values so even unlabeled debug output is safe."""

        try:
            payload = json.loads(
                (self.project_dir / "auth_state.json").read_text(encoding="utf-8")
            )
        except (OSError, ValueError, TypeError):
            return set()
        if not isinstance(payload, dict):
            return set()
        secrets: set[str] = set()
        for cookie in payload.get("cookies", []):
            if isinstance(cookie, dict) and cookie.get("value"):
                secrets.add(str(cookie["value"]))
        for origin in payload.get("origins", []):
            if not isinstance(origin, dict):
                continue
            for item in origin.get("localStorage", []):
                if isinstance(item, dict) and item.get("value"):
                    secrets.add(str(item["value"]))
        return secrets

    @staticmethod
    def _terminate_process(process: subprocess.Popen[str]) -> None:
        try:
            if process.poll() is not None:
                return
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
        except (OSError, ProcessLookupError, subprocess.SubprocessError):
            pass


__all__ = ["RefreshManager"]
