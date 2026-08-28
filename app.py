from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit

from flask import Flask, abort, jsonify, render_template, request, send_file, url_for

from dashboard_data import discover_workbooks, load_dashboard_data
from refresh_manager import RefreshManager


ROOT = Path(__file__).resolve().parent
ALLOWED_VIEWS = {
    "dashboard",
    "all",
    "new",
    "preferred",
    "removed",
    "spreadsheets",
}


def _timestamp_value(value: object) -> float:
    try:
        parsed = datetime.fromisoformat(str(value))
        return parsed.timestamp()
    except (TypeError, ValueError, OSError, OverflowError):
        return float("-inf")


def _display_timestamp(value: object) -> str:
    try:
        parsed = datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return ""
    hour = parsed.hour % 12 or 12
    period = "AM" if parsed.hour < 12 else "PM"
    return (
        f"{parsed.strftime('%B')} {parsed.day}, {parsed.year} at "
        f"{hour}:{parsed.minute:02d} {period}"
    )


def create_app(
    data_dir: str | Path | None = None,
    update_script: str | Path | None = None,
    refresh_manager: RefreshManager | None = None,
) -> Flask:
    """Create the dashboard, with injectable paths used only by local tests."""
    resolved_data_dir = Path(data_dir or ROOT).resolve()
    resolved_update_script = Path(update_script or (ROOT / "main.py")).resolve()
    app = Flask(__name__)
    app.config.update(
        DATA_DIR=resolved_data_dir,
        UPDATE_SCRIPT=resolved_update_script,
        JSON_SORT_KEYS=False,
        SEND_FILE_MAX_AGE_DEFAULT=0,
        TRUSTED_HOSTS=["127.0.0.1", "localhost", "::1"],
    )
    manager = refresh_manager or RefreshManager(
        project_dir=resolved_data_dir,
        update_script=resolved_update_script,
        loader=load_dashboard_data,
    )
    app.extensions["refresh_manager"] = manager

    def current_snapshot() -> dict:
        snapshot = load_dashboard_data(app.config["DATA_DIR"])
        refresh_status = manager.get_status()
        workbook_timestamp = snapshot.get("last_refreshed_iso", "")
        successful_timestamp = refresh_status.get("last_success_at", "")
        if successful_timestamp and _timestamp_value(
            successful_timestamp
        ) >= _timestamp_value(workbook_timestamp):
            snapshot["last_refreshed_iso"] = successful_timestamp
            snapshot["last_refreshed"] = refresh_status.get(
                "last_success_display"
            ) or _display_timestamp(successful_timestamp)
        elif workbook_timestamp:
            snapshot["last_refreshed"] = _display_timestamp(workbook_timestamp)
        for workbook in snapshot.get("files", []):
            workbook["download_url"] = url_for(
                "download_spreadsheet", filename=workbook["filename"]
            )
        return snapshot

    @app.get("/")
    def dashboard():
        snapshot = current_snapshot()
        requested_view = request.args.get("view", "dashboard").strip().lower()
        initial_view = requested_view if requested_view in ALLOWED_VIEWS else "dashboard"
        requested_month = request.args.get("month", "").strip()
        available_months = {
            month["key"] for month in snapshot.get("months", []) if month.get("key")
        }
        initial_month = requested_month if requested_month in available_months else ""
        return render_template(
            "dashboard.html",
            snapshot=snapshot,
            initial_view=initial_view,
            initial_month=initial_month,
        )

    @app.get("/api/gigs")
    def gigs_api():
        """Return the freshly loaded snapshot for lightweight local integrations."""
        return jsonify(current_snapshot())

    @app.post("/api/refresh")
    def start_refresh():
        """Start the one fixed local updater; request data cannot select a command."""
        request_host = urlsplit(f"//{request.host}").hostname or ""
        origin = request.headers.get("Origin", "").rstrip("/")
        has_browser_marker = request.headers.get("X-US-Event-Refresh") == "1"
        if (
            request_host.casefold() not in {"127.0.0.1", "localhost", "::1"}
            or (origin and origin != request.host_url.rstrip("/"))
            or (origin and not has_browser_marker)
        ):
            response = jsonify(
                {
                    "status": "error",
                    "message": "Refresh requests are accepted only from this local dashboard.",
                    "details": "",
                }
            )
            response.status_code = 403
            response.headers["Cache-Control"] = "no-store"
            return response
        accepted, payload = manager.start_refresh()
        response = jsonify(payload)
        response.status_code = (
            202 if accepted else 409 if payload.get("status") == "running" else 500
        )
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/api/refresh/status")
    def refresh_status():
        """Return only the bounded, sanitized public background-job state."""
        response = jsonify(manager.get_status())
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/spreadsheets/<path:filename>")
    def download_spreadsheet(filename: str):
        """Download only a currently discovered US_Event_*.xlsx workbook."""
        workbook = next(
            (
                path
                for path in discover_workbooks(app.config["DATA_DIR"])
                if path.name == filename
            ),
            None,
        )
        if workbook is None:
            abort(404)
        return send_file(
            workbook,
            as_attachment=True,
            download_name=workbook.name,
            conditional=True,
            max_age=0,
        )

    @app.get("/health")
    def health():
        snapshot = load_dashboard_data(app.config["DATA_DIR"])
        return {
            "status": "ok",
            "workbooks": len(snapshot.get("files", [])),
            "rows": len(snapshot.get("rows", [])),
        }

    return app


app = create_app()


if __name__ == "__main__":
    debug_mode = False

    # With Werkzeug's reloader enabled, only its child process should launch
    # the refresh. Without the reloader, WERKZEUG_RUN_MAIN is unset and this
    # process is the only server process.
    is_reloader_child = os.environ.get("WERKZEUG_RUN_MAIN") == "true"
    if not debug_mode or is_reloader_child:
        manager: RefreshManager = app.extensions["refresh_manager"]
        accepted, startup_status = manager.start_refresh()
        print(startup_status.get("message", "Starting the calendar refresh..."))

    print("US Event Dashboard running at:")
    print("http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=debug_mode)
