from __future__ import annotations

import json
import mimetypes
import os
import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from cgi import FieldStorage

from roster_tool.parsers import assign_sources, doctor_options, export_ics, generate_events, preview_summary


ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"


class AppHandler(BaseHTTPRequestHandler):
    server_version = "RosterCalendarHTTP/0.1"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._serve_file(WEB_ROOT / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/static/"):
            target = (WEB_ROOT / parsed.path.removeprefix("/static/")).resolve()
            if WEB_ROOT not in target.parents and target != WEB_ROOT:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if target.is_file():
                mime = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
                self._serve_file(target, mime)
                return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/analyze":
                self._handle_analyze()
                return
            if parsed.path == "/api/preview":
                self._handle_preview()
                return
            if parsed.path == "/api/export":
                self._handle_export()
                return
            self.send_error(HTTPStatus.NOT_FOUND)
        except ValueError as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except Exception:
            self._send_json({"error": "Unexpected server error."}, status=HTTPStatus.INTERNAL_SERVER_ERROR)
            raise

    def log_message(self, fmt: str, *args) -> None:
        return

    def _handle_analyze(self) -> None:
        with self._uploaded_files() as files:
            options = doctor_options(files["mmc"], files["ddh"])
            self._send_json(
                {
                    "sources": {
                        "mmc": files["mmc_name"],
                        "ddh": files["ddh_name"],
                    },
                    "doctors": [{"key": item.key, "displayName": item.display_name} for item in options],
                }
            )

    def _handle_preview(self) -> None:
        with self._uploaded_files(expect_doctor=True) as files:
            events = generate_events(files["mmc"], files["ddh"], files["doctor_key"])
            summary = preview_summary(events)
            self._send_json(
                {
                    **summary,
                    "events": [_serialize_event(event) for event in events],
                }
            )

    def _handle_export(self) -> None:
        with self._uploaded_files(expect_doctor=True) as files:
            events = generate_events(files["mmc"], files["ddh"], files["doctor_key"])
            if not events:
                raise ValueError("No calendar events were found for the selected doctor.")
            calendar_text = export_ics(events, files["doctor_display"] or files["doctor_key"])
            safe_name = (files["doctor_display"] or "roster").replace("/", "-")
            filename = f"{safe_name} roster.ics"
            payload = calendar_text.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/calendar; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.end_headers()
            self.wfile.write(payload)

    def _serve_file(self, path: Path, content_type: str) -> None:
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _uploaded_files(self, expect_doctor: bool = False):
        return UploadedFiles(self, expect_doctor=expect_doctor)


class UploadedFiles:
    def __init__(self, handler: AppHandler, expect_doctor: bool) -> None:
        self.handler = handler
        self.expect_doctor = expect_doctor
        self.tempdir: tempfile.TemporaryDirectory[str] | None = None
        self.paths: dict[str, str] = {}

    def __enter__(self) -> dict[str, str]:
        self.tempdir = tempfile.TemporaryDirectory(prefix="roster-calendar-")
        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.handler.headers.get("Content-Type", ""),
        }
        form = FieldStorage(
            fp=self.handler.rfile,
            headers=self.handler.headers,
            environ=environ,
        )
        uploads = self._save_uploads(form)
        assignments = assign_sources(uploads)
        self.paths["mmc"] = assignments["mmc"] or ""
        self.paths["ddh"] = assignments["ddh"] or ""
        self.paths["mmc_name"] = Path(assignments["mmc"]).name if assignments["mmc"] else ""
        self.paths["ddh_name"] = Path(assignments["ddh"]).name if assignments["ddh"] else ""
        self.paths["doctor_key"] = form.getfirst("doctorKey", "")
        self.paths["doctor_display"] = form.getfirst("doctorDisplay", "")
        if self.expect_doctor and not self.paths["doctor_key"]:
            raise ValueError("A doctor selection is required.")
        return self.paths

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.tempdir:
            self.tempdir.cleanup()

    def _save_uploads(self, form: FieldStorage) -> list[str]:
        raw_fields = []
        if "rosterFiles" in form:
            field = form["rosterFiles"]
            raw_fields.extend(field if isinstance(field, list) else [field])
        else:
            for legacy_name in ("mmcFile", "ddhFile"):
                if legacy_name in form:
                    raw_fields.append(form[legacy_name])
        uploads: list[str] = []
        for field in raw_fields:
            if getattr(field, "file", None) is None:
                continue
            filename = Path(field.filename or "upload.xlsx").name
            target = Path(self.tempdir.name) / filename
            with target.open("wb") as handle:
                handle.write(field.file.read())
            uploads.append(os.fspath(target))
        if not uploads:
            raise ValueError("Upload at least one roster file.")
        if len(uploads) > 2:
            raise ValueError("Upload no more than two roster files at a time.")
        return uploads


def _serialize_event(event) -> dict[str, str | bool]:
    return {
        "source": event.source,
        "title": event.title,
        "allDay": event.all_day,
        "start": event.start.isoformat(),
        "end": event.end.isoformat(),
        "location": event.location or "",
    }


def main() -> None:
    host = "127.0.0.1"
    port = 8000
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Roster to Calendar Tool running at http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server.")


if __name__ == "__main__":
    main()
